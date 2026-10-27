<#
.SYNOPSIS
    Verifies XSN captions, same-form version fallback, and per-file HTML matching.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path -Path $PSScriptRoot -ChildPath 'XsnTemplates.ps1')
$testRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("InfoPathXsnLabels-" + [guid]::NewGuid().ToString('N'))

function New-LabelCab {
    param([string]$Name, [string]$Version, [string]$Identity, [string]$Namespace, [string]$RequesterLabel)

    $directory = Join-Path -Path $testRoot -ChildPath $Name
    New-Item -Path $directory -ItemType Directory -Force | Out-Null
    $manifest = @"
<xsf:xDocumentClass xmlns:xsf="http://schemas.microsoft.com/office/infopath/2003/solutionDefinition" solutionVersion="$Version" name="$Identity">
  <xsf:views default="Default">
    <xsf:view name="Default"><xsf:mainpane transform="view1.xsl" /></xsf:view>
    <xsf:view name="Alternate"><xsf:mainpane transform="view2.xsl" /></xsf:view>
    <xsf:view name="Third"><xsf:mainpane transform="view3.xsl" /></xsf:view>
  </xsf:views>
</xsf:xDocumentClass>
"@
    $view = @"
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xd="http://schemas.microsoft.com/office/infopath/2003" xmlns:data="$Namespace" version="1.0">
  <xsl:template match="data:FormData">
    <label>$RequesterLabel</label><span xd:binding="data:Requestor_Name">Placeholder</span>
    <table><tr><td><b>Department:</b></td><td><span xd:binding="data:Details/data:Name">Placeholder</span></td></tr></table>
    <table><thead><tr><td>Legal name</td><td>Identifier</td></tr></thead><tbody>
      <xsl:for-each select="data:People/data:Person"><tr>
        <td><span xd:binding="data:Name">Placeholder</span></td>
        <td><span xd:binding="data:Code">Placeholder</span></td>
      </tr></xsl:for-each>
    </tbody></table>
    <span aria-label="Amount &amp; cost &lt; limit" xd:binding="data:Amount" />
    <label for="amb1">First caption</label><span id="amb1" xd:binding="data:AmbiguousField" />
    <label for="amb2">Second caption</label><span id="amb2" xd:binding="data:AmbiguousField" />
    <span xd:binding="data:Fallback_Field" />
  </xsl:template>
</xsl:stylesheet>
"@
    $alternate = @"
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xd="http://schemas.microsoft.com/office/infopath/2003" xmlns:data="$Namespace" version="1.0">
  <xsl:template match="/">
    <span aria-label="Not the default caption" xd:binding="/data:FormData/data:Requestor_Name" />
    <span aria-label="Shared alternative caption" xd:binding="/data:FormData/data:OtherViewOnly" />
    <span aria-label="Conflicting alternate caption" xd:binding="/data:FormData/data:OtherViewConflict" />
  </xsl:template>
</xsl:stylesheet>
"@
    $manifest | Set-Content -LiteralPath (Join-Path -Path $directory -ChildPath 'manifest.xsf') -Encoding UTF8
    $view | Set-Content -LiteralPath (Join-Path -Path $directory -ChildPath 'view1.xsl') -Encoding UTF8
    $alternate | Set-Content -LiteralPath (Join-Path -Path $directory -ChildPath 'view2.xsl') -Encoding UTF8
    $alternate.Replace('Conflicting alternate caption', 'Different alternate caption') |
        Set-Content -LiteralPath (Join-Path -Path $directory -ChildPath 'view3.xsl') -Encoding UTF8
    $ddf = @(
        '.OPTION EXPLICIT'
        '.Set Cabinet=ON'
        '.Set Compress=ON'
        '.Set CabinetNameTemplate=template.xsn'
        ('.Set DiskDirectoryTemplate="{0}"' -f $directory)
        ('.Set RptFileName="{0}"' -f (Join-Path -Path $directory -ChildPath 'makecab.rpt'))
        ('.Set InfFileName="{0}"' -f (Join-Path -Path $directory -ChildPath 'makecab.inf'))
        foreach ($file in @('manifest.xsf', 'view1.xsl', 'view2.xsl', 'view3.xsl')) {
            '"{0}" {1}' -f (Join-Path -Path $directory -ChildPath $file), $file
        }
    )
    $ddfPath = Join-Path -Path $directory -ChildPath 'template.ddf'
    $ddf | Set-Content -LiteralPath $ddfPath -Encoding ASCII
    $output = & "$env:SystemRoot\System32\makecab.exe" /F $ddfPath 2>&1
    if ($LASTEXITCODE -ne 0) { throw "Cannot create label fixture CAB: $($output -join ' ')" }
    return Join-Path -Path $directory -ChildPath 'template.xsn'
}

function Write-DataForm {
    param([string]$Name, [string]$Version, [string]$Identity, [string]$Namespace)

    @"
<?xml version="1.0" encoding="utf-8"?>
<?mso-infoPathSolution solutionVersion="$Version" name="$Identity" href="http://sharepoint/forms/$Identity/template.xsn"?>
<my:FormData xmlns:my="$Namespace">
  <my:Requestor_Name>Alice</my:Requestor_Name>
  <my:Details><my:Name>Operations</my:Name></my:Details>
  <my:People><my:Person><my:Name>Alice</my:Name><my:Code>1</my:Code></my:Person><my:Person><my:Name>Bob</my:Name><my:Code>2</my:Code></my:Person></my:People>
  <my:Amount>10</my:Amount><my:AmbiguousField>Value</my:AmbiguousField><my:Fallback_Field>Unlabelled</my:Fallback_Field>
  <my:OtherViewOnly>Other</my:OtherViewOnly><my:OtherViewConflict>Conflict</my:OtherViewConflict>
</my:FormData>
"@ | Set-Content -LiteralPath (Join-Path -Path $source -ChildPath ($Name + '.xml')) -Encoding UTF8
}

function Assert-XsnHtml {
    param([string]$Path, [string]$Requester, [bool]$Internal = $false, [bool]$Matched = $true)

    $html = Get-Content -LiteralPath $Path -Raw
    $labels = @([regex]::Matches($html, '<(?:div|span) class="(?:section-title|collection-title|field-label)">([^<]*)</(?:div|span)>|<th>([^<]*)</th>') |
        ForEach-Object { [System.Net.WebUtility]::HtmlDecode($_.Groups[1].Value + $_.Groups[2].Value) })
    $expected = if ($Internal) {
        @('my:Requestor_Name', 'my:Name', 'my:Amount', 'my:AmbiguousField', 'my:Fallback_Field', 'my:OtherViewOnly', 'my:OtherViewConflict')
    }
    elseif (-not $Matched) {
        @('Requestor Name', 'Name', 'Amount', 'Ambiguous Field', 'Fallback Field', 'Other View Only', 'Other View Conflict')
    }
    else {
        @($Requester, 'Department:', 'Legal name', 'Identifier', 'Amount & cost < limit', 'Ambiguous Field', 'Fallback Field', 'Shared alternative caption', 'Other View Conflict')
    }
    foreach ($label in $expected) {
        if ($labels -cnotcontains $label) { throw "Missing '$label' in '$Path'. Labels: $($labels -join '; ')" }
    }
    foreach ($unexpected in @('Not the default caption', 'First caption', 'Second caption', 'Conflicting alternate caption', 'Different alternate caption')) {
        if ($labels -ccontains $unexpected) { throw "Ambiguous or lower-priority label '$unexpected' was rendered." }
    }
    foreach ($value in @('Alice', 'Bob', 'Operations', 'Unlabelled')) {
        if (-not $html.Contains(">$value<")) { throw "Lost field value '$value' in '$Path'." }
    }
    if ($html -match '<limit') { throw 'Caption text was emitted as HTML markup.' }
}

try {
    New-Item -Path $testRoot -ItemType Directory | Out-Null
    $unprefixed = [System.Xml.XmlDocument]::new()
    $unprefixed.LoadXml('<control />')
    if ((Convert-XsnBindingPath -Expression '/Form/Field' -Node $unprefixed.DocumentElement) -ne '/:Form/:Field') {
        throw 'Unprefixed XML field bindings were not resolved.'
    }
    $renamedDocument = [System.Xml.XmlDocument]::new()
    $renamedDocument.LoadXml('<?mso-infoPathSolution name="urn:form:Applications" solutionVersion="1.0.0.3416" href="http://sharepoint/forms/template.xsn"?><Form />')
    $linkedRecord = [pscustomobject]@{
        Version = [version]'1.0.0.3443'
        SharePointVersion = [version]'1.0'
        Row = [pscustomobject]@{
            SolutionVersion = '1.0.0.3443'; SolutionName = 'urn:form:template'
            TemplateUrl = 'http://sharepoint/forms/template.xsn'; Modified = '2026-10-27'
        }
    }
    $renamedMatch = Find-XsnTemplate -Document $renamedDocument -Inventory @($linkedRecord)
    if ($renamedMatch.Status -ne 'NextVersion' -or $renamedMatch.Template.Version -ne [version]'1.0.0.3443' -or
        $renamedMatch.MatchBasis -ne 'TemplateUrlIdentityFallback' -or $renamedMatch.LinkedTemplateVersions -ne '1.0.0.3443' -or
        $renamedMatch.LinkedTemplateMessage -notlike '*1.0.0.3443*1.0.0.3416*' -or
        $renamedMatch.Message -notlike "*XML form identity 'urn:form:Applications' differs*") {
        throw 'Renamed-form exact-URL fallback or linked-version diagnostics failed.'
    }
    $renamedInstruction = $renamedDocument.SelectSingleNode("/processing-instruction('mso-infoPathSolution')")
    $originalInstruction = $renamedInstruction.Data
    $renamedInstruction.Data = $originalInstruction.Replace('1.0.0.3416', '1.0.0.3443')
    if ((Find-XsnTemplate -Document $renamedDocument -Inventory @($linkedRecord)).Status -ne 'ExactVersion') {
        throw 'An exact-version renamed form did not match its exact template URL.'
    }
    $renamedInstruction.Data = $originalInstruction.Replace('http://sharepoint/forms/template.xsn', '/forms/template.xsn')
    if ((Find-XsnTemplate -Document $renamedDocument -Inventory @($linkedRecord)).MatchBasis -ne 'TemplateUrlIdentityFallback') {
        throw 'Server-relative exact-URL fallback failed for a renamed form.'
    }
    $renamedInstruction.Data = $originalInstruction.Replace('http://sharepoint/forms/template.xsn', 'http://sharepoint/unrelated/template.xsn')
    if ((Find-XsnTemplate -Document $renamedDocument -Inventory @($linkedRecord)).Status -ne 'NoMatch') {
        throw 'A mismatched form name incorrectly fell back using only a template filename.'
    }
    $renamedInstruction.Data = $originalInstruction
    $conflictingRecord = [pscustomobject]@{
        Version = [version]'1.0.0.3450'
        SharePointVersion = [version]'1.0'
        Row = [pscustomobject]@{
            SolutionVersion = '1.0.0.3450'; SolutionName = 'urn:form:unrelated'
            TemplateUrl = 'http://sharepoint/forms/template.xsn'; Modified = '2026-10-27'
        }
    }
    $ambiguousMatch = Find-XsnTemplate -Document $renamedDocument -Inventory @($linkedRecord, $conflictingRecord)
    if ($ambiguousMatch.Status -ne 'NoMatch' -or $ambiguousMatch.Message -notlike '*ambiguous*' -or
        $ambiguousMatch.LinkedTemplateVersions -ne '1.0.0.3443, 1.0.0.3450') {
        throw 'Ambiguous renamed-form candidates at one template URL were accepted or not diagnosed.'
    }
    $conflictingRecord.Row.SolutionName = 'urn:form:Applications'
    $preferredMatch = Find-XsnTemplate -Document $renamedDocument -Inventory @($linkedRecord, $conflictingRecord)
    if ($preferredMatch.MatchBasis -ne 'SolutionName' -or $preferredMatch.Template.Version -ne [version]'1.0.0.3450') {
        throw 'An eligible explicit form identity did not take precedence over URL fallback.'
    }
    $linkedRecord.Version = [version]'1.0.0.3410'
    $linkedRecord.Row.SolutionVersion = '1.0.0.3410'
    if ((Find-XsnTemplate -Document $renamedDocument -Inventory @($linkedRecord)).Status -ne 'NoMatch') {
        throw 'Renamed-form URL fallback selected an older template.'
    }
    $exportRoot = Join-Path -Path $testRoot -ChildPath 'Export'
    New-Item -Path $exportRoot -ItemType Directory | Out-Null
    $downloads = @(
        foreach ($spec in @(
            @{ Name = 'A3'; Version = '1.0.0.3'; Identity = 'urn:form:A'; Namespace = 'urn:fields:A'; Label = 'Who is requesting?' }
            @{ Name = 'A10'; Version = '1.0.0.10'; Identity = 'urn:form:A'; Namespace = 'urn:fields:A'; Label = 'Version ten requester' }
            @{ Name = 'B3'; Version = '1.0.0.3'; Identity = 'urn:form:B'; Namespace = 'urn:fields:B'; Label = 'Other form requester' }
        )) {
            $cab = New-LabelCab -Name $spec.Name -Version $spec.Version -Identity $spec.Identity -Namespace $spec.Namespace -RequesterLabel $spec.Label
            $relative = $spec.Name + '.xsn'
            Copy-Item -LiteralPath $cab -Destination (Join-Path -Path $exportRoot -ChildPath $relative)
            [pscustomobject]@{
                SiteUrl = 'http://sharepoint'
                LibraryName = 'Forms'
                TemplateUrl = "http://sharepoint/forms/$($spec.Identity)/template.xsn"
                ConfiguredBy = 'Library'
                SharePointVersion = '1.0'
                IsCurrentVersion = $true
                Modified = '2026-10-26T12:00:00.0000000'
                XsnRelativePath = $relative
                Status = 'Success'
                Error = ''
            }
        }
    )
    $exportCsv = Join-Path -Path $exportRoot -ChildPath 'XsnTemplateDownloads.csv'
    $downloads | Export-Csv -LiteralPath $exportCsv -NoTypeInformation -Encoding UTF8
    $source = Join-Path -Path $testRoot -ChildPath 'Source'
    New-Item -Path $source -ItemType Directory | Out-Null
    Write-DataForm -Name 'AExact' -Version '1.0.0.3' -Identity 'urn:form:A' -Namespace 'urn:fields:A'
    Write-DataForm -Name 'ANext' -Version '1.0.0.2' -Identity 'urn:form:A' -Namespace 'urn:fields:A'
    Write-DataForm -Name 'AHigher' -Version '1.0.0.4' -Identity 'urn:form:A' -Namespace 'urn:fields:A'
    Write-DataForm -Name 'ATooNew' -Version '1.0.0.11' -Identity 'urn:form:A' -Namespace 'urn:fields:A'
    Write-DataForm -Name 'BExact' -Version '1.0.0.3' -Identity 'urn:form:B' -Namespace 'urn:fields:B'
    Write-DataForm -Name 'CUnknown' -Version '1.0.0.3' -Identity 'urn:form:C' -Namespace 'urn:fields:C'
    $originalHashes = @(Get-ChildItem -LiteralPath $source -File | Get-FileHash | Select-Object Path, Hash)
    $pipelineOutput = Join-Path -Path $testRoot -ChildPath 'Output'
    & (Join-Path -Path $repoRoot -ChildPath 'RunInfoPathPipeline.ps1') `
        -SourceDirectory $source -OutputRootDirectory $pipelineOutput -XsnExportCsvPath $exportCsv -StartStep 1 6>$null

    $inventoryPath = Join-Path -Path $pipelineOutput -ChildPath 'XSNTemplates\XsnTemplateVersions.csv'
    $inventoryRows = @(Import-Csv -LiteralPath $inventoryPath)
    if ($inventoryRows.Count -ne 3 -or @($inventoryRows | Where-Object { $_.Status -ne 'Success' -or [int]$_.LabelCount -ne 6 -or [int]$_.ViewCount -ne 3 }).Count -ne 0) {
        throw "Expected three XSNs with six resolved labels each. Actual label counts: $($inventoryRows.LabelCount -join ',')"
    }
    $records = Import-XsnInventory -Path $inventoryPath
    $a10 = @($records | Where-Object { $_.Row.SolutionName -ceq 'urn:form:A' -and $_.Version -eq [version]'1.0.0.10' })[0]
    $a10.Row.TemplateUrl = 'http://sharepoint/forms/NewLocation/template.xsn'
    $higherDocument = Read-XsnXml -Path (Join-Path -Path $source -ChildPath 'AHigher.xml')
    $relocatedMatch = Find-XsnTemplate -Document $higherDocument -Inventory $records
    if ($relocatedMatch.Status -ne 'NextVersion' -or $relocatedMatch.Template.Row.TemplateUrl -ne $a10.Row.TemplateUrl) {
        throw 'Same-form next-version matching failed after a template URL changed.'
    }
    $hrefDocument = Read-XsnXml -Path (Join-Path -Path $source -ChildPath 'ANext.xml')
    $instruction = $hrefDocument.SelectSingleNode("/processing-instruction('mso-infoPathSolution')")
    $instruction.Data = $instruction.Data.Replace(' name="urn:form:A"', '').Replace('http://sharepoint/forms/urn:form:A/template.xsn', '/forms/urn:form:A/template.xsn')
    $hrefMatch = Find-XsnTemplate -Document $hrefDocument -Inventory $records
    if ($hrefMatch.Status -ne 'NextVersion' -or $hrefMatch.Template.Row.SolutionName -cne 'urn:form:A') {
        throw 'Server-relative href did not isolate the correct form when name was absent.'
    }
    $instruction.Data = $instruction.Data.Replace('/forms/urn:form:A/template.xsn', 'template.xsn')
    if ((Find-XsnTemplate -Document $hrefDocument -Inventory $records).Status -ne 'NoMatch') {
        throw 'An ambiguous template filename matched a different form.'
    }
    $htmlRoot = Join-Path -Path $pipelineOutput -ChildPath 'Html'
    Assert-XsnHtml -Path (Join-Path -Path $htmlRoot -ChildPath 'AExact.html') -Requester 'Who is requesting?'
    Assert-XsnHtml -Path (Join-Path -Path $htmlRoot -ChildPath 'ANext.html') -Requester 'Who is requesting?'
    Assert-XsnHtml -Path (Join-Path -Path $htmlRoot -ChildPath 'AHigher.html') -Requester 'Version ten requester'
    Assert-XsnHtml -Path (Join-Path -Path $htmlRoot -ChildPath 'ATooNew.html') -Requester '' -Matched $false
    Assert-XsnHtml -Path (Join-Path -Path $htmlRoot -ChildPath 'BExact.html') -Requester 'Other form requester'
    Assert-XsnHtml -Path (Join-Path -Path $htmlRoot -ChildPath 'CUnknown.html') -Requester '' -Matched $false
    $matches = @(Import-Csv -LiteralPath (Join-Path -Path $htmlRoot -ChildPath 'InfoPathXsnMatches.csv'))
    foreach ($case in @(
        @{ File = 'AExact.xml'; Status = 'ExactVersion'; Version = '1.0.0.3' }
        @{ File = 'ANext.xml'; Status = 'NextVersion'; Version = '1.0.0.3' }
        @{ File = 'AHigher.xml'; Status = 'NextVersion'; Version = '1.0.0.10' }
        @{ File = 'ATooNew.xml'; Status = 'NoMatch'; Version = '' }
        @{ File = 'BExact.xml'; Status = 'ExactVersion'; Version = '1.0.0.3' }
        @{ File = 'CUnknown.xml'; Status = 'NoMatch'; Version = '' }
    )) {
        $row = @($matches | Where-Object { [System.IO.Path]::GetFileName($_.SourceFilePath) -eq $case.File })[0]
        if ($row.MatchStatus -ne $case.Status -or $row.MatchedVersion -ne $case.Version -or $row.Error) {
            throw "Incorrect per-file XSN match for '$($case.File)'."
        }
    }
    $mapping = @(Import-Csv -LiteralPath (Join-Path -Path $pipelineOutput -ChildPath 'Templates\InfoPathXsltGroupMap.csv'))
    $sampleMap = @($mapping | Where-Object { [System.IO.Path]::GetFileName($_.SampleSourceFilePath) -eq 'AExact.xml' })[0]
    if ($sampleMap.XsnMatchStatus -ne 'ExactVersion' -or $sampleMap.XsnSolutionVersion -ne '1.0.0.3') {
        throw "Generated-template XSN tracking is incorrect: status=$($sampleMap.XsnMatchStatus), version=$($sampleMap.XsnSolutionVersion), sample=$($sampleMap.SampleSourceFilePath)."
    }
    $direct = Join-Path -Path $testRoot -ChildPath 'Direct.html'
    $transform = [System.Xml.Xsl.XslCompiledTransform]::new()
    $transform.Load($sampleMap.XsltPath)
    $transform.Transform($sampleMap.SampleSourceFilePath, $direct)
    Assert-XsnHtml -Path $direct -Requester 'Who is requesting?'

    & (Join-Path $repoRoot 'ProcessInfopathToHTML.ps1') -InfoPathDirectory $source `
        -MetadataCsvPath (Join-Path $pipelineOutput 'Reports\InfoPathVersionMap.csv') `
        -GroupingCsvPath (Join-Path $pipelineOutput 'Reports\InfoPathGroupingDetailed.csv') `
        -TemplateMappingCsvPath (Join-Path $pipelineOutput 'Templates\InfoPathXsltGroupMap.csv') `
        -OutputDirectory $htmlRoot -XsnInventoryCsvPath $inventoryPath -UseDisplayName $false 6>$null
    Assert-XsnHtml -Path (Join-Path -Path $htmlRoot -ChildPath 'AExact.html') -Requester '' -Internal $true
    $internalMatches = @(Import-Csv -LiteralPath (Join-Path -Path $htmlRoot -ChildPath 'InfoPathXsnMatches.csv'))
    if (@($internalMatches | Where-Object { [int]$_.AppliedLabelCount -ne 0 }).Count -gt 0) { throw 'Internal-name mode applied display labels.' }
    $currentHashes = @(Get-ChildItem -LiteralPath $source -File | Get-FileHash | Select-Object Path, Hash)
    if (Compare-Object $originalHashes $currentHashes -Property Path, Hash) { throw 'XSN label processing modified source XML.' }
    Write-Host 'PASS: XSN captions, namespaces, repeating headers, view precedence, ambiguity, numeric next-version matching, same-form isolation, tracking, escaping, internal mode, and resume.'
}
finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
