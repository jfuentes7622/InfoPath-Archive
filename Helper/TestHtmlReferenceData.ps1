<#
.SYNOPSIS
    Tests offline list/user resolution in generated and XSN archival HTML.
#>
[CmdletBinding()]
param([string]$ExtractedSampleDirectory, [switch]$PrintToPdf)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'HtmlRendering.ps1')
. (Join-Path $PSScriptRoot 'XsnTemplates.ps1')
. (Join-Path $PSScriptRoot 'XsnHtmlRendering.ps1')
. (Join-Path $PSScriptRoot 'HtmlReferenceData.ps1')
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('InfoPathReferences-' + [guid]::NewGuid().ToString('N'))

function Assert-Reference {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Assert-ReferencePdf {
    param([string]$Path)
    if (-not $PrintToPdf) { return }
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    Assert-Reference ($bytes.Length -gt 1000 -and [System.Text.Encoding]::ASCII.GetString($bytes, 0, 5) -eq '%PDF-') 'Reference-enabled PDF output is invalid.'
}

try {
    $source = Join-Path $testRoot 'Source'
    $bundle = Join-Path $testRoot 'ReferenceData'
    $resources = Join-Path $testRoot 'Resources'
    New-Item $source, $bundle, $resources -ItemType Directory -Force | Out-Null
    $site = 'http://sharepoint/sites/Finance'
    $listId = '11111111-1111-1111-1111-111111111111'
    $otherId = '22222222-2222-2222-2222-222222222222'
    @(
        [pscustomobject]@{ ID = '1'; Value = 'US'; Title = 'United States & Territories' }
        [pscustomobject]@{ ID = '2'; Value = 'GB'; Title = 'Changed current label' }
        [pscustomobject]@{ ID = '3'; Value = 'dup'; Title = 'First label' }
        [pscustomobject]@{ ID = '4'; Value = 'dup'; Title = 'Conflicting label' }
        [pscustomobject]@{ ID = '5'; Value = 'Y'; Title = 'Affirmative' }
    ) | Export-Csv (Join-Path $bundle 'Countries.csv') -NoTypeInformation -Encoding UTF8
    [pscustomobject]@{ ID = '1'; Value = 'US'; Title = 'Different list label' } |
        Export-Csv (Join-Path $bundle 'Other.csv') -NoTypeInformation -Encoding UTF8
    @(
        [pscustomobject]@{ UserId = '42'; LoginName = 'i:0#.w|DOMAIN\archivist'; DisplayName = 'Archive User'; Email = '' }
        [pscustomobject]@{ UserId = '43'; LoginName = 'i:0#.f|membership|person@example.test'; DisplayName = 'Membership User'; Email = '' }
    ) | Export-Csv (Join-Path $bundle 'Users.csv') -NoTypeInformation -Encoding UTF8
    $bindings = @(
        foreach ($field in @('Country', 'CachedCountry', 'UnknownCountry', 'DuplicateCountry', 'OverrideCountry', 'Radio')) {
            [pscustomobject]@{
                XPath = '/my:Form/my:' + $field; ConnectionName = 'Countries'; SiteUrl = $site; ListId = $listId
                ValueField = 'Value'; DisplayField = 'Title'; CacheRowXPath = ''; ValueXPath = 'd:Value'; DisplayXPath = 'd:Title'
            }
        }
        [pscustomobject]@{
            XPath = '/my:Form/my:Row/my:Country'; ConnectionName = 'Countries'; SiteUrl = $site; ListId = $listId
            ValueField = 'Value'; DisplayField = 'Title'; CacheRowXPath = ''; ValueXPath = 'd:Value'; DisplayXPath = 'd:Title'
        }
        [pscustomobject]@{
            XPath = '/my:Form/my:OtherCountry'; ConnectionName = 'Other'; SiteUrl = $site; ListId = $otherId
            ValueField = 'Value'; DisplayField = 'Title'; CacheRowXPath = ''; ValueXPath = 'd:Value'; DisplayXPath = 'd:Title'
        }
    )
    $template = [pscustomobject]@{
        SolutionName = 'urn:reference:test'; SolutionVersion = '1.0.0.2'; TemplateUrl = "$site/Applications/Forms/template.xsn"
        Sha256 = 'sample'; UserSiteUrl = $site
        Namespaces = [pscustomobject]@{ my = 'urn:reference:fields'; pc = 'urn:reference:person'; d = 'urn:reference:secondary' }
        Bindings = $bindings
    }
    $manifest = [pscustomobject]@{
        SchemaVersion = 1; ExportedAtUtc = '2026-10-27T00:00:00Z'
        Dependencies = [pscustomobject]@{ SchemaVersion = 1; Templates = @($template); Lists = @(); UserSites = @($site); Warnings = @() }
        Lists = @(
            [pscustomobject]@{ SiteUrl = $site; ListId = $listId; Status = 'Success'; Error = ''; DataPath = 'Countries.csv'; RowCount = 5; SiteCollectionUrl = $site }
            [pscustomobject]@{ SiteUrl = $site; ListId = $otherId; Status = 'Success'; Error = ''; DataPath = 'Other.csv'; RowCount = 1; SiteCollectionUrl = $site }
        )
        Users = @([pscustomobject]@{
            SiteCollectionUrl = $site; SourceSiteUrls = @($site); Status = 'Success'; Error = ''; DataPath = 'Users.csv'; RowCount = 2
        })
    }
    $manifestPath = Join-Path $bundle 'Manifest.json'
    $manifest | ConvertTo-Json -Depth 30 | Set-Content $manifestPath -Encoding UTF8
    $xmlPath = Join-Path $source 'Sample.xml'
    @'
<?mso-infoPathSolution solutionVersion="1.0.0.1" name="urn:reference:test" href="http://sharepoint/sites/Finance/Applications/Forms/template.xsn"?>
<my:Form xmlns:my="urn:reference:fields" xmlns:pc="urn:reference:person" xmlns:d="urn:reference:secondary">
 <my:Country>US</my:Country><my:CachedCountry>GB</my:CachedCountry><my:OtherCountry>US</my:OtherCountry>
 <my:UnknownCountry>ZZ</my:UnknownCountry><my:DuplicateCountry>dup</my:DuplicateCountry><my:OverrideCountry>US</my:OverrideCountry>
 <my:SavedLookup>1;#Saved lookup label</my:SavedLookup><my:Radio>Y</my:Radio>
 <my:Owner><pc:Person><pc:DisplayName/><pc:AccountId>DOMAIN\archivist</pc:AccountId></pc:Person></my:Owner>
 <my:SavedOwner><pc:Person><pc:DisplayName>Historical Person</pc:DisplayName><pc:AccountId>DOMAIN\archivist</pc:AccountId></pc:Person></my:SavedOwner>
 <my:NumericOwner><pc:Person><pc:DisplayName/><pc:AccountId>42</pc:AccountId></pc:Person></my:NumericOwner>
 <my:AutoOwner><pc:Person><pc:DisplayName/><pc:AccountId>DOMAIN\archivist</pc:AccountId></pc:Person></my:AutoOwner>
 <my:UnknownOwner><pc:Person><pc:AccountId>unknown\account</pc:AccountId></pc:Person></my:UnknownOwner>
 <my:Row><my:Country>US</my:Country></my:Row><my:Row><my:Country>GB</my:Country></my:Row>
 <my:DataConnection Name="Countries"><d:Row><d:Value>GB</d:Value><d:Title>Historical country label</d:Title></d:Row></my:DataConnection>
</my:Form>
'@ | Set-Content $xmlPath -Encoding UTF8
    $originalHash = (Get-FileHash $xmlPath).Hash
    $rulesPath = Join-Path $testRoot 'Rules.json'
    @'
{"Namespaces":{"my":"urn:reference:fields"},"Fields":[
 {"XPath":"/my:Form/my:Owner","Type":"Person"},
 {"XPath":"/my:Form/my:SavedOwner","Type":"Person"},
 {"XPath":"/my:Form/my:NumericOwner","Type":"Person"},
 {"XPath":"/my:Form/my:UnknownOwner","Type":"Person"},
 {"XPath":"/my:Form/my:OverrideCountry","Type":"Dropdown","Values":{"US":"Explicit override"}}
]}
'@ | Set-Content $rulesPath -Encoding UTF8
    $data = Import-HtmlReferenceData $bundle
    $configuration = Import-HtmlFieldRules $rulesPath
    $document = Read-XsnXml $xmlPath
    $audit = [System.Collections.Generic.List[object]]::new()
    Set-HtmlReferenceValues $document $data $audit $xmlPath
    Assert-Reference ($document.DocumentElement.SelectSingleNode('*[local-name()="Country"]').InnerText -eq 'US') 'Resolution changed raw field values.'
    Convert-HtmlFieldValues $document $configuration
    Assert-Reference ($document.DocumentElement.SelectSingleNode('*[local-name()="Country"]').InnerText -eq 'United States & Territories') 'Value-to-Title resolution failed.'
    Assert-Reference ($document.DocumentElement.SelectSingleNode('*[local-name()="CachedCountry"]').InnerText -eq 'Historical country label') 'Saved cache did not take precedence.'
    Assert-Reference ($document.DocumentElement.SelectSingleNode('*[local-name()="Owner"]').InnerText -eq 'Archive User') 'Windows claims alias did not resolve.'
    Assert-Reference ($document.DocumentElement.SelectSingleNode('*[local-name()="SavedOwner"]').InnerText -eq 'Historical Person') 'Saved person name was overwritten.'
    Assert-Reference ($document.DocumentElement.SelectSingleNode('*[local-name()="NumericOwner"]').InnerText -eq 'Archive User') 'Scoped user ID did not resolve.'
    Assert-Reference ($document.DocumentElement.SelectSingleNode('*[local-name()="AutoOwner"]').InnerText -eq 'Archive User') 'Automatic person rendering duplicated names/account IDs.'
    Assert-Reference ($document.DocumentElement.SelectSingleNode('*[local-name()="OverrideCountry"]').InnerText -eq 'Explicit override') 'Explicit field rule did not override snapshot.'
    Assert-Reference (@($audit | Where-Object Status -eq 'Ambiguous').Count -eq 1) 'Conflicting lookup keys were not audited.'
    Assert-Reference (@($audit | Where-Object Status -eq 'Unresolved').Count -eq 2) 'Unknown IDs/accounts were not audited.'
    Assert-Reference (@($audit | Where-Object Source -eq 'SavedCache').Count -eq 2) 'Cached lookup provenance was not preserved.'

    @'
<xsf:xDocumentClass xmlns:xsf="http://schemas.microsoft.com/office/infopath/2003/solutionDefinition" solutionVersion="1.0.0.2" name="urn:reference:test">
 <xsf:views default="Archive"><xsf:view name="Archive"><xsf:mainpane transform="view1.xsl"/></xsf:view></xsf:views>
</xsf:xDocumentClass>
'@ | Set-Content (Join-Path $resources 'manifest.xsf') -Encoding UTF8
    @'
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:my="urn:reference:fields" xmlns:xd="http://schemas.microsoft.com/office/infopath/2003" version="1.0">
 <xsl:output method="html"/><xsl:template match="/"><html><body><h1>Reference archive</h1><xsl:for-each select="my:Form">
 <span xd:binding="my:Country"/><span xd:binding="my:CachedCountry"/><span xd:binding="my:OtherCountry"/>
 <span xd:binding="my:UnknownCountry"/><span xd:binding="my:DuplicateCountry"/><span xd:binding="my:OverrideCountry"/>
 <span xd:binding="my:Owner"/><span xd:binding="my:SavedOwner"/><span xd:binding="my:NumericOwner"/><span xd:binding="my:AutoOwner"/>
 <input type="radio" value="Y" xd:binding="my:Radio"><xsl:if test="my:Radio='Y'"><xsl:attribute name="checked">checked</xsl:attribute></xsl:if></input>
 <xsl:for-each select="my:Row"><span xd:binding="my:Country"/></xsl:for-each>
 </xsl:for-each></body></html></xsl:template>
</xsl:stylesheet>
'@ | Set-Content (Join-Path $resources 'view1.xsl') -Encoding UTF8
    $output = Join-Path $testRoot 'Pipeline'
    $pdfSetting = if ($PrintToPdf) { 'true' } else { 'false' }
    & (Join-Path $repoRoot 'RunInfoPathPipeline.ps1') -SourceDirectory $source -OutputRootDirectory $output -StartStep 1 -FieldRulesPath $rulesPath -ReferenceDataPath $bundle -PrintToPdf $pdfSetting 6>$null
    $htmlPath = Join-Path $output 'Html\Sample.html'
    $generated = Get-Content $htmlPath -Raw
    foreach ($value in @('United States &amp; Territories', 'Historical country label', 'Different list label', 'Archive User', 'Historical Person', 'Explicit override', 'ZZ', 'dup', 'Saved lookup label')) {
        Assert-Reference ($generated.Contains($value)) "Generated HTML missing '$value'."
    }
    $pipelineAudit = @(Import-Csv (Join-Path $output 'Html\InfoPathReferenceResolutions.csv'))
    Assert-Reference ($pipelineAudit.Count -eq $audit.Count) 'Pipeline reference audit differs from direct resolution.'
    Assert-ReferencePdf (Join-Path $output 'Html\Sample.pdf')
    $csv = @(Get-ChildItem (Join-Path $output 'Reports\Normalized') -Filter Main.csv -Recurse)[0]
    $normalizedBefore = (Get-FileHash $csv.FullName).Hash
    Assert-Reference (-not (Get-Content $csv.FullName -Raw).Contains('United States & Territories')) 'Snapshot labels leaked into CSV.'

    $xsnRoot = Join-Path $output 'XSNTemplates'
    New-Item $xsnRoot -ItemType Directory | Out-Null
    Copy-Item $resources (Join-Path $xsnRoot 'Resources') -Recurse
    '{"Labels":[]}' | Set-Content (Join-Path $xsnRoot 'Labels.json') -Encoding UTF8
    [pscustomobject]@{
        TemplateUrl = $template.TemplateUrl; SolutionVersion = $template.SolutionVersion; SolutionName = $template.SolutionName
        LabelsRelativePath = 'Labels.json'; SharePointVersion = '1.0'; Modified = ''; Status = 'Success'
        ResourcesRelativePath = 'Resources'; Sha256 = 'sample'; XsnPath = ''
    } | Export-Csv (Join-Path $xsnRoot 'XsnTemplateVersions.csv') -NoTypeInformation -Encoding UTF8
    & (Join-Path $repoRoot 'RunInfoPathPipeline.ps1') -SourceDirectory $source -OutputRootDirectory $output -StartStep 7 -FieldRulesPath $rulesPath -ReferenceDataPath $manifestPath -UseXSN4HTML $true -PrintToPdf $pdfSetting 6>$null
    $xsnHtml = Get-Content $htmlPath -Raw
    foreach ($value in @('Reference archive', 'United States &amp; Territories', 'Historical country label', 'Different list label', 'Archive User', 'Historical Person', 'Explicit override', '(X)', 'ZZ', 'dup')) {
        Assert-Reference ($xsnHtml.Contains($value)) "XSN HTML missing '$value'."
    }
    Assert-Reference ((Import-Csv (Join-Path $output 'Html\InfoPathXsnMatches.csv')).ActualRenderer -eq 'XSN') 'Reference-enabled XSN rendering fell back.'
    Assert-ReferencePdf (Join-Path $output 'Html\Sample.pdf')
    Assert-Reference ((Get-FileHash $csv.FullName).Hash -eq $normalizedBefore) 'Step 7 changed normalized CSV.'
    Assert-Reference ((Get-FileHash $xmlPath).Hash -eq $originalHash) 'Source XML changed.'

    $liveSource = Join-Path $testRoot 'LiveSource'
    $liveOutput = Join-Path $testRoot 'LiveOutput'
    New-Item $liveSource -ItemType Directory | Out-Null
    Copy-Item $xmlPath (Join-Path $liveSource 'First.xml')
    Copy-Item $xmlPath (Join-Path $liveSource 'Second.xml')
    $liveAuditPath = Join-Path $liveOutput 'Html\InfoPathReferenceResolutions.csv'
    $liveAuditState = [pscustomobject]@{ Observed = $false; Existed = $false; Rows = @() }
    & (Join-Path $repoRoot 'RunInfoPathPipeline.ps1') -SourceDirectory $liveSource -OutputRootDirectory $liveOutput `
        -StartStep 1 -FieldRulesPath $rulesPath -ReferenceDataPath $bundle 6>&1 | ForEach-Object {
            if (-not $liveAuditState.Observed -and $_.ToString() -like '*unresolved/ambiguous reference(s)*') {
                $liveAuditState.Existed = Test-Path -LiteralPath $liveAuditPath -PathType Leaf
                if ($liveAuditState.Existed) { $liveAuditState.Rows = @(Import-Csv -LiteralPath $liveAuditPath) }
                $liveAuditState.Observed = $true
            }
        }
    Assert-Reference ($liveAuditState.Observed -and $liveAuditState.Existed) 'Reference audit was not created before its warning was printed.'
    Assert-Reference ($liveAuditState.Rows.Count -eq $audit.Count) 'Current file reference rows were not flushed before its warning.'
    Assert-Reference (@($liveAuditState.Rows.SourceFilePath | Select-Object -Unique).Count -eq 1) 'Live audit contains unexpected files.'
    Assert-Reference (@(Import-Csv -LiteralPath $liveAuditPath).Count -eq ($audit.Count * 2)) 'Incremental audit lost records from either file.'
    & (Join-Path $repoRoot 'RunInfoPathPipeline.ps1') -SourceDirectory $liveSource -OutputRootDirectory $liveOutput `
        -StartStep 7 -FieldRulesPath $rulesPath -ReferenceDataPath $bundle 6>$null
    $resumedRows = @(Import-Csv -LiteralPath $liveAuditPath)
    Assert-Reference ($resumedRows.Count -eq ($audit.Count * 2)) 'Rerun duplicated or lost reference rows.'
    Assert-Reference (@($resumedRows.SourceFilePath | Select-Object -Unique).Count -eq 2) 'Rerun did not retain both files in the audit.'
    Assert-Reference (@(Get-Content -LiteralPath $liveAuditPath | Where-Object { $_ -like '"SourceFilePath",*' }).Count -eq 1) 'Incremental audit contains duplicate headers.'
    & (Join-Path $repoRoot 'RunInfoPathPipeline.ps1') -SourceDirectory $liveSource -OutputRootDirectory $liveOutput `
        -StartStep 7 -FieldRulesPath $rulesPath 6>$null
    Assert-Reference (@(Import-Csv -LiteralPath $liveAuditPath).Count -eq 0) 'Run without reference data retained stale audit rows.'

    $scopedData = Import-HtmlReferenceData $bundle
    $scopedData.Users[0].Entry.SourceSiteUrls = @('http://sharepoint/sites/OtherCollection')
    $scopedAudit = [System.Collections.Generic.List[object]]::new()
    $scopedDocument = Read-XsnXml $xmlPath
    Set-HtmlReferenceValues $scopedDocument $scopedData $scopedAudit $xmlPath
    Assert-Reference (@($scopedAudit | Where-Object { $_.StoredValue -eq '42' -and $_.Status -eq 'MissingExport' }).Count -eq 1) 'User IDs leaked across site collections.'
    $selectedXsn = [pscustomobject]@{
        Template = [pscustomobject]@{
            Version = [version]'1.0.0.2'
            Row = [pscustomobject]@{
                TemplateUrl = $template.TemplateUrl; SolutionVersion = '1.0.0.2'; SolutionName = $template.SolutionName; Sha256 = 'different-package'
            }
        }
    }
    $hashAudit = [System.Collections.Generic.List[object]]::new()
    Set-HtmlReferenceValues (Read-XsnXml $xmlPath) $data $hashAudit $xmlPath -XsnMatch $selectedXsn
    Assert-Reference ($hashAudit.Count -eq 1 -and $hashAudit[0].Status -eq 'NoTemplate') 'Mismatched XSN reference definition was accepted.'
    $claimsIndex = Get-HtmlReferenceIndex $data.Users[0] 'LoginName' 'DisplayName' -IgnoreCase
    Assert-Reference (-not $claimsIndex.ContainsKey('person@example.test')) 'Membership claims identity was reduced to an unscoped email.'
    $legacyData = Import-HtmlReferenceData $bundle
    $legacyData.Users[0].Rows[0].LoginName = 'DOMAIN\archivist'
    $claimsDocument = Read-XsnXml $xmlPath
    $claimsDocument.SelectSingleNode("//*[local-name()='Owner']//*[local-name()='AccountId']").InnerText = 'i:0#.w|DOMAIN\archivist'
    $claimsAudit = [System.Collections.Generic.List[object]]::new()
    Set-HtmlReferenceValues $claimsDocument $legacyData $claimsAudit $xmlPath
    Assert-Reference (@($claimsAudit | Where-Object { $_.StoredValue -eq 'i:0#.w|DOMAIN\archivist' -and $_.DisplayValue -eq 'Archive User' }).Count -eq 1) 'Windows claims-to-legacy login resolution failed.'

    $manifest.Lists[0].Status = 'Failed'
    $manifest.Lists[0].Error = 'Access denied'
    $manifest | ConvertTo-Json -Depth 30 | Set-Content $manifestPath -Encoding UTF8
    $partial = Import-HtmlReferenceData $bundle
    $partialDocument = Read-XsnXml $xmlPath
    $partialAudit = [System.Collections.Generic.List[object]]::new()
    Set-HtmlReferenceValues $partialDocument $partial $partialAudit $xmlPath
    Assert-Reference (@($partialAudit | Where-Object Status -eq 'MissingExport').Count -gt 0) 'Partial export was not reported.'
    Assert-Reference (@($partialAudit | Where-Object Source -eq 'SavedCache').Count -eq 2) 'Partial export lost saved cached labels.'
    $manifest.Lists[0].Status = 'Success'
    $manifest.Lists[0].RowCount = 99
    $manifest | ConvertTo-Json -Depth 30 | Set-Content $manifestPath -Encoding UTF8
    $rejected = $false
    try { $null = Import-HtmlReferenceData $bundle } catch { $rejected = $true }
    Assert-Reference $rejected 'Truncated reference CSV was accepted.'
    $manifest.Lists[0].RowCount = 5
    $manifest.Lists[0].DataPath = '..\outside.csv'
    $manifest | ConvertTo-Json -Depth 30 | Set-Content $manifestPath -Encoding UTF8
    $rejected = $false
    try { $null = Import-HtmlReferenceData $bundle } catch { $rejected = $true }
    Assert-Reference $rejected 'Escaping reference CSV path was accepted.'
    if ($ExtractedSampleDirectory) {
        $sampleRoot = Join-Path $testRoot 'RealSample'
        New-Item $sampleRoot -ItemType Directory | Out-Null
        $dependencyPath = Join-Path $sampleRoot 'Dependencies.json'
        & (Join-Path $repoRoot 'GetXsnReferenceDependencies.ps1') -ExtractedXsnDirectory $ExtractedSampleDirectory `
            -TemplateUrl 'http://sharepoint/sites/notices/red/Applications/Forms/template.xsn' -OutputPath $dependencyPath
        $dependencies = Get-Content $dependencyPath -Raw | ConvertFrom-Json
        $sampleTemplate = $dependencies.Templates[0]
        $countries = @($sampleTemplate.Bindings | Where-Object ConnectionName -eq 'NoticeCountries')
        Assert-Reference ($countries.Count -eq 4) 'Sample discovery did not identify all four Country controls.'
        Assert-Reference (@($countries | Where-Object { $_.ValueField -ne 'Value' -or $_.DisplayField -ne 'Title' }).Count -eq 0) 'Sample Country value/display columns are incorrect.'
        [pscustomobject]@{ ID = '999'; Value = 'ARCHIVE_COUNTRY'; Title = 'Offline Country Resolution' } |
            Export-Csv (Join-Path $sampleRoot 'Countries.csv') -NoTypeInformation -Encoding UTF8
        [pscustomobject]@{
            SchemaVersion = 1; ExportedAtUtc = '2026-10-27T00:00:00Z'; Dependencies = $dependencies
            Lists = @([pscustomobject]@{
                SiteUrl = $countries[0].SiteUrl; ListId = $countries[0].ListId
                Status = 'Success'; Error = ''; DataPath = 'Countries.csv'; RowCount = 1; SiteCollectionUrl = 'http://sharepoint/sites/notices'
            })
            Users = @()
        } | ConvertTo-Json -Depth 40 | Set-Content (Join-Path $sampleRoot 'Manifest.json') -Encoding UTF8
        $sampleDocument = Read-XsnXml (Join-Path $ExtractedSampleDirectory 'sampledata.xml')
        $instruction = $sampleDocument.SelectSingleNode("/processing-instruction('mso-infoPathSolution')")
        if ($null -ne $instruction) { $null = $sampleDocument.RemoveChild($instruction) }
        $instruction = $sampleDocument.CreateProcessingInstruction('mso-infoPathSolution',
            ('name="{0}" solutionVersion="{1}" href="{2}"' -f $sampleTemplate.SolutionName, $sampleTemplate.SolutionVersion, $sampleTemplate.TemplateUrl))
        $null = $sampleDocument.InsertBefore($instruction, $sampleDocument.DocumentElement)
        $sampleManager = New-HtmlReferenceNamespaceManager $sampleDocument $sampleTemplate
        $sampleFieldCount = 0
        foreach ($binding in $countries) {
            foreach ($node in $sampleDocument.SelectNodes($binding.XPath, $sampleManager)) {
                $node.InnerText = 'ARCHIVE_COUNTRY'
                $sampleFieldCount++
            }
        }
        Assert-Reference ($sampleFieldCount -eq 4) 'Sample dependency paths did not select all four Country fields.'
        $sampleAudit = [System.Collections.Generic.List[object]]::new()
        Set-HtmlReferenceValues $sampleDocument (Import-HtmlReferenceData $sampleRoot) $sampleAudit 'sampledata.xml'
        Assert-Reference (@($sampleAudit | Where-Object DisplayValue -eq 'Offline Country Resolution').Count -eq 4) 'Sample Country fields did not resolve offline.'
        $sampleHtmlPath = Join-Path $sampleRoot 'Sample.html'
        $null = Write-XsnArchiveHtml -Document $sampleDocument -Configuration (Import-HtmlFieldRules) `
            -ResourceDirectory $ExtractedSampleDirectory -TemplateHash 'sample-offline' -HtmlPath $sampleHtmlPath -OutputRoot $sampleRoot -Cache @{}
        $sampleHtml = Get-Content $sampleHtmlPath -Raw
        Assert-Reference ([regex]::Matches($sampleHtml, 'Offline Country Resolution').Count -eq 4) 'Sample XSN HTML did not display all four offline Country labels.'
        Assert-Reference (-not $sampleHtml.Contains('ARCHIVE_COUNTRY')) 'Sample XSN HTML retained a resolved Country code.'
        Write-Host 'PASS: real XSN dependency discovery and offline labels in all four Country controls.'
    }
    Write-Host 'PASS: offline references, cached labels, scoped users, ambiguity, both HTML paths, raw selections, CSV preservation, and partial exports.'
}
finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
