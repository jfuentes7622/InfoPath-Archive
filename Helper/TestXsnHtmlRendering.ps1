<#
.SYNOPSIS
    Verifies optional XSN archival rendering, CAB assets, fallback, and field-rule formats.
#>
[CmdletBinding()]
param([string]$ExtractedSampleDirectory, [switch]$PrintToPdf)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'HtmlRendering.ps1')
. (Join-Path $PSScriptRoot 'XsnTemplates.ps1')
. (Join-Path $PSScriptRoot 'XsnHtmlRendering.ps1')
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('InfoPathXsnHtmlTest-' + [guid]::NewGuid().ToString('N'))

try {
    $package = Join-Path $testRoot 'Export'
    $source = Join-Path $testRoot 'Source'
    New-Item $package, $source -ItemType Directory -Force | Out-Null
    @'
<xsf:xDocumentClass xmlns:xsf="http://schemas.microsoft.com/office/infopath/2003/solutionDefinition" solutionVersion="1.0.0.1" name="urn:archive:test">
 <xsf:package><xsf:files><xsf:file name="view1.xsl"/><xsf:file name="image.png"/><xsf:file name="styles.css"/></xsf:files></xsf:package>
 <xsf:views default="Archive"><xsf:view name="Archive"><xsf:mainpane transform="view1.xsl"/></xsf:view></xsf:views>
</xsf:xDocumentClass>
'@ | Set-Content (Join-Path $package 'manifest.xsf') -Encoding UTF8
    @'
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:my="urn:archive:fields" xmlns:xd="http://schemas.microsoft.com/office/infopath/2003" xmlns:xdFormatting="http://schemas.microsoft.com/office/infopath/2003/xslt/Formatting" xmlns:xdXDocument="http://schemas.microsoft.com/office/infopath/2003/xslt/xDocument" version="1.0">
 <xsl:output method="html"/>
 <xsl:template match="/">
 <html><head><link rel="stylesheet" href="styles.css"/></head><body>
 <h1>Original form layout</h1><img src="image.png"/>
 <xsl:for-each select="my:Form">
 <span xd:binding="my:Date"><xsl:value-of select="xdFormatting:formatString(my:Date,'date','short')"/></span>
 <span xd:binding="my:Calculated"><xsl:value-of select="my:Calculated * 99"/></span>
 <select xd:binding="my:Lookup"><xsl:for-each select="xdXDocument:GetDOM('LiveSharePoint')/*"><option>Live lookup</option></xsl:for-each></select>
 <span xd:binding="my:SavedText"/>
 <input type="checkbox" xd:binding="my:Flag"><xsl:if test="my:Flag='1'"><xsl:attribute name="CHECKED">CHECKED</xsl:attribute></xsl:if></input>
 <input type="radio" value="Y" xd:binding="my:Radio"><xsl:if test="my:Radio='Y'"><xsl:attribute name="checked">checked</xsl:attribute></xsl:if></input>
 <input type="radio" value="N" xd:binding="my:Radio"><xsl:if test="my:Radio='N'"><xsl:attribute name="checked">checked</xsl:attribute></xsl:if></input>
 <input type="button" value="Submit"/>
 <xsl:if test="false()"><div style="display:none; visibility:hidden"><h2>Hidden section</h2><span xd:binding="my:Empty"/></div></xsl:if>
 <xsl:for-each select="my:Row"><div class="record"><span xd:binding="my:Value"/></div></xsl:for-each>
 <div xd:binding="my:Rich"/>
 </xsl:for-each>
 </body></html>
 </xsl:template>
</xsl:stylesheet>
'@ | Set-Content (Join-Path $package 'view1.xsl') -Encoding UTF8
    [System.IO.File]::WriteAllBytes((Join-Path $package 'image.png'), [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aCJkAAAAASUVORK5CYII='))
    '.record { background-image:url("image.png"); }' | Set-Content (Join-Path $package 'styles.css') -Encoding ASCII
    @'
<?mso-infoPathSolution solutionVersion="1.0.0.1" name="urn:archive:test" href="http://sharepoint/forms/test.xsn"?>
<my:Form xmlns:my="urn:archive:fields" xmlns:h="http://www.w3.org/1999/xhtml">
 <my:Date>2026-10-27</my:Date><my:Calculated>17</my:Calculated><my:Lookup>314</my:Lookup><my:SavedText>314;#Saved label</my:SavedText>
 <my:Flag>1</my:Flag><my:Radio>Y</my:Radio><my:Empty/>
 <my:Row><my:Value>First row</my:Value></my:Row><my:Row><my:Value>Second row</my:Value></my:Row>
 <my:Rich><h:p><h:b>Rich content</h:b><h:script>unsafe()</h:script></h:p></my:Rich>
</my:Form>
'@ | Set-Content (Join-Path $source 'Sample.xml') -Encoding UTF8
    $originalHash = (Get-FileHash (Join-Path $source 'Sample.xml')).Hash
    $nested = Join-Path $source 'Nested Forms'
    New-Item $nested -ItemType Directory | Out-Null
    Copy-Item (Join-Path $source 'Sample.xml') (Join-Path $nested 'Sample #1.xml')
    $rulesPath = Join-Path $testRoot 'HtmlFieldRules.json'
    '{"DateFormat":"dd.MM.yyyy","DateTimeFormat":"dd.MM.yyyy HH:mm"}' | Set-Content $rulesPath -Encoding UTF8
    $ddf = Join-Path $testRoot 'template.ddf'
    @(
        '.OPTION EXPLICIT'
        '.Set Cabinet=ON'
        '.Set Compress=ON'
        '.Set CabinetNameTemplate=template.xsn'
        ('.Set DiskDirectoryTemplate="{0}"' -f $package)
        ('.Set RptFileName="{0}"' -f (Join-Path $testRoot 'makecab.rpt'))
        ('.Set InfFileName="{0}"' -f (Join-Path $testRoot 'makecab.inf'))
        ('"{0}" manifest.xsf' -f (Join-Path $package 'manifest.xsf'))
        ('"{0}" view1.xsl' -f (Join-Path $package 'view1.xsl'))
        ('"{0}" image.png' -f (Join-Path $package 'image.png'))
        ('"{0}" styles.css' -f (Join-Path $package 'styles.css'))
    ) | Set-Content $ddf -Encoding ASCII
    $cabOutput = & "$env:SystemRoot\System32\makecab.exe" /F $ddf 2>&1
    if ($LASTEXITCODE -ne 0) { throw "makecab failed: $($cabOutput -join ' ')" }
    $exportCsv = Join-Path $package 'XsnTemplateDownloads.csv'
    [pscustomobject]@{
        SiteUrl='http://sharepoint'; LibraryName='Forms'; TemplateUrl='http://sharepoint/forms/test.xsn'
        ConfiguredBy='Library'; SharePointVersion='1.0'; IsCurrentVersion='True'; Modified='2026-10-27'
        XsnRelativePath='template.xsn'; Status='Success'; Error=''
    } | Export-Csv $exportCsv -NoTypeInformation
    foreach ($mode in @('Xsn', 'Default', 'Missing', 'Broken', 'NextVersion', 'NoHigherVersion', 'OtherForm')) {
        $output = Join-Path $testRoot $mode
        $parameters = @{
            SourceDirectory=$source; OutputRootDirectory=$output; StartStep=1; FieldRulesPath=$rulesPath
            UseXSN4HTML=($mode -ne 'Default'); Recurse=$true
        }
        if ($mode -ne 'Missing') { $parameters.XsnExportCsvPath = $exportCsv }
        if ($mode -eq 'Broken') { $parameters.XsnExportCsvPath = Join-Path $package 'missing.csv' }
        if ($mode -in @('NextVersion', 'NoHigherVersion', 'OtherForm')) {
            $versionSource = Join-Path $testRoot ($mode + 'Source')
            New-Item $versionSource -ItemType Directory | Out-Null
            $versionXml = Get-Content (Join-Path $source 'Sample.xml') -Raw
            if ($mode -eq 'OtherForm') {
                $versionXml = $versionXml.Replace('name="urn:archive:test"', 'name="urn:archive:other"')
            }
            else {
                $requestedVersion = if ($mode -eq 'NextVersion') { '1.0.0.0' } else { '1.0.0.2' }
                $versionXml = $versionXml.Replace('solutionVersion="1.0.0.1"', ('solutionVersion="' + $requestedVersion + '"'))
            }
            $versionXml | Set-Content (Join-Path $versionSource 'Sample.xml') -Encoding UTF8
            $parameters.SourceDirectory = $versionSource
        }
        if ($PrintToPdf -and $mode -in @('Xsn', 'NextVersion')) { $parameters.PrintToPdf = 'true' }
        & (Join-Path $repoRoot 'RunInfoPathPipeline.ps1') @parameters 6>$null
        $htmlPath = Join-Path $output 'Html\Sample.html'
        $html = Get-Content $htmlPath -Raw
        $audit = @(Import-Csv (Join-Path $output 'Html\InfoPathXsnMatches.csv'))[0]
        if (-not $html.Contains('27.10.2026')) { throw "$mode did not use JSON date format." }
        if ($mode -eq 'Xsn') {
            foreach ($value in @('Original form layout', 'Hidden section', 'First row', 'Second row', 'Saved label', '314', '17', '[X]', '(X)', '( )', 'Rich content')) {
                if (-not $html.Contains($value)) { throw "XSN output missing '$value'." }
            }
            if ($html -match '(?i)<input|<select|<script|display\s*:\s*none|visibility\s*:\s*hidden|unsafe\(\)|Live lookup') { throw 'XSN output retained interactive, hidden, or unsafe content.' }
            if ($audit.ActualRenderer -ne 'XSN' -or $audit.ViewName -ne 'Archive' -or $audit.FallbackReason) { throw 'XSN renderer was not used.' }
            $asset = @(Get-ChildItem (Join-Path $output 'Html\Assets') -Filter image.png -Recurse)
            if ($asset.Count -ne 1 -or (Get-FileHash $asset[0].FullName).Hash -ne (Get-FileHash (Join-Path $package 'image.png')).Hash) { throw 'Image was not copied once with original bytes.' }
            if (@(Get-ChildItem (Join-Path $output 'Html\Assets') -Filter styles.css -Recurse).Count -ne 1) { throw 'CSS was not copied.' }
            $nestedHtmlPath = Join-Path $output 'Html\Nested Forms\Sample #1.html'
            $nestedHtml = Get-Content $nestedHtmlPath -Raw
            $imageReference = [regex]::Match($nestedHtml, '<img[^>]+src="([^"]+)"').Groups[1].Value
            $resolvedImage = [uri]::new([uri]::new($nestedHtmlPath), $imageReference).LocalPath
            if (-not (Test-Path -LiteralPath $resolvedImage)) { throw 'Nested HTML image reference is broken.' }
            if ($PrintToPdf) {
                $pdf = [System.IO.File]::ReadAllBytes((Join-Path $output 'Html\Sample.pdf'))
                if ($pdf.Length -lt 1000 -or [System.Text.Encoding]::ASCII.GetString($pdf, 0, 5) -ne '%PDF-' -or -not $audit.PdfPath) { throw 'XSN PDF output or audit is invalid.' }
            }
        }
        elseif ($mode -eq 'NextVersion') {
            if ($audit.ActualRenderer -ne 'XSN' -or $audit.MatchStatus -ne 'NextVersion' -or
                $audit.RequestedVersion -ne '1.0.0.0' -or $audit.MatchedVersion -ne '1.0.0.1' -or
                $audit.ViewName -ne 'Archive' -or $audit.FallbackReason -or -not $html.Contains('Original form layout')) {
                throw 'Nearest-higher XSN view was not rendered or audited correctly.'
            }
            if ($PrintToPdf) {
                $pdf = [System.IO.File]::ReadAllBytes((Join-Path $output 'Html\Sample.pdf'))
                if ($pdf.Length -lt 1000 -or [System.Text.Encoding]::ASCII.GetString($pdf, 0, 5) -ne '%PDF-' -or -not $audit.PdfPath) {
                    throw 'Nearest-higher XSN PDF output or audit is invalid.'
                }
            }
        }
        elseif ($audit.ActualRenderer -ne 'Generated' -or ($mode -ne 'Default' -and -not $audit.FallbackReason)) { throw "$mode fallback audit is incorrect." }
        if ($mode -in @('NoHigherVersion', 'OtherForm') -and $audit.MatchStatus -ne 'NoMatch') {
            throw "$mode incorrectly matched an older or different form's template."
        }
        $csv = @(Get-ChildItem (Join-Path $output 'Reports\Normalized') -Filter Main.csv -Recurse)[0]
        if (-not (Get-Content $csv.FullName -Raw).Contains('2026-10-27')) { throw 'CSV data was reformatted.' }
    }
    if ((Get-FileHash (Join-Path $source 'Sample.xml')).Hash -ne $originalHash) { throw 'Source XML changed.' }
    $xsnOutput = Join-Path $testRoot 'Xsn'
    $processedInventory = Join-Path $xsnOutput 'XSNTemplates\XsnTemplateVersions.csv'
    $processedRow = @(Import-Csv $processedInventory)[0]
    $generatedMap = @(Import-Csv (Join-Path $xsnOutput 'Templates\InfoPathXsltGroupMap.csv'))[0]
    $backupXslt = $generatedMap.XsltPath + '.saved'
    Move-Item -LiteralPath $generatedMap.XsltPath -Destination $backupXslt
    & (Join-Path $repoRoot 'RunInfoPathPipeline.ps1') -SourceDirectory $source -OutputRootDirectory $xsnOutput -StartStep 7 -UseXSN4HTML $true -Recurse -FieldRulesPath $rulesPath 6>$null
    $independentAudit = @(Import-Csv (Join-Path $xsnOutput 'Html\InfoPathXsnMatches.csv'))
    if (@($independentAudit | Where-Object ActualRenderer -ne 'XSN').Count -ne 0) { throw 'XSN renderer unnecessarily required the generated XSLT.' }
    Move-Item -LiteralPath $backupXslt -Destination $generatedMap.XsltPath
    $extractedView = Join-Path (Split-Path $processedInventory -Parent) ($processedRow.ResourcesRelativePath + '\view1.xsl')
    (Get-Content $extractedView -Raw).Replace('<h1>Original form layout</h1>', '<h1><xsl:value-of select="sum(//my:Calculated)"/></h1>') |
        Set-Content $extractedView -Encoding UTF8
    & (Join-Path $repoRoot 'RunInfoPathPipeline.ps1') -SourceDirectory $source -OutputRootDirectory $xsnOutput -StartStep 7 -UseXSN4HTML $true -Recurse -FieldRulesPath $rulesPath 6>$null
    $fallbackAudit = @(Import-Csv (Join-Path $xsnOutput 'Html\InfoPathXsnMatches.csv'))
    if (@($fallbackAudit | Where-Object { $_.ActualRenderer -ne 'Generated' -or $_.FallbackReason -notlike '*runtime computation*' }).Count -ne 0) { throw 'Unbound computed output did not trigger an explicit fallback.' }
    Move-Item -LiteralPath $generatedMap.XsltPath -Destination $backupXslt
    $failedExplicitly = $false
    try {
        & (Join-Path $repoRoot 'RunInfoPathPipeline.ps1') -SourceDirectory $source -OutputRootDirectory $xsnOutput -StartStep 7 -UseXSN4HTML $true -Recurse -FieldRulesPath $rulesPath 6>$null
    }
    catch {
        if ($_.Exception.Message -notlike '*rendering failed*') { throw }
        $failedExplicitly = $true
    }
    if (-not $failedExplicitly) { throw 'Failure of both renderers was reported as success.' }
    Move-Item -LiteralPath $backupXslt -Destination $generatedMap.XsltPath
    foreach ($format in @('', '%')) {
        @{ DateFormat=$format; Fields=@() } | ConvertTo-Json | Set-Content $rulesPath
        $rejected = $false
        try { $null = Import-HtmlFieldRules $rulesPath } catch { $rejected = $true }
        if (-not $rejected) { throw 'Invalid JSON date format was accepted.' }
    }
    if ($ExtractedSampleDirectory) {
        $sampleOutput = Join-Path $testRoot 'RealSample'
        New-Item $sampleOutput -ItemType Directory | Out-Null
        $sampleHtml = Join-Path $sampleOutput 'Sample.html'
        $viewName = Write-XsnArchiveHtml -Document (Read-XsnXml (Join-Path $ExtractedSampleDirectory 'sampledata.xml')) -Configuration (Import-HtmlFieldRules) -ResourceDirectory $ExtractedSampleDirectory -TemplateHash 'sample' -HtmlPath $sampleHtml -OutputRoot $sampleOutput -Cache @{}
        $sampleText = Get-Content $sampleHtml -Raw
        if ((Get-Item $sampleHtml).Length -lt 10000 -or $sampleText -match '(?i)<input|<select|<script|display\s*:\s*none|visibility\s*:\s*hidden') { throw 'Extracted sample rendering failed static-output checks.' }
        Write-Host "PASS: Extracted sample view '$viewName'; $(@(Get-ChildItem (Join-Path $sampleOutput 'Assets') -File -Recurse).Count) local assets."
    }
    Write-Host 'PASS: XSN CAB rendering, exact/nearest-higher matching, older/different-form exclusion, stored values, show-all sections, static selections, resources, JSON formats, CSV independence, and fallback.'
}
finally {
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
}
