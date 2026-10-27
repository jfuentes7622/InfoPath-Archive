<#
.SYNOPSIS
    Verifies dependency discovery directly from downloaded XSN CAB exports.
#>
#requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repository = Split-Path $PSScriptRoot -Parent
$work = Join-Path ([IO.Path]::GetTempPath()) ('InfoPathDownloadDiscoveryTest-' + [guid]::NewGuid().ToString('N'))
$originalTemp = $env:TEMP
$originalTmp = $env:TMP

function Assert-DownloadDiscovery {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Assert-ExtractionCleaned {
    $remaining = @(Get-ChildItem -LiteralPath $temporaryRoot -Directory)
    Assert-DownloadDiscovery ($remaining.Count -eq 0) 'Temporary extraction directories were not removed.'
}

try {
    $export = Join-Path $work 'Downloads\Run1'
    $temporaryRoot = Join-Path $work 'Temp'
    $package = Join-Path $work 'Package'
    New-Item $export, $temporaryRoot, $package -ItemType Directory -Force | Out-Null
    $env:TEMP = $temporaryRoot
    $env:TMP = $temporaryRoot
    $view = @'
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:my="urn:download:fields" xmlns:xd="http://schemas.microsoft.com/office/infopath/2003" xmlns:xdXDocument="http://schemas.microsoft.com/office/infopath/2003/xslt/xDocument" xmlns:dfs="urn:dfs" xmlns:d="urn:rows">
 <xsl:template match="/"><html><body><xsl:for-each select="my:Root"><select xd:binding="my:Country">
  <xsl:for-each select="xdXDocument:GetDOM(&quot;Countries&quot;)/dfs:myFields/dfs:dataFields/d:Item">
   <option><xsl:attribute name="value"><xsl:value-of select="d:Value"/></xsl:attribute><xsl:value-of select="d:Title"/></option>
  </xsl:for-each>
 </select></xsl:for-each></body></html></xsl:template>
</xsl:stylesheet>
'@
    $view | Set-Content (Join-Path $package 'first.xsl') -Encoding UTF8
    $view.Replace('my:Country', 'my:OtherCountry') | Set-Content (Join-Path $package 'second.xsl') -Encoding UTF8
    $rows = @(
        foreach ($version in 1..2) {
            @"
<xsf:xDocumentClass xmlns:xsf="http://schemas.microsoft.com/office/infopath/2003/solutionDefinition" xmlns:my="urn:download:fields" name="urn:download:test" solutionVersion="1.0.0.$version">
 <xsf:views default="Primary"><xsf:view name="Primary"><xsf:mainpane transform="first.xsl"/></xsf:view><xsf:view name="Second"><xsf:mainpane transform="second.xsl"/></xsf:view></xsf:views>
 <xsf:dataObjects><xsf:dataObject name="Countries"><xsf:query><xsf:sharepointListAdapterRW name="Countries" siteURL="/sites/lookups/" sharePointListID="{11111111-1111-1111-1111-111111111111}"><xsf:field internalName="Value"/><xsf:field internalName="Title"/></xsf:sharepointListAdapterRW></xsf:query></xsf:dataObject></xsf:dataObjects>
</xsf:xDocumentClass>
"@ | Set-Content (Join-Path $package 'manifest.xsf') -Encoding UTF8
            $cabDirectory = Join-Path $export "hash\version$version"
            New-Item $cabDirectory -ItemType Directory -Force | Out-Null
            $ddf = Join-Path $package 'template.ddf'
            @(
                '.OPTION EXPLICIT'
                '.Set Cabinet=ON'
                '.Set Compress=ON'
                '.Set CabinetNameTemplate=template.xsn'
                ('.Set DiskDirectoryTemplate="{0}"' -f $cabDirectory)
                ('.Set RptFileName="{0}"' -f (Join-Path $package 'makecab.rpt'))
                ('.Set InfFileName="{0}"' -f (Join-Path $package 'makecab.inf'))
                ('"{0}" manifest.xsf' -f (Join-Path $package 'manifest.xsf'))
                ('"{0}" first.xsl' -f (Join-Path $package 'first.xsl'))
                ('"{0}" second.xsl' -f (Join-Path $package 'second.xsl'))
            ) | Set-Content $ddf -Encoding ASCII
            $cabOutput = & "$env:SystemRoot\System32\makecab.exe" /F $ddf 2>&1
            if ($LASTEXITCODE -ne 0) { throw "makecab failed: $($cabOutput -join ' ')" }
            [pscustomobject]@{
                SiteUrl = 'https://example.test/sites/owner'; LibraryName = 'Applications'
                TemplateUrl = 'https://example.test/sites/owner/Applications/Forms/template.xsn'; ConfiguredBy = 'Library'
                SharePointVersion = "$version.0"; IsCurrentVersion = ($version -eq 2); Modified = '2026-10-27'
                XsnRelativePath = "hash\version$version\template.xsn"; Status = 'Success'; Error = ''
            }
        }
    )
    $csv = Join-Path $export 'XsnTemplateDownloads.csv'
    $rows | Export-Csv $csv -NoTypeInformation -Encoding UTF8
    $originalCsvHash = (Get-FileHash $csv).Hash
    $originalHashes = @($rows | ForEach-Object { (Get-FileHash (Join-Path $export $_.XsnRelativePath)).Hash })
    $discovery = Join-Path $repository 'GetXsnReferenceDependencies.ps1'
    foreach ($inputPath in @($csv, (Join-Path $work 'Downloads'))) {
        $output = Join-Path $work ('Results\' + [guid]::NewGuid().ToString('N') + '\Dependencies.json')
        $result = & $discovery -XsnExportCsvPath $inputPath -OutputPath $output 6>$null
        $dependencies = Get-Content $output -Raw | ConvertFrom-Json
        Assert-DownloadDiscovery ($result -eq $output) 'Discovery did not return the dependency filename.'
        Assert-DownloadDiscovery ($dependencies.Templates.Count -eq 2) 'Not all downloaded template versions were discovered.'
        Assert-DownloadDiscovery ($dependencies.Lists.Count -eq 1) 'Shared list dependencies were not deduplicated.'
        Assert-DownloadDiscovery ($dependencies.Warnings.Count -eq 0) ("Valid downloaded templates produced discovery warnings: " + ($dependencies.Warnings -join '; '))
        foreach ($template in $dependencies.Templates) {
            Assert-DownloadDiscovery ($originalHashes -contains $template.Sha256) 'Dependency hash does not match the original downloaded XSN.'
            Assert-DownloadDiscovery ($template.Bindings.Count -eq 2) 'Declared views were not both inspected.'
            Assert-DownloadDiscovery (@($template.Bindings | Where-Object { $_.ValueField -ne 'Value' -or $_.DisplayField -ne 'Title' }).Count -eq 0) 'Actual value/display mapping was lost during temporary extraction.'
        }
        Assert-DownloadDiscovery (-not (Get-Content $output -Raw).Contains($temporaryRoot)) 'Dependency manifest depends on temporary extraction paths.'
        Assert-DownloadDiscovery ((Get-Content ($output + '.log') -Raw).Contains('solutionVersion=1.0.0.2')) 'Preparation diagnostics were not retained.'
        Assert-ExtractionCleaned
    }

    $directoryRejected = $false
    try { $null = & $discovery -XsnExportCsvPath $csv -OutputPath $work }
    catch { $directoryRejected = $_.Exception.Message -like '*JSON filename*' }
    Assert-DownloadDiscovery $directoryRejected 'Folder-only OutputPath did not produce an actionable error.'
    Assert-ExtractionCleaned

    $blockedOutput = Join-Path $work 'Blocked.json'
    New-Item ($blockedOutput + '.log') -ItemType Directory | Out-Null
    $writeThrew = $false
    try { $null = & $discovery -XsnExportCsvPath $csv -OutputPath $blockedOutput 6>$null }
    catch { $writeThrew = $true }
    Assert-DownloadDiscovery $writeThrew 'Log output failure was not surfaced.'
    Assert-ExtractionCleaned

    $failedRow = $rows[0].PSObject.Copy()
    $failedRow.Status = 'Failed'
    $failedRow.Error = 'Synthetic download failure'
    @($rows) + @($failedRow) | Export-Csv $csv -NoTypeInformation -Encoding UTF8
    $partialPath = Join-Path $work 'Partial.json'
    $partialThrew = $false
    try { $null = & $discovery -XsnExportCsvPath $csv -OutputPath $partialPath 6>$null }
    catch { $partialThrew = $_.Exception.Message -like '*Dependency discovery had failures*' }
    $partial = Get-Content $partialPath -Raw | ConvertFrom-Json
    Assert-DownloadDiscovery ($partialThrew -and $partial.Templates.Count -eq 2) 'Partial download failure lost successful dependency results or was reported as success.'
    Assert-DownloadDiscovery ((Get-Content ($partialPath + '.log') -Raw).Contains('Synthetic download failure')) 'Download failure details were lost during cleanup.'
    Assert-ExtractionCleaned
    $rows | Export-Csv $csv -NoTypeInformation -Encoding UTF8

    $badPath = Join-Path $export 'bad.xsn'
    [IO.File]::WriteAllText($badPath, 'Not a CAB')
    $badRow = $rows[0].PSObject.Copy()
    $badRow.XsnRelativePath = 'bad.xsn'
    $badExport = Join-Path $work 'Corrupt'
    New-Item $badExport -ItemType Directory | Out-Null
    Copy-Item $badPath (Join-Path $badExport 'bad.xsn')
    $badRow | Export-Csv (Join-Path $badExport 'XsnTemplateDownloads.csv') -NoTypeInformation -Encoding UTF8
    $badOutput = Join-Path $work 'Corrupt.json'
    $badThrew = $false
    try { $null = & $discovery -XsnExportCsvPath $badExport -OutputPath $badOutput 6>$null }
    catch { $badThrew = $true }
    $badResult = Get-Content $badOutput -Raw | ConvertFrom-Json
    Assert-DownloadDiscovery ($badThrew -and $badResult.Templates.Count -eq 0 -and $badResult.Warnings.Count -gt 0) 'Corrupt CAB was not reported explicitly.'
    Assert-DownloadDiscovery (Test-Path -LiteralPath ($badOutput + '.log')) 'Corrupt CAB diagnostics were not preserved.'
    Assert-ExtractionCleaned

    $ambiguousRoot = Join-Path $work 'Ambiguous'
    New-Item (Join-Path $ambiguousRoot 'One'), (Join-Path $ambiguousRoot 'Two') -ItemType Directory -Force | Out-Null
    Copy-Item $csv (Join-Path $ambiguousRoot 'One\XsnTemplateDownloads.csv')
    Copy-Item $csv (Join-Path $ambiguousRoot 'Two\XsnTemplateDownloads.csv')
    $ambiguousOutput = Join-Path $work 'Ambiguous.json'
    $ambiguousThrew = $false
    try { $null = & $discovery -XsnExportCsvPath $ambiguousRoot -OutputPath $ambiguousOutput 6>$null }
    catch { $ambiguousThrew = $true }
    Assert-DownloadDiscovery ($ambiguousThrew -and (Get-Content ($ambiguousOutput + '.log') -Raw).Contains('Multiple XSN download inventories')) 'Multiple exports were not rejected with useful diagnostics.'
    Assert-ExtractionCleaned
    Assert-DownloadDiscovery ((Get-FileHash $csv).Hash -eq $originalCsvHash) 'Download inventory was changed.'
    for ($index = 0; $index -lt $rows.Count; $index++) {
        Assert-DownloadDiscovery ((Get-FileHash (Join-Path $export $rows[$index].XsnRelativePath)).Hash -eq $originalHashes[$index]) 'Downloaded XSN was changed.'
    }
    Write-Host 'PASS: direct downloaded-CAB discovery, recursive folders, all versions/views, hashes, partial failures, diagnostics, output filenames, and temporary cleanup.'
}
finally {
    $env:TEMP = $originalTemp
    $env:TMP = $originalTmp
    if (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force }
}
