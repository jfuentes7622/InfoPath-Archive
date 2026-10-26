<#
.SYNOPSIS
    Verifies local XSN metadata processing, copied exports, and pipeline step 1.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$processor = Join-Path -Path $repoRoot -ChildPath 'ProcessXsnTemplates.ps1'
$testRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("InfoPathXsnProcessTest-" + [guid]::NewGuid().ToString('N'))

function New-TestCab {
    param([string]$Name, [string]$Manifest)

    $directory = Join-Path -Path $testRoot -ChildPath $Name
    New-Item -Path $directory -ItemType Directory -Force | Out-Null
    $manifestPath = Join-Path -Path $directory -ChildPath 'manifest.xsf'
    $viewPath = Join-Path -Path $directory -ChildPath 'view1.xsl'
    $cabPath = Join-Path -Path $directory -ChildPath 'template.xsn'
    $ddfPath = Join-Path -Path $directory -ChildPath 'template.ddf'
    $Manifest | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    '<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0" />' |
        Set-Content -LiteralPath $viewPath -Encoding UTF8
    @(
        '.OPTION EXPLICIT'
        '.Set Cabinet=ON'
        '.Set Compress=ON'
        '.Set CabinetNameTemplate=template.xsn'
        ('.Set DiskDirectoryTemplate="{0}"' -f $directory)
        ('.Set RptFileName="{0}"' -f (Join-Path -Path $directory -ChildPath 'makecab.rpt'))
        ('.Set InfFileName="{0}"' -f (Join-Path -Path $directory -ChildPath 'makecab.inf'))
        ('"{0}" manifest.xsf' -f $manifestPath)
        ('"{0}" view1.xsl' -f $viewPath)
    ) | Set-Content -LiteralPath $ddfPath -Encoding ASCII
    $output = & "$env:SystemRoot\System32\makecab.exe" /F $ddfPath 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "makecab failed: $($output -join ' ')"
    }
    return ,([System.IO.File]::ReadAllBytes($cabPath))
}

function Invoke-TestProcessing {
    param([string]$Case, [string]$ExpectedError = '')

    $output = Join-Path -Path $testRoot -ChildPath $Case
    $caught = ''
    try {
        & $processor -InputCsvPath $copiedCsv -OutputDirectory $output 6>$null
    }
    catch {
        $caught = $_.Exception.Message
    }
    if (($ExpectedError -and $caught -notlike $ExpectedError) -or (-not $ExpectedError -and $caught)) {
        throw "Case '$Case' expected '$ExpectedError', got '$caught'."
    }
    if (-not (Test-Path -LiteralPath (Join-Path -Path $output -ChildPath 'ProcessXsnTemplates.log'))) {
        throw "Case '$Case' did not retain its processing log."
    }
    $inventory = Join-Path -Path $output -ChildPath 'XsnTemplateVersions.csv'
    if (Test-Path -LiteralPath $inventory) {
        return @(Import-Csv -LiteralPath $inventory)
    }
}

function Write-TestInventory {
    param([object[]]$Rows)
    $Rows | Export-Csv -LiteralPath $copiedCsv -NoTypeInformation -Encoding UTF8
}

try {
    New-Item -Path $testRoot -ItemType Directory | Out-Null
    $manifestFormat = '<xsf:xDocumentClass xmlns:xsf="http://schemas.microsoft.com/office/infopath/2003/solutionDefinition" solutionVersion="{0}" name="{1}" />'
    $bytes3 = New-TestCab -Name 'Cab3' -Manifest ($manifestFormat -f '1.0.0.3', 'urn:form:A')
    $bytes2 = New-TestCab -Name 'Cab2' -Manifest ($manifestFormat -f '1.0.0.2', 'urn:form:A')
    $bytesB = New-TestCab -Name 'CabB' -Manifest ($manifestFormat -f '1.0.0.3', 'urn:form:B')
    $serverExport = Join-Path -Path $testRoot -ChildPath 'ServerExport'
    $downloadRows = @(
        foreach ($spec in @(
            @{ Folder = 'A'; Label = '3.0'; Current = $true; Bytes = $bytes3 }
            @{ Folder = 'A'; Label = '2.1'; Current = $false; Bytes = $bytes2 }
            @{ Folder = 'B'; Label = '1.0'; Current = $true; Bytes = $bytesB }
        )) {
            $relative = Join-Path -Path $spec.Folder -ChildPath ("template_sp-{0}.xsn" -f $spec.Label)
            $path = Join-Path -Path $serverExport -ChildPath $relative
            New-Item -Path (Split-Path -Path $path -Parent) -ItemType Directory -Force | Out-Null
            [System.IO.File]::WriteAllBytes($path, $spec.Bytes)
            [pscustomobject][ordered]@{
                SiteUrl = 'http://sharepoint/sites/Finance'
                LibraryName = 'Applications'
                TemplateUrl = "http://sharepoint/sites/Finance/$($spec.Folder)/template.xsn"
                ConfiguredBy = 'Library'
                SharePointVersion = $spec.Label
                IsCurrentVersion = $spec.Current
                Modified = '2026-10-26T12:00:00.0000000'
                XsnRelativePath = $relative
                XsnPath = $path
                Status = 'Success'
                Error = ''
            }
        }
    )
    $serverCsv = Join-Path -Path $serverExport -ChildPath 'XsnTemplateDownloads.csv'
    $downloadRows | Export-Csv -LiteralPath $serverCsv -NoTypeInformation -Encoding UTF8
    $copiedExport = Join-Path -Path $testRoot -ChildPath 'Copied Export'
    Copy-Item -LiteralPath $serverExport -Destination $copiedExport -Recurse
    Remove-Item -LiteralPath $serverExport -Recurse -Force
    $copiedCsv = Join-Path -Path $copiedExport -ChildPath 'XsnTemplateDownloads.csv'
    $originalCsvHash = (Get-FileHash -LiteralPath $copiedCsv).Hash
    $rows = @(Invoke-TestProcessing -Case 'Processed')
    if ($rows.Count -ne 3 -or @($rows | Where-Object Status -ne 'Success').Count -ne 0) {
        throw 'Expected three locally processed XSN versions.'
    }
    if (($rows.SharePointVersion -join ',') -ne '3.0,2.1,1.0' -or
        ($rows.SolutionVersion -join ',') -ne '1.0.0.3,1.0.0.2,1.0.0.3' -or
        ($rows.SolutionName -join ',') -ne 'urn:form:A,urn:form:A,urn:form:B') {
        throw 'Local processing did not preserve SharePoint history or extract solution versions and identities.'
    }
    foreach ($row in $rows) {
        if (-not $row.XsnPath.StartsWith($copiedExport + '\') -or
            -not $row.ManifestPath.StartsWith((Join-Path -Path $testRoot -ChildPath 'Processed') + '\') -or
            -not (Test-Path -LiteralPath $row.ManifestPath) -or
            (Get-FileHash -LiteralPath $row.XsnPath -Algorithm SHA256).Hash -ne $row.Sha256) {
            throw 'Relocated XSN paths, manifest output paths, or hashes are incorrect.'
        }
    }
    if ((Get-FileHash -LiteralPath $copiedCsv).Hash -ne $originalCsvHash) {
        throw 'Processing modified the download inventory.'
    }

    $currentPath = Join-Path -Path $copiedExport -ChildPath $downloadRows[0].XsnRelativePath
    [System.IO.File]::WriteAllBytes($currentPath, [System.Text.Encoding]::UTF8.GetBytes('not a cabinet'))
    $rows = @(Invoke-TestProcessing -Case 'Processed' -ExpectedError '*processing incomplete: 1 failure*')
    if ($rows[0].Status -ne 'Failed' -or $rows[0].SolutionVersion -or $rows[0].ManifestPath -or @($rows | Where-Object Status -eq 'Success').Count -ne 2) {
        throw 'A corrupt XSN reused stale metadata or prevented other versions from processing.'
    }
    [System.IO.File]::WriteAllBytes($currentPath, (New-TestCab -Name 'NoVersion' -Manifest '<xDocumentClass />'))
    $rows = @(Invoke-TestProcessing -Case 'MissingVersion' -ExpectedError '*processing incomplete: 1 failure*')
    if ($rows[0].Error -notlike '*no solutionVersion*') { throw 'Missing solutionVersion was not reported.' }
    [System.IO.File]::WriteAllBytes($currentPath, $bytes3)
    Remove-Item -LiteralPath $currentPath
    $rows = @(Invoke-TestProcessing -Case 'MissingFile' -ExpectedError '*processing incomplete: 1 failure*')
    if ($rows[0].Error -notlike '*file is missing*') { throw 'Missing copied XSN was not reported.' }
    [System.IO.File]::WriteAllBytes($currentPath, $bytes3)

    $badRows = @($downloadRows | Select-Object *)
    $badRows[0].Status = 'Failed'
    $badRows[0].Error = 'Access denied on SPS'
    Write-TestInventory -Rows $badRows
    $rows = @(Invoke-TestProcessing -Case 'FailedDownload' -ExpectedError '*processing incomplete: 1 failure*')
    if ($rows[0].DownloadStatus -ne 'Failed' -or $rows[0].Error -notlike '*Access denied on SPS*') {
        throw 'Failed server download was not preserved as a local failure.'
    }
    $badRows[0].Status = 'Success'
    $badRows[0].XsnRelativePath = '..\outside.xsn'
    Write-TestInventory -Rows $badRows
    $rows = @(Invoke-TestProcessing -Case 'EscapingPath' -ExpectedError '*processing incomplete: 1 failure*')
    if ($rows[0].Error -notlike '*escapes the copied export folder*') { throw 'Path escaping the export was accepted.' }
    $badRows[0].XsnRelativePath = $currentPath
    Write-TestInventory -Rows $badRows
    $rows = @(Invoke-TestProcessing -Case 'AbsolutePath' -ExpectedError '*processing incomplete: 1 failure*')
    if ($rows[0].Error -notlike '*must be a nonempty relative path*') { throw 'Absolute relative-path value was accepted.' }
    $downloadRows | Select-Object * -ExcludeProperty XsnRelativePath | Export-Csv -LiteralPath $copiedCsv -NoTypeInformation
    $null = Invoke-TestProcessing -Case 'MissingColumn' -ExpectedError "*missing 'XsnRelativePath'*"
    Write-TestInventory -Rows $downloadRows

    $source = Join-Path -Path $testRoot -ChildPath 'Source'
    New-Item -Path $source -ItemType Directory | Out-Null
    @'
<?xml version="1.0" encoding="utf-8"?>
<?mso-infoPathSolution solutionVersion="1.0.0.3" name="urn:form:A" href="template.xsn"?>
<my:FormData xmlns:my="urn:form:A"><my:Requestor_Name>Alice</my:Requestor_Name></my:FormData>
'@ | Set-Content -LiteralPath (Join-Path -Path $source -ChildPath 'Sample.xml') -Encoding UTF8
    $pipelineOutput = Join-Path -Path $testRoot -ChildPath 'PipelineOutput'
    & (Join-Path -Path $repoRoot -ChildPath 'RunInfoPathPipeline.ps1') `
        -SourceDirectory $source -OutputRootDirectory $pipelineOutput -XsnExportCsvPath $copiedCsv -StartStep 1 6>$null
    $pipelineInventory = Join-Path -Path $pipelineOutput -ChildPath 'XSNTemplates\XsnTemplateVersions.csv'
    if (-not (Test-Path -LiteralPath $pipelineInventory) -or
        @(Import-Csv -LiteralPath $pipelineInventory).Count -ne 3 -or
        -not (Test-Path -LiteralPath (Join-Path -Path $pipelineOutput -ChildPath 'Html\Sample.html'))) {
        throw 'The full pipeline did not process XSN information under its output root and render the XML.'
    }
    $badRows[0].Status = 'Failed'
    Write-TestInventory -Rows $badRows
    $failedPipelineOutput = Join-Path -Path $testRoot -ChildPath 'FailedPipeline'
    $messages = [System.Collections.Generic.List[string]]::new()
    $caught = $false
    try {
        & (Join-Path -Path $repoRoot -ChildPath 'RunInfoPathPipeline.ps1') `
            -SourceDirectory $source -OutputRootDirectory $failedPipelineOutput -XsnExportCsvPath $copiedCsv -StartStep 1 6>&1 |
            ForEach-Object { $messages.Add($_.ToString()) }
    }
    catch {
        if ($_.Exception.Message -notlike '*XSN processing incomplete*') { throw }
        $caught = $true
    }
    if (-not $caught -or
        -not ($messages | Where-Object { $_ -like '[[]Error] Pipeline failed at Step 1/7:*' }) -or
        (Test-Path -LiteralPath (Join-Path -Path $failedPipelineOutput -ChildPath 'Reports\InfoPathVersionMap.csv'))) {
        throw 'Failed XSN processing did not identify step 1 and stop the remaining pipeline.'
    }
    Write-Host 'PASS: Local manifest processing, relocated exports, output folders, hashes, partial failures, stale-output protection, and seven-step pipeline.'
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
