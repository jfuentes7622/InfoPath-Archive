<#
.SYNOPSIS
    Collects SharePoint InfoPath archival inputs into one portable export bundle.
.DESCRIPTION
    Run on SharePoint Server in Windows PowerShell 5.1 / Management Shell.
    Coordinates library files, library metadata, XSN versions, dependency
    discovery (including temporary extraction), and lookup/user snapshots.
    Each new run creates InfoPathExport_<timestamp>_<id> under OutputRootDirectory.
    Copy that entire folder offline, including nested XSN/reference exports.

    Continues independent steps after errors and uses partial XSN/dependency
    results where available. ExportManifest.json and Export.log persist progress.
    Any failed, blocked, or unfinished requested step makes the bundle Incomplete;
    throws only after attempting the remaining steps. Never changes SharePoint.
    This is not a transactional or historical snapshot. Protect exported data.
.PARAMETER SiteUrl
    Absolute URL of the SharePoint web containing the library, for a new run.
.PARAMETER LibraryName
    Display name of the source library, for a new run.
.PARAMETER OutputRootDirectory
    Parent folder for a NEW timestamped bundle. Required with SiteUrl/LibraryName.
.PARAMETER ResumeExportDirectory
    Existing bundle folder containing ExportManifest.json. Restores its source
    and options. Does not create a new bundle. Incompatible with new-run inputs.
.PARAMETER StartStep
    New runs start at 1. Resumes default to the first unfinished/failed step.
    An explicit 1-5 reruns that step and every later step. Earlier failed steps
    remain failed unless rerun; an incomplete bundle is never reported complete.
.PARAMETER SkipXsn
    Explicit XML/metadata-only export; skips steps 3-5. Optional, default off.
.PARAMETER Overwrite
    Forwarded to library-file download. On resume, replace already saved files.
.PARAMETER PageSize
    Number of items per library/reference query, default 500.
.PARAMETER IncludeFolders
    Include folder items in metadata; downloads always retain folder hierarchy.
.PARAMETER UserSiteUrl
    Optional owning-web override forwarded to dependency discovery.
.EXAMPLE
    .\RunInfoPathPipeline_SPS.ps1 `
      -SiteUrl "http://sharepoint/sites/Finance" `
      -LibraryName "Applications" `
      -OutputRootDirectory "D:\SPExports"
.EXAMPLE
    .\RunInfoPathPipeline_SPS.ps1 `
      -ResumeExportDirectory "D:\SPExports\InfoPathExport_20261027_191800_12345678"

    Resume from the first failed/unfinished step, reusing the same bundle.
.EXAMPLE
    .\RunInfoPathPipeline_SPS.ps1 `
      -ResumeExportDirectory "D:\SPExports\InfoPathExport_20261027_191800_12345678" `
      -StartStep 4

    Retry discovery and reference-data export without redownloading forms/XSNs.
.EXAMPLE
    .\RunInfoPathPipeline_SPS.ps1 `
      -SiteUrl "http://sharepoint/sites/Finance" `
      -LibraryName "Applications" `
      -OutputRootDirectory "D:\SPExports" -SkipXsn
#>
#requires -Version 5.1
[CmdletBinding(DefaultParameterSetName = 'New')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'New')][ValidateNotNullOrEmpty()][string]$SiteUrl,
    [Parameter(Mandatory = $true, ParameterSetName = 'New')][ValidateNotNullOrEmpty()][string]$LibraryName,
    [Parameter(Mandatory = $true, ParameterSetName = 'New')][ValidateNotNullOrEmpty()][string]$OutputRootDirectory,
    [Parameter(Mandatory = $true, ParameterSetName = 'Resume')][ValidateNotNullOrEmpty()][string]$ResumeExportDirectory,
    [ValidateRange(1, 5)][int]$StartStep = 1,
    [ValidateRange(1, 10000)][int]$PageSize = 500,
    [switch]$IncludeFolders,
    [switch]$Overwrite,
    [switch]$SkipXsn,
    [string]$UserSiteUrl
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$logEntries = [System.Collections.Generic.List[string]]::new()
$stepDefinitions = @(
    @{ Name = 'Download library files'; Script = 'MoveLibraryFilesToLocal_SPS.ps1' }
    @{ Name = 'Export library metadata'; Script = 'GetMetadataforLib_SPS.ps1' }
    @{ Name = 'Download XSN versions'; Script = 'GetLibraryXsnVersions_SPS.ps1' }
    @{ Name = 'Discover XSN reference dependencies'; Script = 'GetXsnReferenceDependencies.ps1' }
    @{ Name = 'Export lookup lists and site users'; Script = 'GetXsnReferenceData_SPS.ps1' }
)

function Write-SpsPipelineStage {
    param([string]$Message, [ValidateSet('Init', 'Run', 'Done', 'Warn', 'Error')][string]$Group = 'Run')
    $color = switch ($Group) { 'Init' { 'Cyan' } 'Done' { 'Green' } 'Warn' { 'DarkYellow' } 'Error' { 'Red' } default { 'White' } }
    Write-Host "[$Group] $Message" -ForegroundColor $color
    $logEntries.Add(('[{0}] [{1}] {2}' -f [datetime]::UtcNow.ToString('o'), $Group, $Message))
}

function Get-SpsBundlePath {
    param([string]$RelativePath)
    if ([string]::IsNullOrWhiteSpace($RelativePath) -or [IO.Path]::IsPathRooted($RelativePath) -or
        $RelativePath.Contains(':') -or '..' -in ($RelativePath -split '[\\/]')) {
        throw "Expected a bundle-relative path: '$RelativePath'."
    }
    $path = [IO.Path]::GetFullPath((Join-Path $bundleRoot $RelativePath))
    if (-not $path.StartsWith($bundleRoot.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Export path escapes its bundle: '$RelativePath'."
    }
    return $path
}

function Get-SpsBundleRelativePath {
    param([string]$Path)
    $prefix = $bundleRoot.TrimEnd('\') + '\'
    $fullPath = [IO.Path]::GetFullPath($Path)
    if (-not $fullPath.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Child export is outside its bundle: '$Path'."
    }
    return $fullPath.Substring($prefix.Length)
}

function Save-SpsExportManifest {
    $manifest.UpdatedAtUtc = [datetime]::UtcNow.ToString('o')
    $temporaryManifest = $manifestPath + '.tmp'
    try {
        $manifest | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $temporaryManifest -Encoding UTF8
        if (Test-Path -LiteralPath $manifestPath -PathType Leaf) { [IO.File]::Replace($temporaryManifest, $manifestPath, [NullString]::Value) }
        else { [IO.File]::Move($temporaryManifest, $manifestPath) }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryManifest -PathType Leaf) { Remove-Item -LiteralPath $temporaryManifest }
    }
    $logEntries | Set-Content -LiteralPath $logPath -Encoding UTF8
}

function Get-SpsNewInventory {
    param([string]$Directory, [string]$FileName, [string[]]$PreviousPaths)
    $newFiles = @(Get-ChildItem -LiteralPath $Directory -Filter $FileName -File -Recurse |
        Where-Object { $PreviousPaths -notcontains $_.FullName })
    if ($newFiles.Count -gt 1) { throw "Child script produced multiple new '$FileName' files under '$Directory'; refusing to guess." }
    if ($newFiles.Count -eq 1) { return $newFiles[0].FullName }
    return ''
}

if ($PSCmdlet.ParameterSetName -eq 'New') {
    if ($StartStep -ne 1) { throw 'New exports must start at step 1. Use ResumeExportDirectory for later steps.' }
    $uri = $null
    if (-not [uri]::TryCreate($SiteUrl, [UriKind]::Absolute, [ref]$uri) -or
        $uri.Scheme -notin @('http', 'https') -or $uri.Query -or $uri.Fragment -or $uri.UserInfo) {
        throw 'SiteUrl must be an absolute HTTP/HTTPS web URL without credentials, query, or fragment.'
    }
    $SiteUrl = $uri.AbsoluteUri.TrimEnd('/')
    $name = 'InfoPathExport_{0}_{1}' -f (Get-Date -Format 'yyyyMMdd_HHmmss'), [guid]::NewGuid().ToString('N').Substring(0, 8)
    $bundleRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $OutputRootDirectory $name))
    New-Item -Path $bundleRoot -ItemType Directory | Out-Null
    $manifest = [pscustomobject][ordered]@{
        SchemaVersion = 1; SiteUrl = $SiteUrl; LibraryName = $LibraryName
        CreatedAtUtc = [datetime]::UtcNow.ToString('o'); UpdatedAtUtc = ''; Status = 'Running'
        Options = [pscustomobject]@{
            PageSize = $PageSize; IncludeFolders = $IncludeFolders.IsPresent; Overwrite = $Overwrite.IsPresent
            SkipXsn = $SkipXsn.IsPresent; UserSiteUrl = $UserSiteUrl
        }
        Paths = [pscustomobject]@{ XmlDirectory = 'Xml'; MetadataCsv = ''; XsnInventoryCsv = ''; ReferenceDependencies = ''; ReferenceDataDirectory = '' }
        Steps = @(
            for ($index = 0; $index -lt $stepDefinitions.Count; $index++) {
                [pscustomobject]@{
                    Number = $index + 1; Name = $stepDefinitions[$index].Name; Script = $stepDefinitions[$index].Script
                    Status = 'Pending'; StartedAtUtc = ''; CompletedAtUtc = ''; Error = ''
                }
            }
        )
    }
}
else {
    $bundleRoot = (Resolve-Path -LiteralPath $ResumeExportDirectory).ProviderPath
    $resumeManifest = Join-Path $bundleRoot 'ExportManifest.json'
    if (-not (Test-Path -LiteralPath $resumeManifest -PathType Leaf)) { throw "Export manifest not found: $resumeManifest" }
    $manifest = Get-Content -LiteralPath $resumeManifest -Raw | ConvertFrom-Json
    if ($manifest.SchemaVersion -ne 1 -or $manifest.Steps -isnot [array] -or $manifest.Steps.Count -ne 5) { throw 'Unsupported SPS export manifest.' }
    for ($index = 0; $index -lt 5; $index++) {
        if ($manifest.Steps[$index].Number -ne ($index + 1) -or $manifest.Steps[$index].Script -ne $stepDefinitions[$index].Script) {
            throw 'SPS export manifest has an invalid step definition.'
        }
    }
    $SiteUrl = [string]$manifest.SiteUrl
    $LibraryName = [string]$manifest.LibraryName
    foreach ($option in @('PageSize', 'IncludeFolders', 'Overwrite', 'SkipXsn', 'UserSiteUrl')) {
        if (-not $PSBoundParameters.ContainsKey($option)) { Set-Variable -Name $option -Value $manifest.Options.$option }
    }
    $IncludeFolders = [switch][bool]$IncludeFolders
    $Overwrite = [switch][bool]$Overwrite
    $SkipXsn = [switch][bool]$SkipXsn
    $manifest.Options.PageSize = $PageSize
    $manifest.Options.IncludeFolders = $IncludeFolders.IsPresent
    $manifest.Options.Overwrite = $Overwrite.IsPresent
    $manifest.Options.SkipXsn = $SkipXsn.IsPresent
    $manifest.Options.UserSiteUrl = $UserSiteUrl
    if (-not $PSBoundParameters.ContainsKey('StartStep')) {
        $unfinished = @($manifest.Steps | Where-Object Status -notin @('Completed', 'Skipped'))
        if ($unfinished.Count -eq 0) { throw 'All steps already completed. Specify StartStep to explicitly rerun steps.' }
        $StartStep = $unfinished[0].Number
    }
}

$manifestPath = Join-Path $bundleRoot 'ExportManifest.json'
$logPath = Join-Path $bundleRoot 'Export.log'
if (Test-Path -LiteralPath $logPath -PathType Leaf) {
    foreach ($line in Get-Content -LiteralPath $logPath) { $logEntries.Add($line) }
}
$lock = [IO.File]::Open((Join-Path $bundleRoot 'Export.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
try {
    $manifest.Status = 'Running'
    foreach ($step in @($manifest.Steps | Where-Object Number -ge $StartStep)) {
        $step.Status = 'Pending'; $step.Error = ''; $step.StartedAtUtc = ''; $step.CompletedAtUtc = ''
    }
    foreach ($directory in @('Xml', 'Metadata', 'XSNTemplates', 'Dependencies', 'ReferenceData')) {
        New-Item -Path (Get-SpsBundlePath $directory) -ItemType Directory -Force | Out-Null
    }
    Write-SpsPipelineStage -Group Init -Message "SPS export bundle: $bundleRoot"
    Write-SpsPipelineStage -Group Init -Message "Site: $SiteUrl; library: $LibraryName; starting step: $StartStep"
    foreach ($step in $manifest.Steps) { Write-Host ("  {0}) {1} ({2})" -f $step.Number, $step.Name, $step.Script) }
    Save-SpsExportManifest
    foreach ($step in @($manifest.Steps | Where-Object Number -ge $StartStep)) {
        $step.StartedAtUtc = [datetime]::UtcNow.ToString('o')
        $step.Status = 'Running'
        switch ($step.Number) {
            2 { $manifest.Paths.MetadataCsv = '' }
            3 {
                $manifest.Paths.XsnInventoryCsv = ''
                $manifest.Paths.ReferenceDependencies = ''
                $manifest.Paths.ReferenceDataDirectory = ''
            }
            4 {
                $manifest.Paths.ReferenceDependencies = ''
                $manifest.Paths.ReferenceDataDirectory = ''
            }
            5 { $manifest.Paths.ReferenceDataDirectory = '' }
        }
        Save-SpsExportManifest
        try {
            if ($SkipXsn -and $step.Number -ge 3) {
                $step.Status = 'Skipped'
                Write-SpsPipelineStage -Group Warn -Message "Step $($step.Number)/5 explicitly skipped by SkipXsn."
                continue
            }
            $scriptPath = Join-Path $PSScriptRoot $step.Script
            if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) { throw "Required child script not found: $scriptPath" }
            Write-SpsPipelineStage -Message "Step $($step.Number)/5: $($step.Name) ($($step.Script))"
            switch ($step.Number) {
                1 {
                    $null = & $scriptPath -SiteUrl $SiteUrl -LibraryName $LibraryName -LocalPath (Get-SpsBundlePath 'Xml') -PageSize $PageSize -Overwrite:$Overwrite
                }
                2 {
                    $metadataPath = Get-SpsBundlePath 'Metadata\LibraryMetadata.csv'
                    $null = & $scriptPath -SiteUrl $SiteUrl -LibraryName $LibraryName -OutputCsvPath $metadataPath -PageSize $PageSize -IncludeFolders:$IncludeFolders
                    if (-not (Test-Path -LiteralPath $metadataPath -PathType Leaf)) { throw 'Metadata export did not create its CSV.' }
                    $manifest.Paths.MetadataCsv = Get-SpsBundleRelativePath $metadataPath
                }
                3 {
                    $xsnRoot = Get-SpsBundlePath 'XSNTemplates'
                    $previous = @(Get-ChildItem -LiteralPath $xsnRoot -Filter 'XsnTemplateDownloads.csv' -File -Recurse | Select-Object -ExpandProperty FullName)
                    try { $null = & $scriptPath -SiteUrl $SiteUrl -LibraryName $LibraryName -LocalPath $xsnRoot }
                    finally {
                        $inventory = Get-SpsNewInventory $xsnRoot 'XsnTemplateDownloads.csv' $previous
                        if ($inventory) { $manifest.Paths.XsnInventoryCsv = Get-SpsBundleRelativePath $inventory }
                    }
                    if (-not $manifest.Paths.XsnInventoryCsv) { throw 'XSN export did not create a download inventory.' }
                }
                4 {
                    if (-not $manifest.Paths.XsnInventoryCsv -or -not (Test-Path -LiteralPath (Get-SpsBundlePath $manifest.Paths.XsnInventoryCsv) -PathType Leaf)) {
                        $step.Status = 'Blocked'; throw 'Dependency discovery requires an XSN download inventory from step 3.'
                    }
                    $dependencyPath = Get-SpsBundlePath ('Dependencies\ReferenceDependencies_{0}.json' -f [guid]::NewGuid().ToString('N'))
                    try {
                        $null = & $scriptPath -XsnExportCsvPath (Get-SpsBundlePath $manifest.Paths.XsnInventoryCsv) -OutputPath $dependencyPath -UserSiteUrl $UserSiteUrl
                    }
                    finally {
                        if (Test-Path -LiteralPath $dependencyPath -PathType Leaf) { $manifest.Paths.ReferenceDependencies = Get-SpsBundleRelativePath $dependencyPath }
                    }
                    if (-not $manifest.Paths.ReferenceDependencies) { throw 'Dependency discovery did not create its JSON report.' }
                }
                5 {
                    if (-not $manifest.Paths.ReferenceDependencies -or -not (Test-Path -LiteralPath (Get-SpsBundlePath $manifest.Paths.ReferenceDependencies) -PathType Leaf)) {
                        $step.Status = 'Blocked'; throw 'Reference export requires dependency JSON from step 4.'
                    }
                    $dependencies = Get-Content -LiteralPath (Get-SpsBundlePath $manifest.Paths.ReferenceDependencies) -Raw | ConvertFrom-Json
                    if ($dependencies.Templates -isnot [array] -or $dependencies.Templates.Count -eq 0) {
                        $step.Status = 'Blocked'; throw 'Dependency discovery produced no usable templates; reference export cannot proceed.'
                    }
                    $referenceRoot = Get-SpsBundlePath 'ReferenceData'
                    $previous = @(Get-ChildItem -LiteralPath $referenceRoot -Filter 'Manifest.json' -File -Recurse | Select-Object -ExpandProperty FullName)
                    try { $null = & $scriptPath -DependencyManifestPath (Get-SpsBundlePath $manifest.Paths.ReferenceDependencies) -LocalPath $referenceRoot -PageSize $PageSize }
                    finally {
                        $referenceManifest = Get-SpsNewInventory $referenceRoot 'Manifest.json' $previous
                        if ($referenceManifest) { $manifest.Paths.ReferenceDataDirectory = Get-SpsBundleRelativePath (Split-Path $referenceManifest -Parent) }
                    }
                    if (-not $manifest.Paths.ReferenceDataDirectory) { throw 'Reference export did not create a snapshot manifest.' }
                }
            }
            $step.Status = 'Completed'
            Write-SpsPipelineStage -Group Done -Message "Step $($step.Number)/5 completed."
        }
        catch {
            if ($step.Status -ne 'Blocked') { $step.Status = 'Failed' }
            $step.Error = $_.Exception.Message
            Write-SpsPipelineStage -Group Warn -Message "Step $($step.Number)/5 $($step.Status): $($step.Error) Continuing available exports."
        }
        finally {
            $step.CompletedAtUtc = [datetime]::UtcNow.ToString('o')
            Save-SpsExportManifest
        }
    }
    $incomplete = @($manifest.Steps | Where-Object Status -notin @('Completed', 'Skipped'))
    $manifest.Status = if ($incomplete.Count -eq 0) { 'Complete' } else { 'Incomplete' }
    Write-SpsPipelineStage -Group $(if ($incomplete.Count) { 'Warn' } else { 'Done' }) -Message "Bundle status: $($manifest.Status). Copy the entire folder: $bundleRoot"
    Save-SpsExportManifest
    if ($incomplete.Count -gt 0) { throw "SPS export incomplete: $($incomplete.Count) failed/blocked/unfinished step(s). Review '$manifestPath' and '$logPath'; partial exports are retained." }
    Write-Output $bundleRoot
}
finally {
    try {
        if ($manifest.Status -eq 'Running') {
            $manifest.Status = 'Incomplete'
            Save-SpsExportManifest
        }
    }
    finally { $lock.Dispose() }
}
