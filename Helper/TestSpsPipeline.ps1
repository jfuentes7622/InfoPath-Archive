<#
.SYNOPSIS
    Tests portable SPS bundles, step orchestration, partial failures, and resume.
.DESCRIPTION
    Uses isolated child-script stubs. Never connects to SharePoint.
    Also tests actual downloader failure reporting with an in-process facade.
#>
#requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repository = Split-Path $PSScriptRoot -Parent
$work = Join-Path ([IO.Path]::GetTempPath()) ('InfoPathSpsPipelineTest-' + [guid]::NewGuid().ToString('N'))

function Assert-SpsPipeline {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Set-SpsScenario {
    param([int[]]$FailSteps = @(), [bool]$NoXsnInventory = $false)
    @{ FailSteps = @($FailSteps); NoXsnInventory = $NoXsnInventory } |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $tools 'Scenario.json') -Encoding UTF8
    if (Test-Path -LiteralPath $callsPath) { Remove-Item -LiteralPath $callsPath }
}

function Read-SpsCalls {
    return @(Get-Content -LiteralPath $callsPath | ForEach-Object { $_ | ConvertFrom-Json })
}

function Invoke-SpsTestExport {
    param([string]$Name, [bool]$ExpectFailure = $false, [switch]$SkipXsn)
    $parent = Join-Path $work $Name
    $failed = $false
    try {
        $null = & $pipeline -SiteUrl 'https://example.test/sites/owner/red/' -LibraryName 'Applications' `
            -OutputRootDirectory $parent -PageSize 1000 -IncludeFolders -Overwrite -SkipXsn:$SkipXsn `
            -UserSiteUrl 'https://example.test/sites/owner/red' 6>$null
    }
    catch {
        if (-not $ExpectFailure -or $_.Exception.Message -notlike '*SPS export incomplete*') { throw }
        $failed = $true
    }
    Assert-SpsPipeline ($failed -eq $ExpectFailure) "Incorrect failure outcome for $Name."
    $folders = @(Get-ChildItem -LiteralPath $parent -Directory)
    Assert-SpsPipeline ($folders.Count -eq 1) 'Expected exactly one timestamped bundle.'
    return $folders[0].FullName
}

try {
    $tools = Join-Path $work 'Tools'
    New-Item $tools -ItemType Directory -Force | Out-Null
    $pipeline = Join-Path $tools 'RunInfoPathPipeline_SPS.ps1'
    Copy-Item (Join-Path $repository 'RunInfoPathPipeline_SPS.ps1') $pipeline
    $callsPath = Join-Path $tools 'Calls.jsonl'
    $scriptNames = @('MoveLibraryFilesToLocal_SPS.ps1', 'GetMetadataforLib_SPS.ps1', 'GetLibraryXsnVersions_SPS.ps1', 'GetXsnReferenceDependencies.ps1', 'GetXsnReferenceData_SPS.ps1')
    $stub = @'
[CmdletBinding()]
param(
 [string]$SiteUrl, [string]$LibraryName, [string]$LocalPath, [string]$OutputCsvPath,
 [int]$PageSize, [switch]$Overwrite, [switch]$IncludeFolders, [string]$XsnExportCsvPath,
 [string]$OutputPath, [string]$DependencyManifestPath, [string]$UserSiteUrl
)
$names = @('MoveLibraryFilesToLocal_SPS.ps1', 'GetMetadataforLib_SPS.ps1', 'GetLibraryXsnVersions_SPS.ps1', 'GetXsnReferenceDependencies.ps1', 'GetXsnReferenceData_SPS.ps1')
$number = [array]::IndexOf($names, (Split-Path $PSCommandPath -Leaf)) + 1
@{ Number = $number; Parameters = $PSBoundParameters } | ConvertTo-Json -Depth 5 -Compress |
 Add-Content -LiteralPath (Join-Path $PSScriptRoot 'Calls.jsonl')
$scenario = Get-Content (Join-Path $PSScriptRoot 'Scenario.json') -Raw | ConvertFrom-Json
switch ($number) {
 1 { '<Form>Offline</Form>' | Set-Content (Join-Path $LocalPath 'Sample.xml') -Encoding UTF8 }
 2 { [pscustomobject]@{ FileName = 'Sample.xml'; ItemId = 1 } | Export-Csv $OutputCsvPath -NoTypeInformation -Encoding UTF8 }
 3 {
  if (-not $scenario.NoXsnInventory) {
   $root = Join-Path $LocalPath ('XsnTemplates_' + [guid]::NewGuid().ToString('N'))
   New-Item $root -ItemType Directory | Out-Null
   'CAB fixture' | Set-Content (Join-Path $root 'template.xsn') -Encoding UTF8
   [pscustomobject]@{ SiteUrl=$SiteUrl; LibraryName=$LibraryName; TemplateUrl=($SiteUrl + '/Applications/Forms/template.xsn'); ConfiguredBy='Library'; SharePointVersion='1.0'; IsCurrentVersion=$true; Modified='2026-10-27'; XsnRelativePath='template.xsn'; Status='Success'; Error='' } |
    Export-Csv (Join-Path $root 'XsnTemplateDownloads.csv') -NoTypeInformation -Encoding UTF8
  }
 }
 4 {
  if (-not (Test-Path -LiteralPath $XsnExportCsvPath)) { throw 'Missing download inventory.' }
  @{ SchemaVersion=1; Templates=@(@{ SolutionName='urn:test' }); Lists=@(); UserSites=@($UserSiteUrl); Warnings=@() } |
   ConvertTo-Json -Depth 10 | Set-Content $OutputPath -Encoding UTF8
 }
 5 {
  if (-not (Test-Path -LiteralPath $DependencyManifestPath)) { throw 'Missing dependencies.' }
  $root = Join-Path $LocalPath ('XsnReferences_' + [guid]::NewGuid().ToString('N'))
  New-Item $root -ItemType Directory | Out-Null
  @{ SchemaVersion=1; Dependencies=(Get-Content $DependencyManifestPath -Raw | ConvertFrom-Json); Lists=@(); Users=@() } |
   ConvertTo-Json -Depth 20 | Set-Content (Join-Path $root 'Manifest.json') -Encoding UTF8
  Write-Output $root
 }
}
if ($scenario.FailSteps -contains $number) { throw "Synthetic failure in step $number." }
'@
    foreach ($name in $scriptNames) { $stub | Set-Content (Join-Path $tools $name) -Encoding UTF8 }
    Set-SpsScenario
    $root = Invoke-SpsTestExport 'Complete'
    $manifest = Get-Content (Join-Path $root 'ExportManifest.json') -Raw | ConvertFrom-Json
    $calls = @(Read-SpsCalls)
    Assert-SpsPipeline (($calls.Number -join ',') -ceq '1,2,3,4,5') 'Steps were not invoked in order.'
    Assert-SpsPipeline ($manifest.Status -ceq 'Complete' -and @($manifest.Steps | Where-Object Status -ne 'Completed').Count -eq 0) 'Successful bundle was not marked complete.'
    Assert-SpsPipeline (-not (Test-Path -LiteralPath (Join-Path $root 'ExportManifest.json.tmp'))) 'Temporary manifest was not cleaned after saving.'
    Assert-SpsPipeline ($calls[0].Parameters.PageSize -eq 1000 -and $calls[0].Parameters.Overwrite.IsPresent -and $calls[1].Parameters.IncludeFolders.IsPresent) 'Download/metadata options were not forwarded.'
    Assert-SpsPipeline ($calls[3].Parameters.UserSiteUrl -ceq 'https://example.test/sites/owner/red') 'User-site override was not forwarded.'
    foreach ($property in $manifest.Paths.PSObject.Properties) {
        Assert-SpsPipeline (-not [IO.Path]::IsPathRooted($property.Value) -and (Test-Path -LiteralPath (Join-Path $root $property.Value))) "Nonportable or missing bundle output: $($property.Name)."
    }
    $initialXsn = $manifest.Paths.XsnInventoryCsv
    $initialReference = $manifest.Paths.ReferenceDataDirectory
    Set-SpsScenario
    $resumed = & $pipeline -ResumeExportDirectory $root -StartStep 4 6>$null
    $resumedManifest = Get-Content (Join-Path $root 'ExportManifest.json') -Raw | ConvertFrom-Json
    Assert-SpsPipeline ($resumed -eq $root -and ((@(Read-SpsCalls).Number -join ',') -ceq '4,5')) 'Explicit resume did not reuse the bundle or skip earlier steps.'
    Assert-SpsPipeline ($resumedManifest.Paths.XsnInventoryCsv -eq $initialXsn -and $resumedManifest.Paths.ReferenceDataDirectory -ne $initialReference) 'Resume selected stale artifacts or redownloaded templates.'
    Assert-SpsPipeline ($resumedManifest.Options.IncludeFolders -is [bool] -and $resumedManifest.Options.Overwrite -is [bool]) 'Resume serialized switches as objects instead of booleans.'
    $relocated = Join-Path $work 'RelocatedBundle'
    Copy-Item -LiteralPath $root -Destination $relocated -Recurse
    Set-SpsScenario
    $null = & $pipeline -ResumeExportDirectory $relocated -StartStep 4 -Overwrite:$false 6>$null
    $relocatedManifest = Get-Content (Join-Path $relocated 'ExportManifest.json') -Raw | ConvertFrom-Json
    Assert-SpsPipeline ($relocatedManifest.Status -ceq 'Complete' -and $relocatedManifest.Options.Overwrite -eq $false) 'Relocated bundle could not resume or option override was lost.'
    Set-SpsScenario
    $beforeLockHash = (Get-FileHash (Join-Path $root 'ExportManifest.json')).Hash
    $heldLock = [IO.File]::Open((Join-Path $root 'Export.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $lockRejected = $false
    try {
        try { $null = & $pipeline -ResumeExportDirectory $root -StartStep 5 6>$null }
        catch { $lockRejected = $true }
    }
    finally { $heldLock.Dispose() }
    Assert-SpsPipeline ($lockRejected -and -not (Test-Path -LiteralPath $callsPath) -and (Get-FileHash (Join-Path $root 'ExportManifest.json')).Hash -eq $beforeLockHash) 'Concurrent resume changed the active bundle or ran child scripts.'
    $completedResumeRejected = $false
    try { $null = & $pipeline -ResumeExportDirectory $root 6>$null }
    catch { $completedResumeRejected = $_.Exception.Message -like '*All steps already completed*' }
    Assert-SpsPipeline $completedResumeRejected 'Completed bundle silently started another export.'
    $invalidStartRejected = $false
    try {
        $null = & $pipeline -SiteUrl 'https://example.test' -LibraryName 'Applications' -OutputRootDirectory (Join-Path $work 'InvalidStart') -StartStep 3 6>$null
    }
    catch { $invalidStartRejected = $_.Exception.Message -like '*New exports must start at step 1*' }
    Assert-SpsPipeline $invalidStartRejected 'Fresh export accepted a start step without prerequisites.'
    Set-SpsScenario -FailSteps @(3) -NoXsnInventory $true
    $retryFailed = $false
    try { $null = & $pipeline -ResumeExportDirectory $relocated -StartStep 3 6>$null }
    catch {
        if ($_.Exception.Message -notlike '*SPS export incomplete*') { throw }
        $retryFailed = $true
    }
    $failedRetryManifest = Get-Content (Join-Path $relocated 'ExportManifest.json') -Raw | ConvertFrom-Json
    Assert-SpsPipeline ($retryFailed -and -not $failedRetryManifest.Paths.XsnInventoryCsv -and -not $failedRetryManifest.Paths.ReferenceDependencies -and -not $failedRetryManifest.Paths.ReferenceDataDirectory) 'Failed retry selected stale outputs from an earlier attempt.'
    Assert-SpsPipeline (Test-Path -LiteralPath (Join-Path $relocated $initialXsn)) 'Failed retry deleted previously collected artifacts.'

    Set-SpsScenario -FailSteps @(1, 3, 4)
    $partial = Invoke-SpsTestExport 'Partial' -ExpectFailure $true
    $partialManifest = Get-Content (Join-Path $partial 'ExportManifest.json') -Raw | ConvertFrom-Json
    Assert-SpsPipeline ((@(Read-SpsCalls).Number -join ',') -ceq '1,2,3,4,5') 'Failures prevented collecting remaining available data.'
    Assert-SpsPipeline ($partialManifest.Status -ceq 'Incomplete' -and $partialManifest.Steps[4].Status -ceq 'Completed') 'Partial dependencies were not used or incomplete bundle was reported as complete.'
    Assert-SpsPipeline (@($partialManifest.Steps | Where-Object Status -eq 'Failed').Count -eq 3) 'Failed steps were not retained in the manifest.'
    Set-SpsScenario
    $null = & $pipeline -ResumeExportDirectory $partial 6>$null
    Assert-SpsPipeline ((@(Read-SpsCalls).Number -join ',') -ceq '1,2,3,4,5') 'Automatic resume did not select the earliest failed step.'
    Assert-SpsPipeline ((Get-Content (Join-Path $partial 'ExportManifest.json') -Raw | ConvertFrom-Json).Status -ceq 'Complete') 'Successful retry did not complete the bundle.'

    Set-SpsScenario -FailSteps @(3) -NoXsnInventory $true
    $blocked = Invoke-SpsTestExport 'Blocked' -ExpectFailure $true
    $blockedManifest = Get-Content (Join-Path $blocked 'ExportManifest.json') -Raw | ConvertFrom-Json
    Assert-SpsPipeline ((@(Read-SpsCalls).Number -join ',') -ceq '1,2,3') 'Dependent scripts ran without their required input.'
    Assert-SpsPipeline ($blockedManifest.Steps[3].Status -ceq 'Blocked' -and $blockedManifest.Steps[4].Status -ceq 'Blocked') 'Missing prerequisite artifacts were not marked blocked.'
    Set-SpsScenario
    $stillIncomplete = $false
    try { $null = & $pipeline -ResumeExportDirectory $blocked -StartStep 4 6>$null }
    catch {
        if ($_.Exception.Message -notlike '*SPS export incomplete*') { throw }
        $stillIncomplete = $true
    }
    $afterSkippedFailure = Get-Content (Join-Path $blocked 'ExportManifest.json') -Raw | ConvertFrom-Json
    Assert-SpsPipeline ($stillIncomplete -and $afterSkippedFailure.Status -ceq 'Incomplete' -and $afterSkippedFailure.Steps[2].Status -ceq 'Failed') 'Resuming later steps erased an earlier failure.'

    Set-SpsScenario
    $skipped = Invoke-SpsTestExport 'XmlOnly' -SkipXsn
    $skippedManifest = Get-Content (Join-Path $skipped 'ExportManifest.json') -Raw | ConvertFrom-Json
    Assert-SpsPipeline ((@(Read-SpsCalls).Number -join ',') -ceq '1,2') 'XML-only export invoked XSN/reference scripts.'
    Assert-SpsPipeline ($skippedManifest.Status -ceq 'Complete' -and @($skippedManifest.Steps | Where-Object Status -eq 'Skipped').Count -eq 3) 'Intentional XML-only export was not recorded accurately.'
    Set-SpsScenario
    $null = & $pipeline -ResumeExportDirectory $skipped -StartStep 3 -SkipXsn:$false 6>$null
    $enabledManifest = Get-Content (Join-Path $skipped 'ExportManifest.json') -Raw | ConvertFrom-Json
    Assert-SpsPipeline ($enabledManifest.Status -ceq 'Complete' -and $enabledManifest.Options.SkipXsn -eq $false -and ((@(Read-SpsCalls).Number -join ',') -ceq '3,4,5')) 'Resume did not honor an explicit false switch override.'
    $second = & $pipeline -SiteUrl 'https://example.test/sites/owner/red' -LibraryName 'Applications' -OutputRootDirectory (Join-Path $work 'Complete') 6>$null
    Assert-SpsPipeline ($second -ne $root -and @(Get-ChildItem (Join-Path $work 'Complete') -Directory).Count -eq 2) 'New exports did not get separate timestamped bundles.'
    $metadataScript = Join-Path $tools 'GetMetadataforLib_SPS.ps1'
    Move-Item $metadataScript ($metadataScript + '.saved')
    try {
        Set-SpsScenario
        $missingChild = Invoke-SpsTestExport 'MissingChild' -ExpectFailure $true
        $missingManifest = Get-Content (Join-Path $missingChild 'ExportManifest.json') -Raw | ConvertFrom-Json
        Assert-SpsPipeline ($missingManifest.Steps[1].Status -ceq 'Failed' -and $missingManifest.Steps[4].Status -ceq 'Completed') 'Missing child script was not explicit or prevented independent exports.'
    }
    finally { Move-Item ($metadataScript + '.saved') $metadataScript }

    $downloadPage = [System.Collections.ArrayList]::new()
    $null = $downloadPage.Add(@{ ID = 1; FSObjType = 0; File = $null })
    Add-Member -InputObject $downloadPage -MemberType NoteProperty -Name ListItemCollectionPosition -Value $null
    $downloadList = [pscustomobject]@{ RootFolder = [pscustomobject]@{ ServerRelativeUrl = '/sites/owner/Applications' }; Page = $downloadPage }
    $downloadList | Add-Member ScriptMethod GetItems { param($Query) return ,$this.Page }
    $downloadLists = [pscustomobject]@{ List = $downloadList }
    $downloadLists | Add-Member ScriptMethod TryGetList { param($Name) return $this.List }
    $downloadWeb = [pscustomobject]@{ Title = 'Fixture'; Lists = $downloadLists; Disposed = $false }
    $downloadWeb | Add-Member ScriptMethod Dispose { $this.Disposed = $true }
    function Get-PSSnapin {
        [CmdletBinding()]param([switch]$Registered)
        return [pscustomobject]@{ Name = 'Microsoft.SharePoint.PowerShell' }
    }
    function Get-SPWeb {
        [CmdletBinding()]param([string]$Identity)
        return $downloadWeb
    }
    function New-Object {
        [CmdletBinding()]param([Parameter(Position = 0)][string]$TypeName)
        if ($TypeName -eq 'Microsoft.SharePoint.SPQuery') {
            return [pscustomobject]@{ RowLimit = 0; ViewAttributes = ''; ListItemCollectionPosition = $null }
        }
        return ,(Microsoft.PowerShell.Utility\New-Object -TypeName $TypeName)
    }
    $downloadRoot = Join-Path $work 'ActualPartialDownload'
    $downloadThrew = $false
    try {
        & (Join-Path $repository 'MoveLibraryFilesToLocal_SPS.ps1') -SiteUrl 'https://example.test/sites/owner' `
            -LibraryName 'Applications' -LocalPath $downloadRoot 6>$null
    }
    catch {
        if ($_.Exception.Message -notlike '*Library download incomplete: 1 error(s)*') { throw }
        $downloadThrew = $true
    }
    $downloadLog = @(Get-ChildItem $downloadRoot -Filter '*.log' -File)
    Assert-SpsPipeline ($downloadThrew -and $downloadWeb.Disposed -and $downloadLog.Count -eq 1) 'Actual downloader did not throw after preserving its partial-failure log and disposing the web.'
    Assert-SpsPipeline ((Get-Content $downloadLog[0].FullName -Raw).Contains('Errors=1')) 'Downloader failure details were not retained.'
    Write-Host 'PASS: SPS numbered steps, timestamped portable bundles, options, partial failures, blocked prerequisites, automatic/explicit resume, relocation, XML-only exports, and downloader error reporting.'
}
finally {
    if (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force }
}
