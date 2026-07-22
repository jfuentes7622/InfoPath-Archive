<#
.SYNOPSIS
    Runs the local InfoPath processing pipeline end-to-end.

.DESCRIPTION
    Orchestrates the non-SharePoint child scripts in this folder in dependency order.
    This intentionally excludes SharePoint instance scripts such as:
      - MoveLibraryFilesToLocal_SPS.ps1
      - GetMetadataforLib_SPS.ps1

    Pipeline order:
      1) ClassifyInfoPathVersions.ps1
      2) ProcessInfoPathAnalyze.ps1
      3) ProcessInfoPathProperties.ps1
      4) AnalyzeInfoPathGroupingFromCsv.ps1
      5) BuildXsltFromGroupingCsv.ps1
      6) ProcessInfopathToHTML.ps1

.PARAMETER SourceDirectory
    Root directory containing InfoPath XML files.

.PARAMETER OutputRootDirectory
    Root folder for all pipeline outputs.

.PARAMETER Recurse
    Include subfolders when searching for XML files.

.PARAMETER MaxFilesToAnalyze
    Max sample files per version for ProcessInfoPathAnalyze.ps1.
    Default: 5

.PARAMETER PrintToPdf
    Pass-through value for ProcessInfopathToHTML.ps1.
    Accepted values: yes, no, true, false, 1, 0.
    Default: false

.EXAMPLE
    .\RunInfoPathPipeline.ps1 `
      -SourceDirectory "C:\Data\InfoPath" `
      -OutputRootDirectory "C:\Data\PipelineOutput" `
      -Recurse

.EXAMPLE
    .\RunInfoPathPipeline.ps1 `
      -SourceDirectory "C:\Data\InfoPath" `
      -OutputRootDirectory "C:\Data\PipelineOutput" `
      -PrintToPdf "true"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SourceDirectory,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputRootDirectory,

    [switch]$Recurse,

    [ValidateRange(1, 100)]
    [int]$MaxFilesToAnalyze = 5,

    [string]$PrintToPdf = "false"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Stage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("Init", "Input", "Run", "Done", "Warn", "Error")]
        [string]$Group = "Run"
    )

    $color = switch ($Group) {
        "Init" { "Cyan" }
        "Input" { "DarkCyan" }
        "Run" { "White" }
        "Done" { "Green" }
        "Warn" { "DarkYellow" }
        "Error" { "Red" }
        default { "White" }
    }

    Write-Host "[$Group] $Message" -ForegroundColor $color
}

function Assert-ScriptExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath
    )

    if (-not (Test-Path -LiteralPath $ScriptPath -PathType Leaf)) {
        throw "Required child script not found: $ScriptPath"
    }
}

if (-not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) {
    throw "SourceDirectory does not exist or is not a folder: $SourceDirectory"
}

if (-not (Test-Path -LiteralPath $OutputRootDirectory -PathType Container)) {
    New-Item -Path $OutputRootDirectory -ItemType Directory -Force | Out-Null
}

$sourceRoot = (Resolve-Path -LiteralPath $SourceDirectory).Path
$outputRoot = (Resolve-Path -LiteralPath $OutputRootDirectory).Path

$reportsDir = Join-Path -Path $outputRoot -ChildPath "Reports"
$analysisDir = Join-Path -Path $outputRoot -ChildPath "Analysis"
$templatesDir = Join-Path -Path $outputRoot -ChildPath "Templates"
$htmlDir = Join-Path -Path $outputRoot -ChildPath "Html"

foreach ($dir in @($reportsDir, $analysisDir, $templatesDir, $htmlDir)) {
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
}

$versionCsvPath = Join-Path -Path $reportsDir -ChildPath "InfoPathVersionMap.csv"
$groupingDetailedPath = Join-Path -Path $reportsDir -ChildPath "InfoPathGroupingDetailed.csv"
$groupingSummaryPath = Join-Path -Path $reportsDir -ChildPath "InfoPathGroupingSummary.csv"
$templateMappingCsvPath = Join-Path -Path $templatesDir -ChildPath "InfoPathXsltGroupMap.csv"

$classifyScript = Join-Path -Path $PSScriptRoot -ChildPath "ClassifyInfoPathVersions.ps1"
$analyzeScript = Join-Path -Path $PSScriptRoot -ChildPath "ProcessInfoPathAnalyze.ps1"
$propertiesScript = Join-Path -Path $PSScriptRoot -ChildPath "ProcessInfoPathProperties.ps1"
$groupingScript = Join-Path -Path $PSScriptRoot -ChildPath "AnalyzeInfoPathGroupingFromCsv.ps1"
$buildXsltScript = Join-Path -Path $PSScriptRoot -ChildPath "BuildXsltFromGroupingCsv.ps1"
$htmlScript = Join-Path -Path $PSScriptRoot -ChildPath "ProcessInfopathToHTML.ps1"

foreach ($script in @($classifyScript, $analyzeScript, $propertiesScript, $groupingScript, $buildXsltScript, $htmlScript)) {
    Assert-ScriptExists -ScriptPath $script
}

Write-Stage -Group Init -Message "Starting InfoPath local pipeline orchestration."
Write-Stage -Group Input -Message "Source directory: $sourceRoot"
Write-Stage -Group Input -Message "Output root: $outputRoot"
Write-Stage -Group Input -Message "Reports directory: $reportsDir"
Write-Stage -Group Input -Message "Analysis directory: $analysisDir"
Write-Stage -Group Input -Message "Templates directory: $templatesDir"
Write-Stage -Group Input -Message "HTML directory: $htmlDir"

$startTime = Get-Date

try {
    Write-Stage -Group Run -Message "Step 1/6: Classify versions"
    & $classifyScript -SourceDirectory $sourceRoot -OutputCsvPath $versionCsvPath -Recurse:$Recurse.IsPresent

    Write-Stage -Group Run -Message "Step 2/6: Analyze structure by version"
    & $analyzeScript -SourceDirectory $sourceRoot -OutputDirectory $analysisDir -VersioningCsvPath $versionCsvPath -MaxFilesToAnalyze $MaxFilesToAnalyze -Recurse:$Recurse.IsPresent

    Write-Stage -Group Run -Message "Step 3/6: Extract normalized properties"
    & $propertiesScript -SourceDirectory $sourceRoot -OutputDirectory $reportsDir -VersioningCsvPath $versionCsvPath -AnalysisDirectory $analysisDir -Recurse:$Recurse.IsPresent

    Write-Stage -Group Run -Message "Step 4/6: Analyze grouping from version CSV"
    & $groupingScript -InputCsvPath $versionCsvPath -OutputCsvPath $groupingDetailedPath -SummaryCsvPath $groupingSummaryPath

    Write-Stage -Group Run -Message "Step 5/6: Build XSLT templates per group"
    & $buildXsltScript -GroupingCsvPath $groupingDetailedPath -TemplateOutputDirectory $templatesDir -MappingCsvPath $templateMappingCsvPath

    Write-Stage -Group Run -Message "Step 6/6: Transform InfoPath XML to HTML"
    & $htmlScript -InfoPathDirectory $sourceRoot -MetadataCsvPath $versionCsvPath -GroupingCsvPath $groupingDetailedPath -TemplateMappingCsvPath $templateMappingCsvPath -OutputDirectory $htmlDir -PrintToPdf $PrintToPdf -Recurse:$Recurse.IsPresent

    $duration = (Get-Date) - $startTime
    Write-Stage -Group Done -Message "Pipeline completed successfully in $($duration.ToString())."
    Write-Host "Version CSV         : $versionCsvPath" -ForegroundColor Magenta
    Write-Host "Grouping CSV        : $groupingDetailedPath" -ForegroundColor Magenta
    Write-Host "Grouping Summary    : $groupingSummaryPath" -ForegroundColor Magenta
    Write-Host "Template Mapping    : $templateMappingCsvPath" -ForegroundColor Magenta
    Write-Host "Normalized CSV root : $(Join-Path -Path $reportsDir -ChildPath 'Normalized')" -ForegroundColor Magenta
    Write-Host "HTML output root    : $htmlDir" -ForegroundColor Magenta
}
catch {
    Write-Stage -Group Error -Message "Pipeline failed: $($_.Exception.Message)"
    throw
}
