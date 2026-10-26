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

.PARAMETER StartStep
    First pipeline step to run (1 through 6). If omitted, show the numbered
    steps and prompt; Enter defaults to step 1. Later steps reuse existing
    prerequisite outputs from OutputRootDirectory.

.PARAMETER Recurse
    Include subfolders when searching for XML files.

.PARAMETER MaxFilesToAnalyze
    Max sample files per version for ProcessInfoPathAnalyze.ps1.
    Default: 5

.PARAMETER PrintToPdf
    Pass-through value for ProcessInfopathToHTML.ps1.
    Accepted values: yes, no, true, false, 1, 0.
    Default: false

.PARAMETER UseDisplayName
    Use readable labels derived from XML names in HTML. Set to $false for exact
    internal XML names. Structural identifiers are unchanged. Default: $true.

.PARAMETER DateFormat
    .NET format for date-only values displayed in HTML (and printed PDFs).
    Does not change source XML, normalized CSVs, or metadata. Default: MM/dd/yyyy.

.PARAMETER DateTimeFormat
    .NET format for date/time values in HTML. Stored clock times and offsets are
    retained. Default: MM/dd/yyyy HH:mm.

.PARAMETER FieldRulesPath
    Optional JSON field definitions and display-value mappings for HTML output.
    Passed to template generation and HTML rendering; CSV extraction is unchanged.

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

    [string]$PrintToPdf = "false",

    [bool]$UseDisplayName = $true,

    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        $null = [datetime]::Now.ToString($_, [System.Globalization.CultureInfo]::InvariantCulture)
        return $true
    })]
    [string]$DateFormat = "MM/dd/yyyy",

    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        $null = [datetimeoffset]::Now.ToString($_, [System.Globalization.CultureInfo]::InvariantCulture)
        return $true
    })]
    [string]$DateTimeFormat = "MM/dd/yyyy HH:mm",

    [string]$FieldRulesPath,

    [ValidateRange(1, 6)]
    [int]$StartStep = 1
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

$steps = @(
    @{
        Name = "Classify versions"
        Script = "ClassifyInfoPathVersions.ps1"
        Action = {
            & $scriptPath -SourceDirectory $sourceRoot -OutputCsvPath $versionCsvPath -Recurse:$Recurse.IsPresent
        }
    }
    @{
        Name = "Analyze structure by version"
        Script = "ProcessInfoPathAnalyze.ps1"
        Action = {
            & $scriptPath -SourceDirectory $sourceRoot -OutputDirectory $analysisDir -VersioningCsvPath $versionCsvPath -MaxFilesToAnalyze $MaxFilesToAnalyze -Recurse:$Recurse.IsPresent
        }
    }
    @{
        Name = "Extract normalized properties"
        Script = "ProcessInfoPathProperties.ps1"
        Action = {
            & $scriptPath -SourceDirectory $sourceRoot -OutputDirectory $reportsDir -VersioningCsvPath $versionCsvPath -AnalysisDirectory $analysisDir -Recurse:$Recurse.IsPresent
        }
    }
    @{
        Name = "Analyze grouping from version CSV"
        Script = "AnalyzeInfoPathGroupingFromCsv.ps1"
        Action = {
            & $scriptPath -InputCsvPath $versionCsvPath -OutputCsvPath $groupingDetailedPath -SummaryCsvPath $groupingSummaryPath
        }
    }
    @{
        Name = "Build XSLT templates per group"
        Script = "BuildXsltFromGroupingCsv.ps1"
        Action = {
            & $scriptPath -GroupingCsvPath $groupingDetailedPath -TemplateOutputDirectory $templatesDir -MappingCsvPath $templateMappingCsvPath -UseDisplayName $UseDisplayName -FieldRulesPath $FieldRulesPath
        }
    }
    @{
        Name = "Transform InfoPath XML to HTML"
        Script = "ProcessInfopathToHTML.ps1"
        Action = {
            & $scriptPath -InfoPathDirectory $sourceRoot -MetadataCsvPath $versionCsvPath -GroupingCsvPath $groupingDetailedPath -TemplateMappingCsvPath $templateMappingCsvPath -OutputDirectory $htmlDir -PrintToPdf $PrintToPdf -Recurse:$Recurse.IsPresent -UseDisplayName $UseDisplayName -DateFormat $DateFormat -DateTimeFormat $DateTimeFormat -FieldRulesPath $FieldRulesPath
        }
    }
)

Write-Stage -Group Init -Message "InfoPath local pipeline steps:"
for ($index = 0; $index -lt $steps.Count; $index++) {
    Write-Host ("  {0}) {1} ({2})" -f ($index + 1), $steps[$index].Name, $steps[$index].Script)
}

if (-not $PSBoundParameters.ContainsKey('StartStep')) {
    while ($true) {
        $selection = Read-Host "Enter the starting step (1-$($steps.Count)) [1]"
        if ([string]::IsNullOrWhiteSpace($selection)) {
            $StartStep = 1
            break
        }

        $selectedStep = 0
        if ([int]::TryParse($selection, [ref]$selectedStep) -and $selectedStep -ge 1 -and $selectedStep -le $steps.Count) {
            $StartStep = $selectedStep
            break
        }

        Write-Stage -Group Warn -Message "Invalid step. Enter a whole number from 1 to $($steps.Count)."
    }
}

Write-Stage -Group Input -Message "Starting at step $StartStep/$($steps.Count): $($steps[$StartStep - 1].Name)"
if ($StartStep -gt 1) {
    Write-Stage -Group Warn -Message "Earlier steps will be skipped. Their required outputs must already exist in OutputRootDirectory."
}

. (Join-Path -Path $PSScriptRoot -ChildPath 'Helper\HtmlRendering.ps1')
$null = Import-HtmlFieldRules -Path $FieldRulesPath

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

foreach ($step in $steps[($StartStep - 1)..($steps.Count - 1)]) {
    Assert-ScriptExists -ScriptPath (Join-Path -Path $PSScriptRoot -ChildPath $step.Script)
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
    for ($index = $StartStep - 1; $index -lt $steps.Count; $index++) {
        $step = $steps[$index]
        $stageLabel = "Step $($index + 1)/$($steps.Count): $($step.Name) ($($step.Script))"
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath $step.Script
        Write-Stage -Group Run -Message $stageLabel
        & $step.Action
        Write-Stage -Group Done -Message "$stageLabel completed."
    }

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
    Write-Stage -Group Error -Message "Pipeline failed at ${stageLabel}: $($_.Exception.Message)"
    throw
}
