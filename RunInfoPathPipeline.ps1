<#
.SYNOPSIS
    Runs the local InfoPath processing pipeline end-to-end.

.DESCRIPTION
    Orchestrates the non-SharePoint child scripts in this folder in dependency order.
    This intentionally excludes SharePoint instance scripts such as:
      - MoveLibraryFilesToLocal_SPS.ps1
      - GetMetadataforLib_SPS.ps1

    Pipeline order:
      1) ProcessXsnTemplates.ps1 (when XsnExportCsvPath is supplied)
      2) ClassifyInfoPathVersions.ps1
      3) ProcessInfoPathAnalyze.ps1
      4) ProcessInfoPathProperties.ps1
      5) AnalyzeInfoPathGroupingFromCsv.ps1
      6) BuildXsltFromGroupingCsv.ps1
      7) ProcessInfopathToHTML.ps1

.PARAMETER SourceDirectory
    Root directory containing InfoPath XML files.

.PARAMETER OutputRootDirectory
    Root folder for all pipeline outputs.

.PARAMETER StartStep
    First pipeline step to run (1 through 7). If omitted, show the numbered
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

.PARAMETER UseXSN4HTML
    Prefer the matched exact-version XSN view for read-only archival HTML,
    otherwise the nearest higher solution version of the same form.
    Falls back to generated HTML with a warning and audit reason. Default: $false.

.PARAMETER FieldRulesPath
    Optional JSON field definitions, DateFormat, DateTimeFormat, and display-value mappings for HTML output.
    Passed to template generation and HTML rendering; CSV extraction is unchanged.

.PARAMETER XsnExportCsvPath
    Optional path to XsnTemplateDownloads.csv copied from the SPS export along
    with its XSN files, or a folder containing exactly one such CSV anywhere
    beneath it. Multiple inventories require an explicit selection by path.
    Step 1 extracts manifest information on this machine
    and view captions into OutputRootDirectory\XSNTemplates. Steps 6 and 7 use
    that inventory for labels. With UseXSN4HTML, also extracts rendering assets.
    Preparation errors warn and continue without XSNs, preserving CSV extraction.
    When omitted, step 1 is explicitly skipped;
    existing processed inventory under that folder is reused when resuming.

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

    [bool]$UseXSN4HTML = $false,

    [string]$FieldRulesPath,

    [ValidateRange(1, 7)]
    [int]$StartStep = 1,

    [string]$XsnExportCsvPath
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
        Name = "Extract XSN template information"
        Script = "ProcessXsnTemplates.ps1"
        Action = {
            & $scriptPath -InputCsvPath $XsnExportCsvPath -OutputDirectory $xsnTemplatesDir -ExtractResources:$UseXSN4HTML
        }
    }
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
            & $scriptPath -GroupingCsvPath $groupingDetailedPath -TemplateOutputDirectory $templatesDir -MappingCsvPath $templateMappingCsvPath -FieldRulesPath $FieldRulesPath -XsnInventoryCsvPath $xsnInventoryPath
        }
    }
    @{
        Name = "Transform InfoPath XML to HTML"
        Script = "ProcessInfopathToHTML.ps1"
        Action = {
            & $scriptPath -InfoPathDirectory $sourceRoot -MetadataCsvPath $versionCsvPath -GroupingCsvPath $groupingDetailedPath -TemplateMappingCsvPath $templateMappingCsvPath -OutputDirectory $htmlDir -PrintToPdf $PrintToPdf -Recurse:$Recurse.IsPresent -UseXSN4HTML $UseXSN4HTML -FieldRulesPath $FieldRulesPath -XsnInventoryCsvPath $xsnInventoryPath
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
. (Join-Path -Path $PSScriptRoot -ChildPath 'Helper\XsnTemplates.ps1')
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
$xsnTemplatesDir = Join-Path -Path $outputRoot -ChildPath "XSNTemplates"
$processedXsnCsv = Join-Path -Path $xsnTemplatesDir -ChildPath 'XsnTemplateVersions.csv'
$xsnInventoryPath = if (-not [string]::IsNullOrWhiteSpace($XsnExportCsvPath) -or (Test-Path -LiteralPath $processedXsnCsv -PathType Leaf)) { $processedXsnCsv } else { '' }

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
    if ($step.Script -eq 'ProcessXsnTemplates.ps1' -and [string]::IsNullOrWhiteSpace($XsnExportCsvPath)) {
        continue
    }
    Assert-ScriptExists -ScriptPath (Join-Path -Path $PSScriptRoot -ChildPath $step.Script)
}

Write-Stage -Group Init -Message "Starting InfoPath local pipeline orchestration."
Write-Stage -Group Input -Message "Source directory: $sourceRoot"
Write-Stage -Group Input -Message "Output root: $outputRoot"
Write-Stage -Group Input -Message "Reports directory: $reportsDir"
Write-Stage -Group Input -Message "Analysis directory: $analysisDir"
Write-Stage -Group Input -Message "Templates directory: $templatesDir"
Write-Stage -Group Input -Message "HTML directory: $htmlDir"
if (-not [string]::IsNullOrWhiteSpace($XsnExportCsvPath)) {
    Write-Stage -Group Input -Message "XSN export inventory: $XsnExportCsvPath"
    Write-Stage -Group Input -Message "XSN information directory: $xsnTemplatesDir"
}

$startTime = Get-Date

try {
    for ($index = $StartStep - 1; $index -lt $steps.Count; $index++) {
        $step = $steps[$index]
        $stageLabel = "Step $($index + 1)/$($steps.Count): $($step.Name) ($($step.Script))"
        if ($index -eq 0 -and [string]::IsNullOrWhiteSpace($XsnExportCsvPath)) {
            Write-Stage -Group Warn -Message "Skipping ${stageLabel}: no -XsnExportCsvPath supplied."
            continue
        }
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath $step.Script
        Write-Stage -Group Run -Message $stageLabel
        if ($index -eq 0) {
            try {
                & $step.Action
                Write-Stage -Group Done -Message "$stageLabel completed."
            }
            catch {
                $xsnInventoryPath = ''
                Write-Stage -Group Warn -Message "XSN preparation failed; continuing XML extraction and generated HTML without XSNs: $($_.Exception.Message)"
            }
            continue
        }
        if ($index -ge 5 -and $xsnInventoryPath) {
            try { $null = Import-XsnInventory -Path $xsnInventoryPath }
            catch {
                $xsnInventoryPath = ''
                Write-Stage -Group Warn -Message "XSN inventory unavailable; using generated HTML: $($_.Exception.Message)"
            }
        }
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
    if (-not [string]::IsNullOrWhiteSpace($XsnExportCsvPath)) {
        Write-Host "XSN information root: $xsnTemplatesDir" -ForegroundColor Magenta
    }
}
catch {
    Write-Stage -Group Error -Message "Pipeline failed at ${stageLabel}: $($_.Exception.Message)"
    throw
}
