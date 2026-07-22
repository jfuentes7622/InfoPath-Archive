<#
.SYNOPSIS
    Analyzes InfoPath classification CSV data and computes grouping recommendations.

.DESCRIPTION
    Reads CSV output from ClassifyInfoPathVersions.ps1 and groups files using this rule:
    - Same TemplateVersion => same group.
    - Different TemplateVersion => different group.

    The script writes:
    1) A detailed CSV with SuggestedGroup for each file
    2) A summary CSV with group counts and version spread

.PARAMETER InputCsvPath
    Path to the CSV produced by ClassifyInfoPathVersions.ps1.

.PARAMETER OutputCsvPath
    Detailed output CSV containing per-file SuggestedGroup.
    Default: .\InfoPathGroupingDetailed.csv

.PARAMETER SummaryCsvPath
    Summary output CSV containing one row per SuggestedGroup.
    Default: .\InfoPathGroupingSummary.csv

.EXAMPLE
    .\AnalyzeInfoPathGroupingFromCsv.ps1 `
      -InputCsvPath "C:\Save\Reports\InfoPathVersionMap.csv"

.EXAMPLE
    .\AnalyzeInfoPathGroupingFromCsv.ps1 `
      -InputCsvPath "C:\Save\Reports\InfoPathVersionMap.csv" `
      -OutputCsvPath "C:\Save\Reports\InfoPathGroupingDetailed.csv" `
      -SummaryCsvPath "C:\Save\Reports\InfoPathGroupingSummary.csv"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$InputCsvPath,

    [ValidateNotNullOrEmpty()]
    [string]$OutputCsvPath = ".\InfoPathGroupingDetailed.csv",

    [ValidateNotNullOrEmpty()]
    [string]$SummaryCsvPath = ".\InfoPathGroupingSummary.csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Stage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("Init", "Input", "Analyze", "Write", "Done", "Warn", "Error")]
        [string]$Group = "Analyze"
    )

    $color = switch ($Group) {
        "Init" { "Cyan" }
        "Input" { "DarkCyan" }
        "Analyze" { "Yellow" }
        "Write" { "Green" }
        "Done" { "Magenta" }
        "Warn" { "DarkYellow" }
        "Error" { "Red" }
        default { "White" }
    }

    Write-Host "[$Group] $Message" -ForegroundColor $color
}

if (-not (Test-Path -LiteralPath $InputCsvPath -PathType Leaf)) {
    throw "InputCsvPath does not exist: $InputCsvPath"
}

Write-Stage -Group Init -Message "Starting CSV grouping analysis."
Write-Stage -Group Input -Message "Input CSV: $InputCsvPath"
Write-Stage -Group Write -Message "Detailed output: $OutputCsvPath"
Write-Stage -Group Write -Message "Summary output: $SummaryCsvPath"

$rows = @(Import-Csv -LiteralPath $InputCsvPath)
if ($rows.Count -eq 0) {
    throw "Input CSV has no rows: $InputCsvPath"
}

Write-Stage -Group Input -Message "Loaded $($rows.Count) row(s)."

$detailedRows = New-Object System.Collections.Generic.List[object]

foreach ($row in $rows) {
    $parseStatus = [string]$row.ParseStatus
    $fieldCountValue = 0
    [void][int]::TryParse([string]$row.FieldCount, [ref]$fieldCountValue)

    $templateVersion = [string]$row.TemplateVersionNormalized
    if ([string]::IsNullOrWhiteSpace($templateVersion)) {
        $templateVersion = [string]$row.TemplateVersion
    }
    if ([string]::IsNullOrWhiteSpace($templateVersion)) {
        $templateVersion = "UnknownVersion"
    }

    $suggestedGroup = if ($parseStatus -ne "Success") {
        "ParseFailed"
    }
    else {
        $templateVersion
    }

    $groupReason = if ($parseStatus -ne "Success") {
        "Parse failed"
    }
    else {
        "Grouped by TemplateVersion"
    }

    $detailedRows.Add([pscustomobject]@{
        SourceFilePath      = $row.SourceFilePath
        SourceFileName      = $row.SourceFileName
        ParseStatus         = $parseStatus
        ErrorMessage        = $row.ErrorMessage
        TemplateVersion     = $row.TemplateVersion
        TemplateVersionNorm = $templateVersion
        FieldCount          = $fieldCountValue
        FieldCountVersion   = $row.FieldCountVersion
        SuggestedGroup      = $suggestedGroup
        GroupReason         = $groupReason
    })
}

Write-Stage -Group Analyze -Message "Computing summary per SuggestedGroup."
$summaryRows = @(
    $detailedRows |
        Group-Object -Property SuggestedGroup |
        Sort-Object Name |
        ForEach-Object {
            $groupRows = @($_.Group)
            $versions = @(
                $groupRows |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_.TemplateVersionNorm) -and $_.TemplateVersionNorm -ne "UnknownVersion" } |
                    Select-Object -ExpandProperty TemplateVersionNorm -Unique |
                    Sort-Object
            )

            $fieldCounts = @(
                $groupRows |
                    Select-Object -ExpandProperty FieldCount -Unique |
                    Sort-Object
            )

            [pscustomobject]@{
                SuggestedGroup      = $_.Name
                FileCount           = $groupRows.Count
                FieldCountVariants  = $fieldCounts.Count
                VersionVariants     = $versions.Count
                FieldCounts         = ($fieldCounts -join ";")
                Versions            = ($versions -join ";")
            }
        }
)

foreach ($path in @($OutputCsvPath, $SummaryCsvPath)) {
    $dir = Split-Path -Path $path -Parent
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir -PathType Container)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
}

Write-Stage -Group Write -Message "Writing detailed CSV."
$detailedRows | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8

Write-Stage -Group Write -Message "Writing summary CSV."
$summaryRows | Export-Csv -Path $SummaryCsvPath -NoTypeInformation -Encoding UTF8

Write-Stage -Group Done -Message "Grouping analysis complete."
Write-Host "Detailed CSV: $OutputCsvPath" -ForegroundColor Magenta
Write-Host "Summary CSV : $SummaryCsvPath" -ForegroundColor Magenta
