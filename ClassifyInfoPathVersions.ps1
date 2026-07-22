<#
.SYNOPSIS
    Classifies InfoPath XML files by template version and schema field-count version.

.DESCRIPTION
    Scans InfoPath XML files, extracts template version from the mso-infoPathSolution
    processing instruction, computes schema characteristics, and writes a CSV report
    for downstream grouping analysis.

.PARAMETER SourceDirectory
    Root folder containing InfoPath XML files to analyze.
    This can be scanned recursively with -Recurse.
    Default: .\InfoPathXml

.PARAMETER OutputCsvPath
    Full path to the output CSV report.
    A log file is also written with the same path and .log extension.
    Default: .\InfoPathVersionMap.csv

.PARAMETER Recurse
    Include subfolders when discovering XML files.

.EXAMPLE
    .\ClassifyInfoPathVersions.ps1 `
      -SourceDirectory "C:\Save\InfoPath" `
      -OutputCsvPath "C:\Save\Reports\InfoPathVersionMap.csv" `
      -Recurse
#>

[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string]$SourceDirectory = ".\InfoPathXml",

    [ValidateNotNullOrEmpty()]
    [string]$OutputCsvPath = ".\InfoPathVersionMap.csv",

    [switch]$Recurse
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-LogEntry {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]]$LogEntries,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntries.Add("[$timestamp] $Message") | Out-Null
}

function Write-Stage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("Init", "Input", "Parse", "Analyze", "Write", "Done", "Warn", "Error")]
        [string]$Group = "Analyze"
    )

    $color = switch ($Group) {
        "Init" { "Cyan" }
        "Input" { "DarkCyan" }
        "Parse" { "Yellow" }
        "Analyze" { "White" }
        "Write" { "Magenta" }
        "Done" { "Blue" }
        "Warn" { "DarkYellow" }
        "Error" { "Red" }
        default { "White" }
    }

    Write-Host "[$Group] $Message" -ForegroundColor $color
}

function Get-InfoPathTemplateVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XmlRaw
    )

    $match = [System.Text.RegularExpressions.Regex]::Match(
        $XmlRaw,
        '<\?mso-infoPathSolution[^>]*solutionVersion\s*=\s*"([^"]+)"',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    if ($match.Success) {
        return $match.Groups[1].Value.Trim()
    }

    return $null
}

function Add-LeafFieldPaths {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$Node,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$Paths
    )

    $children = @($Node.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })

    if ($children.Count -eq 0) {
        $null = $Paths.Add($Path)
        return
    }

    foreach ($child in $children) {
        Add-LeafFieldPaths -Node $child -Path "$Path/$($child.Name)" -Paths $Paths
    }
}

if (-not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) {
    throw "SourceDirectory does not exist or is not a folder: $SourceDirectory"
}

$resolvedSourceDirectory = (Resolve-Path -LiteralPath $SourceDirectory).Path

$logEntries = New-Object System.Collections.Generic.List[string]
$startTime = Get-Date

Write-Stage -Group Init -Message "Starting InfoPath version/schema classification."
Write-Stage -Group Input -Message "Source directory: $resolvedSourceDirectory"
Write-Stage -Group Input -Message "Output CSV: $OutputCsvPath"

Write-LogEntry -LogEntries $logEntries -Message "Started classification."
Write-LogEntry -LogEntries $logEntries -Message "SourceDirectory=$resolvedSourceDirectory"
Write-LogEntry -LogEntries $logEntries -Message "OutputCsvPath=$OutputCsvPath"

$searchParameters = @{
    LiteralPath = $resolvedSourceDirectory
    Filter      = '*.xml'
    File        = $true
}

if ($Recurse.IsPresent) {
    $searchParameters.Recurse = $true
}

$xmlFiles = @(Get-ChildItem @searchParameters)
if ($xmlFiles.Count -eq 0) {
    throw "No XML files found in: $resolvedSourceDirectory"
}

Write-Stage -Group Input -Message "Found $($xmlFiles.Count) XML file(s)."

$rows = New-Object System.Collections.Generic.List[hashtable]

foreach ($xmlFile in $xmlFiles) {
    Write-Stage -Group Parse -Message "Parsing: $($xmlFile.FullName)"

    $row = @{
        SourceFilePath            = $xmlFile.FullName
        SourceFileName            = $xmlFile.Name
        TemplateVersion           = $null
        TemplateVersionNormalized = "UnknownVersion"
        FieldCount                = 0
        FieldCountVersion         = "0fd"
        SchemaHash                = $null
        AllTemplateVersionsSame   = $null
        ParseStatus               = "Success"
        ErrorMessage              = $null
        TemplateRoot              = $null
    }

    try {
        # Use a StreamReader for automatic BOM/encoding detection instead of
        # Get-Content which can mis-read UTF-16 or Windows-1252 InfoPath files.
        $streamReader = [System.IO.StreamReader]::new($xmlFile.FullName, $true)
        $rawXml = $streamReader.ReadToEnd()
        $streamReader.Close()
        $streamReader.Dispose()

        $xmlSettings = [System.Xml.XmlReaderSettings]::new()
        $xmlSettings.IgnoreComments = $true
        $xmlSettings.IgnoreWhitespace = $true
        # Prohibit DTD to avoid XXE vulnerabilities.
        $xmlSettings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit

        $xmlDocument = [System.Xml.XmlDocument]::new()
        $xmlDocument.XmlResolver = $null
        $xmlReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($rawXml), $xmlSettings)
        try {
            $xmlDocument.Load($xmlReader)
        }
        finally {
            $xmlReader.Dispose()
        }

        if ($null -eq $xmlDocument.DocumentElement) {
            throw "Document element is missing."
        }

        $templateVersion = Get-InfoPathTemplateVersion -XmlRaw $rawXml
        if (-not [string]::IsNullOrWhiteSpace($templateVersion)) {
            $row.TemplateVersion = $templateVersion
            $row.TemplateVersionNormalized = $templateVersion

            # Template root is the version without the last dot segment.
            if ($templateVersion -match '^(.*)\.[^.]+$') {
                $row.TemplateRoot = $Matches[1]
            }
            else {
                $row.TemplateRoot = $templateVersion
            }
        }
        else {
            $row.TemplateRoot = "UnknownTemplateRoot"
        }

        $leafPaths = New-Object System.Collections.Generic.HashSet[string]
        Add-LeafFieldPaths -Node $xmlDocument.DocumentElement -Path $xmlDocument.DocumentElement.Name -Paths $leafPaths

        $sortedLeafPaths = @($leafPaths | Sort-Object)
        $fieldCount = $sortedLeafPaths.Count
        $schemaText = [string]::Join('|', $sortedLeafPaths)

        $row.FieldCount = $fieldCount
        $row.FieldCountVersion = "{0}fd" -f $fieldCount

        if (-not [string]::IsNullOrWhiteSpace($schemaText)) {
            $hashBytes = [System.Security.Cryptography.SHA256]::Create().ComputeHash([System.Text.Encoding]::UTF8.GetBytes($schemaText))
            $row.SchemaHash = [System.BitConverter]::ToString($hashBytes).Replace('-', '').ToLowerInvariant()
        }

        Write-LogEntry -LogEntries $logEntries -Message ("Parsed '{0}' :: TemplateVersion='{1}' FieldCount={2}" -f $xmlFile.Name, $row.TemplateVersionNormalized, $row.FieldCount)
    }
    catch {
        $row.ParseStatus = "Failed"
        $row.ErrorMessage = $_.Exception.Message
        Write-Stage -Group Warn -Message "Failed to parse: $($xmlFile.FullName)"
        Write-Stage -Group Error -Message $_.Exception.Message
        Write-LogEntry -LogEntries $logEntries -Message ("ERROR: Failed to parse '{0}' :: {1}" -f $xmlFile.FullName, $_.Exception.Message)
    }

    $rows.Add($row)
}

$successfulRows = @($rows | Where-Object {
    $_.ContainsKey('ParseStatus') -and $_['ParseStatus'] -eq 'Success'
})

$allTemplateVersionsSame = $false
try {
    $versionUniverse = @(
        $successfulRows |
            ForEach-Object {
                if ($_.ContainsKey('TemplateVersionNormalized')) {
                    $_['TemplateVersionNormalized']
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )

    $allTemplateVersionsSame = ($versionUniverse.Count -le 1)

    Write-Stage -Group Analyze -Message "All template versions same: $allTemplateVersionsSame"
    if ($allTemplateVersionsSame) {
        Write-Stage -Group Analyze -Message "Using field-count schema groups (example: 10fd, 13fd)."
    }
    else {
        Write-Stage -Group Analyze -Message "Using template version groups; missing versions fall back to field-count groups."
    }
}
catch {
    $allTemplateVersionsSame = $false
    Write-Stage -Group Warn -Message "Analysis step encountered an error; continuing with output generation."
    Write-Stage -Group Error -Message $_.Exception.Message
    Write-LogEntry -LogEntries $logEntries -Message ("WARN: Analysis step failed but processing continued :: {0}" -f $_.Exception.Message)
}

foreach ($row in $rows) {
    $row.AllTemplateVersionsSame = $allTemplateVersionsSame
}

$outputDirectory = Split-Path -Path $OutputCsvPath -Parent
if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

$columnOrder = @(
    'SourceFilePath',
    'SourceFileName',
    'TemplateVersion',
    'TemplateVersionNormalized',
    'TemplateRoot',
    'FieldCount',
    'FieldCountVersion',
    'SchemaHash',
    'AllTemplateVersionsSame',
    'ParseStatus',
    'ErrorMessage'
)

$outputRows = foreach ($row in $rows) {
    $ordered = [ordered]@{}
    foreach ($column in $columnOrder) {
        $ordered[$column] = $row[$column]
    }

    [pscustomobject]$ordered
}

Write-Stage -Group Write -Message "Writing CSV report."
$outputRows | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8

$logPath = [System.IO.Path]::ChangeExtension($OutputCsvPath, '.log')
Write-Stage -Group Write -Message "Writing log file."

$duration = (Get-Date) - $startTime
Write-LogEntry -LogEntries $logEntries -Message "Completed. Duration=$($duration.ToString())"
Write-LogEntry -LogEntries $logEntries -Message ("Summary: Files={0}, Success={1}, Failed={2}, AllTemplateVersionsSame={3}" -f $rows.Count, $successfulRows.Count, ($rows.Count - $successfulRows.Count), $allTemplateVersionsSame)
$logEntries | Set-Content -LiteralPath $logPath -Encoding UTF8

Write-Stage -Group Done -Message "Completed classification workflow."
Write-Host "CSV report: $OutputCsvPath" -ForegroundColor Blue
Write-Host "Log file  : $logPath" -ForegroundColor Blue
