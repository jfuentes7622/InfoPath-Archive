<#
.SYNOPSIS
    Verifies pipeline start-step selection, ordering, logging, and failure handling.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$testRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("InfoPathResumeTest-" + [guid]::NewGuid().ToString('N'))
$scriptNames = @(
    'ProcessXsnTemplates.ps1'
    'ClassifyInfoPathVersions.ps1'
    'ProcessInfoPathAnalyze.ps1'
    'ProcessInfoPathProperties.ps1'
    'AnalyzeInfoPathGroupingFromCsv.ps1'
    'BuildXsltFromGroupingCsv.ps1'
    'ProcessInfopathToHTML.ps1'
)

function Read-Host {
    param([string]$Prompt)
    if ($promptAnswers.Count -eq 0) {
        throw "Unexpected prompt: $Prompt"
    }
    return $promptAnswers.Dequeue()
}

function Assert-ResumeRun {
    param(
        [int]$ExpectedStart,
        [hashtable]$Parameters,
        [string[]]$Answers = @()
    )

    $promptAnswers = [System.Collections.Generic.Queue[string]]::new()
    foreach ($answer in $Answers) {
        $promptAnswers.Enqueue($answer)
    }
    if (Test-Path -LiteralPath $callsPath) {
        Remove-Item -LiteralPath $callsPath
    }
    $messages = @(& $pipelinePath @Parameters 6>&1 | ForEach-Object { $_.ToString() })
    $calls = @(Get-Content -LiteralPath $callsPath | ForEach-Object { $_ | ConvertFrom-Json })
    $skipXsn = -not $Parameters.ContainsKey('XsnExportCsvPath')
    $firstInvocation = if ($skipXsn -and $ExpectedStart -eq 1) { 2 } else { $ExpectedStart }
    $expectedScripts = $scriptNames[($firstInvocation - 1)..6]
    if (($calls.Script -join ',') -ne ($expectedScripts -join ',')) {
        throw "Incorrect execution order for start step ${ExpectedStart}: $($calls.Script -join ',')"
    }
    if ($promptAnswers.Count -ne 0) {
        throw 'Not all prompt answers were consumed.'
    }
    $lastInvocation = -1
    for ($index = 1; $index -le 7; $index++) {
        if (-not ($messages | Where-Object { $_ -like "  $index) *($($scriptNames[$index - 1]))" })) {
            throw "Menu is missing step $index."
        }
        $starts = @($messages | Where-Object { $_ -like "[[]Run] Step $index/7:*" })
        if ($index -lt $ExpectedStart -or ($skipXsn -and $index -eq 1)) {
            if ($starts.Count -ne 0) {
                throw "Skipped step $index was announced as running."
            }
            continue
        }
        if ($starts.Count -ne 1) {
            throw "Expected one start announcement for step $index."
        }
        $startIndex = [array]::IndexOf($messages, $starts[0])
        $invocationIndex = [array]::IndexOf($messages, "Invoked $($scriptNames[$index - 1])")
        $completed = @($messages | Where-Object { $_ -like "[[]Done] Step $index/7:* completed." })
        if ($completed.Count -ne 1 -or $startIndex -le $lastInvocation -or $invocationIndex -le $startIndex -or [array]::IndexOf($messages, $completed[0]) -le $invocationIndex) {
            throw "Step $index was not announced before execution and completed afterward."
        }
        if ($skipXsn -and $ExpectedStart -eq 1 -and -not ($messages | Where-Object { $_ -like '*Skipping Step 1/7:*no -XsnExportCsvPath*' })) {
            throw 'An XML-only run did not explicitly report skipped XSN extraction.'
        }
        $lastInvocation = $invocationIndex
    }
    foreach ($call in $calls) {
        $bound = $call.Parameters
        if ($call.Script -eq 'ProcessXsnTemplates.ps1' -and
            ($bound.InputCsvPath -ne $Parameters.XsnExportCsvPath -or $bound.OutputDirectory -ne (Join-Path -Path $outputPath -ChildPath 'XSNTemplates'))) {
            throw 'XSN input or output directory was not forwarded correctly.'
        }
        foreach ($name in @('SourceDirectory', 'InfoPathDirectory')) {
            if ($bound.PSObject.Properties.Name -contains $name -and $bound.$name -ne $sourcePath) {
                throw "Source path was not forwarded to $($call.Script)."
            }
        }
        if ($bound.PSObject.Properties.Name -contains 'UseXSN4HTML' -and $bound.UseXSN4HTML -ne $true) {
            throw "UseXSN4HTML was not forwarded to $($call.Script)."
        }
        if ($bound.PSObject.Properties.Name -contains 'Recurse' -and -not $bound.Recurse.IsPresent) {
            throw "Recurse was not forwarded to $($call.Script)."
        }
    }
}

try {
    $sourcePath = Join-Path -Path $testRoot -ChildPath 'Source'
    $outputPath = Join-Path -Path $testRoot -ChildPath 'Output'
    $helperPath = Join-Path -Path $testRoot -ChildPath 'Helper'
    New-Item -Path $sourcePath, $helperPath -ItemType Directory -Force | Out-Null
    $pipelinePath = Join-Path -Path $testRoot -ChildPath 'RunInfoPathPipeline.ps1'
    Copy-Item -LiteralPath (Join-Path -Path $repoRoot -ChildPath 'RunInfoPathPipeline.ps1') -Destination $pipelinePath
    'function Import-HtmlFieldRules { param([string]$Path) }' |
        Set-Content -LiteralPath (Join-Path -Path $helperPath -ChildPath 'HtmlRendering.ps1')
    'function Import-XsnInventory { param([string]$Path) }' |
        Set-Content -LiteralPath (Join-Path -Path $helperPath -ChildPath 'XsnTemplates.ps1')
    $callsPath = Join-Path -Path $testRoot -ChildPath 'calls.jsonl'
    $failurePath = Join-Path -Path $testRoot -ChildPath 'FailAt.txt'
    $stub = @'
[CmdletBinding()]
param(
    [string]$SourceDirectory, [string]$InfoPathDirectory, [string]$OutputDirectory,
    [string]$OutputCsvPath, [string]$VersioningCsvPath, [int]$MaxFilesToAnalyze,
    [switch]$Recurse, [string]$AnalysisDirectory, [string]$InputCsvPath,
    [string]$SummaryCsvPath, [string]$GroupingCsvPath, [string]$TemplateOutputDirectory,
    [string]$MappingCsvPath, [bool]$UseDisplayName, [string]$FieldRulesPath,
    [string]$MetadataCsvPath, [string]$TemplateMappingCsvPath, [string]$PrintToPdf,
    [string]$DateFormat, [string]$DateTimeFormat, [string]$XsnInventoryCsvPath,
    [bool]$UseXSN4HTML, [switch]$ExtractResources
)
$name = Split-Path -Path $PSCommandPath -Leaf
@{ Script = $name; Parameters = $PSBoundParameters } | ConvertTo-Json -Depth 5 -Compress |
    Add-Content -LiteralPath (Join-Path -Path $PSScriptRoot -ChildPath 'calls.jsonl')
Write-Host "Invoked $name"
$failure = Join-Path -Path $PSScriptRoot -ChildPath 'FailAt.txt'
if ((Test-Path -LiteralPath $failure) -and (Get-Content -LiteralPath $failure) -eq $name) {
    throw "Intentional failure in $name"
}
'@
    foreach ($name in $scriptNames) {
        Set-Content -LiteralPath (Join-Path -Path $testRoot -ChildPath $name) -Value $stub
    }
    $parameters = @{
        SourceDirectory = $sourcePath
        OutputRootDirectory = $outputPath
        Recurse = $true
        UseXSN4HTML = $true
        XsnExportCsvPath = (Join-Path -Path $testRoot -ChildPath 'Export\XsnTemplateDownloads.csv')
    }
    foreach ($start in 1..7) {
        Assert-ResumeRun -ExpectedStart $start -Parameters ($parameters + @{ StartStep = $start })
        Assert-ResumeRun -ExpectedStart $start -Parameters $parameters -Answers @("$start")
    }
    Assert-ResumeRun -ExpectedStart 1 -Parameters $parameters -Answers @('')
    Assert-ResumeRun -ExpectedStart 1 -Parameters $parameters -Answers @('   ')
    Assert-ResumeRun -ExpectedStart 4 -Parameters $parameters -Answers @('0', '8', '-1', 'abc', '2.5', '9999999999999999', '4')
    $xmlOnlyParameters = $parameters.Clone()
    $xmlOnlyParameters.Remove('XsnExportCsvPath')
    Assert-ResumeRun -ExpectedStart 1 -Parameters ($xmlOnlyParameters + @{ StartStep = 1 })
    foreach ($invalid in @(0, 8)) {
        $rejected = $false
        try {
            & $pipelinePath @parameters -StartStep $invalid 6>$null
        }
        catch [System.Management.Automation.ParameterBindingException] {
            $rejected = $true
        }
        if (-not $rejected) {
            throw "Invalid StartStep $invalid was accepted."
        }
    }
    Remove-Item -LiteralPath $callsPath
    Set-Content -LiteralPath $failurePath -Value $scriptNames[3]
    $failureMessages = [System.Collections.Generic.List[string]]::new()
    $caught = $false
    try {
        & $pipelinePath @parameters -StartStep 3 6>&1 | ForEach-Object { $failureMessages.Add($_.ToString()) }
    }
    catch {
        if ($_.Exception.Message -notlike 'Intentional failure*') {
            throw
        }
        $caught = $true
    }
    $failedCalls = @(Get-Content -LiteralPath $callsPath | ForEach-Object { $_ | ConvertFrom-Json })
    if (-not $caught -or ($failedCalls.Script -join ',') -ne ($scriptNames[2..3] -join ',')) {
        throw 'The pipeline did not stop and rethrow at the failing step.'
    }
    if (-not ($failureMessages | Where-Object { $_ -like '[[]Error] Pipeline failed at Step 4/7:*Intentional failure*' })) {
        throw 'Failure output did not identify step 4.'
    }
    if ($failureMessages | Where-Object { $_ -like '*Pipeline completed successfully*' -or $_ -like '[[]Done] Step 4/7:*' }) {
        throw 'A failed run reported success.'
    }
    Write-Host 'Pipeline resume tests passed.' -ForegroundColor Green
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
