<#
.SYNOPSIS
    Verifies HTML-only date formatting and read-only boolean checkboxes.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$testRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("InfoPathValueTest-" + [guid]::NewGuid().ToString('N'))

function Assert-HtmlValues {
    param(
        [string]$HtmlPath,
        [string[]]$ExpectedDates
    )

    $html = Get-Content -LiteralPath $HtmlPath -Raw
    $checkboxes = @([regex]::Matches($html, '<input\b[^>]*type="checkbox"[^>]*>'))
    if ($checkboxes.Count -ne 8) {
        throw "Expected 8 boolean checkboxes, got $($checkboxes.Count) in '$HtmlPath'."
    }
    $checkedCount = 0
    foreach ($checkbox in $checkboxes) {
        if ($checkbox.Value -notmatch '\bdisabled(?:=|\s|>)') {
            throw "Checkbox is not read-only in '$HtmlPath'."
        }
        if ($checkbox.Value -match '\bchecked(?:=|\s|>)') {
            $checkedCount++
            if ($checkbox.Value -notmatch 'aria-label="True"') {
                throw "Checked checkbox has an incorrect accessible label in '$HtmlPath'."
            }
        }
        elseif ($checkbox.Value -notmatch 'aria-label="False"') {
            throw "Unchecked checkbox has an incorrect accessible label in '$HtmlPath'."
        }
    }
    if ($checkedCount -ne 4) {
        throw "Expected 4 checked checkboxes, got $checkedCount in '$HtmlPath'."
    }
    foreach ($dateGroup in ($ExpectedDates | Group-Object)) {
        $count = [regex]::Matches($html, ('>\s*' + [regex]::Escape($dateGroup.Name) + '\s*<')).Count
        if ($count -ne $dateGroup.Count) {
            throw "Expected $($dateGroup.Count) occurrences of date '$($dateGroup.Name)', got $count in '$HtmlPath'."
        }
    }
    foreach ($value in ($ExpectedDates + @('1', '0', '12', '34', '2026-02-30', 'Not a date', 'Keep &amp; display'))) {
        if ($html -notmatch ('>\s*' + [regex]::Escape($value) + '\s*<')) {
            throw "Missing field value '$value' in '$HtmlPath'."
        }
    }
    foreach ($href in @('docs/file.txt', 'docs/one.txt', 'docs/two.txt')) {
        if (-not $html.Contains('href="' + $href + '"')) {
            throw "Attachment link '$href' was lost in '$HtmlPath'."
        }
    }
    if ($html -match '>\s*(true|false)\s*<') {
        throw "Boolean text was rendered instead of a checkbox in '$HtmlPath'."
    }
}

try {
    $source = Join-Path -Path $testRoot -ChildPath 'Source'
    New-Item -Path $source -ItemType Directory -Force | Out-Null
    $xmlPath = Join-Path -Path $source -ChildPath 'Sample.xml'
    @'
<?xml version="1.0" encoding="utf-8"?>
<?mso-infoPathSolution solutionVersion="1.0.0.1" href="test.xsn"?>
<my:FormData xmlns:my="urn:infopath-value-test" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xs="http://www.w3.org/2001/XMLSchema">
  <my:ApprovedDate>2026-10-23</my:ApprovedDate>
  <my:OffsetDateTime>2026-10-23T00:30:45.1234567+14:00</my:OffsetDateTime>
  <my:UtcDateTime>2026-10-23T23:45:00Z</my:UtcDateTime>
  <my:Details><my:LeapDate>2024-02-29</my:LeapDate></my:Details>
  <my:Approved> True </my:Approved>
  <my:Rejected>FALSE</my:Rejected>
  <my:TypedTrue xsi:type="xs:boolean">1</my:TypedTrue>
  <my:TypedFalse xsi:type="xs:boolean">0</my:TypedFalse>
  <my:NumericOne>1</my:NumericOne>
  <my:NumericZero>0</my:NumericZero>
  <my:EmptyDate />
  <my:EmptyBoolean xsi:type="xs:boolean" />
  <my:NilBoolean xsi:type="xs:boolean" xsi:nil="true" />
  <my:InvalidDate>2026-02-30</my:InvalidDate>
  <my:Text>Not a date</my:Text>
  <my:EscapedText>Keep &amp; display</my:EscapedText>
  <my:Record><my:Date>2025-12-31</my:Date><my:Flag>true</my:Flag><my:Count>12</my:Count></my:Record>
  <my:Record><my:Date>2026-01-02</my:Date><my:Flag>false</my:Flag><my:Count>34</my:Count></my:Record>
  <my:RepeatedFlag>true</my:RepeatedFlag>
  <my:RepeatedFlag>false</my:RepeatedFlag>
  <my:RepeatedDate>2025-12-31</my:RepeatedDate>
  <my:RepeatedDate>2026-01-02</my:RepeatedDate>
  <my:Attachment>ATTACHMENT::docs/file.txt</my:Attachment>
  <my:RepeatedAttachment>ATTACHMENT::docs/one.txt</my:RepeatedAttachment>
  <my:RepeatedAttachment>ATTACHMENT::docs/two.txt</my:RepeatedAttachment>
</my:FormData>
'@ | Set-Content -LiteralPath $xmlPath -Encoding UTF8
    $originalHash = (Get-FileHash -LiteralPath $xmlPath).Hash
    $outputs = @{}

    foreach ($mode in @('Default', 'Custom')) {
        $output = Join-Path -Path $testRoot -ChildPath $mode
        $parameters = @{
            SourceDirectory = $source
            OutputRootDirectory = $output
        }
        $dates = @('10/23/2026', '10/23/2026', '10/23/2026', '02/29/2024', '12/31/2025', '12/31/2025', '01/02/2026', '01/02/2026')
        if ($mode -eq 'Custom') {
            $parameters.DateFormat = 'dd/MM/yyyy'
            $dates = @('23/10/2026', '23/10/2026', '23/10/2026', '29/02/2024', '31/12/2025', '31/12/2025', '02/01/2026', '02/01/2026')
        }
        & (Join-Path -Path $repoRoot -ChildPath 'RunInfoPathPipeline.ps1') @parameters
        Assert-HtmlValues -HtmlPath (Join-Path -Path $output -ChildPath 'Html\Sample.html') -ExpectedDates $dates
        $outputs[$mode] = $output

        $mapping = @(Import-Csv -LiteralPath (Join-Path -Path $output -ChildPath 'Templates\InfoPathXsltGroupMap.csv'))
        if ($mapping.Count -ne 1 -or $mapping[0].GenerationStatus -ne 'Success') {
            throw "Template generation failed for '$mode'."
        }
        $transform = [System.Xml.Xsl.XslCompiledTransform]::new()
        $transform.Load($mapping[0].XsltPath)
        $directHtml = Join-Path -Path $output -ChildPath 'Direct.html'
        $transform.Transform($xmlPath, $directHtml)
        Assert-HtmlValues -HtmlPath $directHtml -ExpectedDates @('2026-10-23', '2026-10-23T00:30:45.1234567+14:00', '2026-10-23T23:45:00Z', '2024-02-29', '2025-12-31', '2025-12-31', '2026-01-02', '2026-01-02')

        $csvPaths = @(Get-ChildItem -LiteralPath (Join-Path -Path $output -ChildPath 'Reports\Normalized') -Filter 'Main.csv' -Recurse)
        if ($csvPaths.Count -ne 1) {
            throw "Expected one normalized Main.csv for '$mode'."
        }
        $csv = Get-Content -LiteralPath $csvPaths[0].FullName -Raw
        foreach ($value in @('2026-10-23', '2026-10-23T00:30:45.1234567+14:00', '2026-10-23T23:45:00Z')) {
            if (-not $csv.Contains($value)) {
                throw "Normalized CSV date '$value' was changed for '$mode'."
            }
        }
    }

    $output = $outputs.Default
    $htmlOutput = Join-Path -Path $testRoot -ChildPath 'Override'
    $htmlParameters = @{
        InfoPathDirectory = $source
        MetadataCsvPath = Join-Path -Path $output -ChildPath 'Reports\InfoPathVersionMap.csv'
        GroupingCsvPath = Join-Path -Path $output -ChildPath 'Reports\InfoPathGroupingDetailed.csv'
        TemplateMappingCsvPath = Join-Path -Path $output -ChildPath 'Templates\InfoPathXsltGroupMap.csv'
        OutputDirectory = $htmlOutput
        DateFormat = 'yyyy-MM-dd HH:mm:ss zzz'
    }
    & (Join-Path -Path $repoRoot -ChildPath 'ProcessInfopathToHTML.ps1') @htmlParameters
    Assert-HtmlValues -HtmlPath (Join-Path -Path $htmlOutput -ChildPath 'Sample.html') -ExpectedDates @('2026-10-23 00:30:45 +14:00', '2026-10-23 23:45:00 +00:00')

    foreach ($script in @('RunInfoPathPipeline.ps1', 'ProcessInfopathToHTML.ps1')) {
        $invalidParameters = if ($script -eq 'RunInfoPathPipeline.ps1') {
            @{ SourceDirectory = $source; OutputRootDirectory = $outputs.Default; DateFormat = '%' }
        }
        else {
            $htmlParameters.Clone()
        }
        $invalidParameters.DateFormat = '%'
        $rejected = $false
        try {
            & (Join-Path -Path $repoRoot -ChildPath $script) @invalidParameters
        }
        catch [System.Management.Automation.ParameterBindingException] {
            $rejected = $true
        }
        if (-not $rejected) {
            throw "Invalid DateFormat was accepted by '$script'."
        }
    }

    if ((Get-FileHash -LiteralPath $xmlPath).Hash -ne $originalHash) {
        throw 'Source XML was modified.'
    }
    Write-Host 'PASS: HTML date defaults and overrides, offsets, boolean checkboxes, repeating fields, attachments, invalid formats, and unchanged source XML/CSV dates.'
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
