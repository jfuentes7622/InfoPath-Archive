<#
.SYNOPSIS
    Verifies PDF browser arguments, failure handling, and real printing with spaced paths.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$testRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("InfoPath PDF Test-" + [guid]::NewGuid().ToString('N'))
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    (Join-Path -Path $repoRoot -ChildPath 'ProcessInfopathToHTML.ps1'), [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) {
    throw "Unable to parse HTML processor: $($parseErrors.Message -join '; ')"
}
foreach ($name in @('Find-PdfBrowser', 'Convert-HtmlToPdf')) {
    $definition = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    }, $true)
    if ($null -eq $definition) {
        throw "Missing function '$name'."
    }
    . ([scriptblock]::Create($definition.Extent.Text))
}
$browser = Find-PdfBrowser
if ($null -eq $browser) {
    throw 'The PDF integration test requires Microsoft Edge or Google Chrome.'
}

try {
    New-Item -Path $testRoot -ItemType Directory -Force | Out-Null
    $htmlPath = Join-Path -Path $testRoot -ChildPath 'Sample Form #1.html'
    $pdfPath = Join-Path -Path $testRoot -ChildPath 'Sample Form #1.pdf'
    '<html><body>PDF test</body></html>' | Set-Content -LiteralPath $htmlPath -Encoding UTF8
    $script:profilePaths = [System.Collections.Generic.List[string]]::new()
    $script:mode = 'Valid'

    function Start-Process {
        param(
            [string]$FilePath,
            [string[]]$ArgumentList,
            [switch]$Wait,
            [switch]$PassThru,
            [switch]$NoNewWindow
        )

        if ($FilePath -ne $browser -or -not $Wait -or -not $PassThru -or -not $NoNewWindow) {
            throw 'Unexpected browser invocation.'
        }
        foreach ($flag in @('--headless', '--no-first-run', '--no-default-browser-check', '--disable-extensions', '--no-pdf-header-footer')) {
            if ($ArgumentList -notcontains $flag) {
                throw "Missing browser flag '$flag'."
            }
        }
        $commandLine = $ArgumentList -join ' '
        if ($commandLine -notmatch '--user-data-dir="([^"]+)"') {
            throw 'Browser profile path was not quoted.'
        }
        $profile = $Matches[1]
        if (-not (Test-Path -LiteralPath $profile -PathType Container) -or
            $script:profilePaths.Contains($profile)) {
            throw 'Browser profile is missing or was reused.'
        }
        $script:profilePaths.Add($profile)
        if ($commandLine -notmatch '--print-to-pdf="([^"]+)"') {
            throw 'PDF path was not quoted.'
        }
        $temporaryPdf = $Matches[1]
        if ($temporaryPdf -eq $pdfPath -or $temporaryPdf -notlike '* *') {
            throw 'Expected an isolated temporary PDF path containing spaces.'
        }
        $expectedUri = [uri]::new($htmlPath).AbsoluteUri
        if ($ArgumentList -notcontains ('"' + $expectedUri + '"') -or
            -not $expectedUri.Contains('%20') -or -not $expectedUri.Contains('%23')) {
            throw 'HTML URI was not quoted and escaped correctly.'
        }

        switch ($script:mode) {
            'Valid' { [System.IO.File]::WriteAllText($temporaryPdf, "%PDF-1.7`nTest output`n%%EOF") }
            'Empty' { [System.IO.File]::WriteAllText($temporaryPdf, '') }
            'Invalid' { [System.IO.File]::WriteAllText($temporaryPdf, "Not a PDF`n%%EOF") }
            'Truncated' { [System.IO.File]::WriteAllText($temporaryPdf, "%PDF-1.7`nNo end marker") }
            'Missing' { }
            'ExitFailure' { }
            default { throw "Unknown mock mode '$script:mode'." }
        }
        $exitCode = if ($script:mode -eq 'ExitFailure') { 9 } else { 0 }
        $result = [pscustomobject]@{ ExitCode = $exitCode }
        $result | Add-Member -MemberType ScriptMethod -Name Dispose -Value {}
        return $result
    }

    try {
        foreach ($case in @('Valid', 'Missing', 'Empty', 'Invalid', 'Truncated', 'ExitFailure')) {
            $script:mode = $case
            [System.IO.File]::WriteAllText($pdfPath, 'Existing destination')
            $failure = $null
            try {
                Convert-HtmlToPdf -BrowserPath $browser -HtmlPath $htmlPath -PdfPath $pdfPath
            }
            catch {
                $failure = $_.Exception.Message
            }
            if ($case -eq 'Valid') {
                if ($null -ne $failure -or -not ([System.IO.File]::ReadAllText($pdfPath)).StartsWith('%PDF-')) {
                    throw "Valid output was not published: $failure"
                }
            }
            else {
                $expectedError = switch ($case) {
                    'Missing' { 'browser did not create a PDF' }
                    'Empty' { 'empty or truncated' }
                    'Invalid' { 'invalid PDF signature or missing end marker' }
                    'Truncated' { 'invalid PDF signature or missing end marker' }
                    'ExitFailure' { 'browser exited with code 9' }
                }
                if ($null -eq $failure -or -not $failure.Contains($expectedError)) {
                    throw "Case '$case' did not report the expected failure: $failure"
                }
                if ([System.IO.File]::ReadAllText($pdfPath) -ne 'Existing destination') {
                    throw "Case '$case' overwrote an existing destination PDF."
                }
            }
            foreach ($profile in $script:profilePaths) {
                if (Test-Path -LiteralPath (Split-Path -Path $profile -Parent)) {
                    throw "Temporary PDF/profile directory was not removed: $profile"
                }
            }
        }
    }
    finally {
        Remove-Item -LiteralPath Function:\Start-Process
    }

    $source = Join-Path -Path $testRoot -ChildPath 'Source Forms'
    $nestedSource = Join-Path -Path $source -ChildPath 'Nested Folder'
    $output = Join-Path -Path $testRoot -ChildPath 'Pipeline Output'
    New-Item -Path $nestedSource -ItemType Directory -Force | Out-Null
    @'
<?xml version="1.0" encoding="utf-8"?>
<?mso-infoPathSolution solutionVersion="1.0.0.1" href="test.xsn"?>
<my:FormData xmlns:my="urn:infopath-pdf-test">
  <my:RequestDate>2026-10-23</my:RequestDate>
  <my:Approved>true</my:Approved>
  <my:Description>PDF generation with spaces in paths</my:Description>
  <my:Notes><div xmlns="http://www.w3.org/1999/xhtml"><p>Rich text <strong>PDF content</strong></p></div></my:Notes>
</my:FormData>
'@ | Set-Content -LiteralPath (Join-Path -Path $nestedSource -ChildPath 'Sample Form #1.xml') -Encoding UTF8
    & (Join-Path -Path $repoRoot -ChildPath 'RunInfoPathPipeline.ps1') `
        -SourceDirectory $source -OutputRootDirectory $output -Recurse -PrintToPdf 'true' -StartStep 1
    $generatedPdf = Join-Path -Path $output -ChildPath 'Html\Nested Folder\Sample Form #1.pdf'
    if (-not (Test-Path -LiteralPath $generatedPdf -PathType Leaf)) {
        throw "Browser failed to generate PDF. See $(Join-Path -Path $output -ChildPath 'Html\ProcessInfopathToHTML-errors.log')."
    }
    $bytes = [System.IO.File]::ReadAllBytes($generatedPdf)
    if ($bytes.Length -lt 1000 -or [System.Text.Encoding]::ASCII.GetString($bytes, 0, 5) -ne '%PDF-' -or
        -not [System.Text.Encoding]::ASCII.GetString($bytes, [Math]::Max(0, $bytes.Length - 1024), [Math]::Min(1024, $bytes.Length)).Contains('%%EOF')) {
        throw 'Real browser output is not a complete PDF.'
    }
    $log = Get-Content -LiteralPath (Join-Path -Path $output -ChildPath 'Html\ProcessInfopathToHTML-errors.log') -Raw
    if (-not $log.Contains('No processing errors.')) {
        throw "Unexpected pipeline errors: $log"
    }
    Write-Host "PASS: Quoted paths, unique profiles, header/footer flags, PDF validation, failure handling, cleanup, preserved existing output, and real browser pipeline PDF ($($bytes.Length) bytes)."
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
