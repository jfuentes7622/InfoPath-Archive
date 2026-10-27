<#
.SYNOPSIS
    Verifies rich-text content, field labels, nested sections, and repeating records.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$testRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("InfoPathRichTextTest-" + [guid]::NewGuid().ToString('N'))

function Get-RenderedDocument {
    param(
        [string]$XsltPath,
        [string]$XmlPath,
        [bool]$DisplayNames
    )

    [xml]$stylesheet = Get-Content -LiteralPath $XsltPath -Raw
    $namespaces = [System.Xml.XmlNamespaceManager]::new($stylesheet.NameTable)
    $namespaces.AddNamespace('xsl', 'http://www.w3.org/1999/XSL/Transform')
    # XML serialization makes the generated DOM queryable without changing the rendering templates.
    $stylesheet.SelectSingleNode('/xsl:stylesheet/xsl:output', $namespaces).SetAttribute('method', 'xml')
    $reader = [System.Xml.XmlNodeReader]::new($stylesheet)
    $transform = [System.Xml.Xsl.XslCompiledTransform]::new()
    try {
        $transform.Load($reader)
    }
    finally {
        $reader.Dispose()
    }
    $arguments = [System.Xml.Xsl.XsltArgumentList]::new()
    $arguments.AddParam('UseDisplayName', '', $DisplayNames)
    $text = [System.Text.StringBuilder]::new()
    $writer = [System.Xml.XmlWriter]::Create($text, $transform.OutputSettings)
    try {
        $transform.Transform($XmlPath, $arguments, $writer)
    }
    finally {
        $writer.Dispose()
    }
    $document = [System.Xml.XmlDocument]::new()
    $document.PreserveWhitespace = $true
    $document.LoadXml($text.ToString())
    return ,$document
}

function Assert-RichText {
    param(
        [System.Xml.XmlDocument]$Document,
        [bool]$DisplayNames
    )

    $labels = @($Document.SelectNodes('//span[@class="field-label"] | //div[@class="section-title" or @class="collection-title"] | //table[@class="data-table"]/thead/tr/th') | ForEach-Object { $_.InnerText })
    foreach ($tag in @('p', 'span', 'h:div', 'h:p', 'h:span', 'b', 'i', 'br', 'ul', 'li')) {
        if ($labels -ccontains $tag) {
            throw "Formatting tag '$tag' became a field label."
        }
    }
    $richLabel = if ($DisplayNames) { 'Rich Notes' } else { 'my:RichNotes' }
    if (@($labels | Where-Object { $_ -ceq $richLabel }).Count -ne 1) {
        throw 'Rich-text field was not rendered as a single labeled value.'
    }
    $realDivLabel = if ($DisplayNames) { 'div' } else { 'my:div' }
    if (@($labels | Where-Object { $_ -ceq $realDivLabel }).Count -ne 1) {
        throw 'A genuine namespaced form group named div was mistaken for formatting.'
    }
    if ($null -eq $Document.SelectSingleNode('//p[span[@style="font-weight:bold"] and contains(., "First bold tail")]') -or
        $null -eq $Document.SelectSingleNode('//p[b="one" and i="two" and br]') -or
        $null -eq $Document.SelectSingleNode('//ul/li[text()="List item"]') -or
        $null -eq $Document.SelectSingleNode('//table/tr/td[@colspan="2" and text()="Rich cell"]') -or
        $null -eq $Document.SelectSingleNode('//a[@href="https://example.invalid/help" and text()="Help"]') -or
        $null -eq $Document.SelectSingleNode('//img[@src="images/photo.png" and @alt="Photo"]')) {
        throw 'Rich-text formatting, tables, links, or images were lost.'
    }
    if (-not $Document.OuterXml.Contains('</b> <i>')) {
        throw 'Whitespace between inline rich-text elements was lost.'
    }
    if ($Document.DocumentElement.InnerText -notmatch 'Before inline after' -or
        $Document.DocumentElement.InnerText -notmatch 'Prefixed content' -or
        $Document.DocumentElement.InnerText -notmatch 'Unqualified content') {
        throw 'Mixed content or alternate XHTML namespace styles were lost.'
    }
    if ($Document.SelectNodes('//*[namespace-uri()="http://www.w3.org/1999/xhtml"]').Count -ne 0) {
        throw 'XHTML was not converted to HTML elements.'
    }
    if ($Document.SelectNodes('//script | //iframe | //@onclick | //a[starts-with(@href, "javascript:")]').Count -ne 0 -or
        $Document.DocumentElement.InnerText.Contains('Hidden script')) {
        throw 'Executable rich-text markup was copied into the generated form.'
    }
    $records = $Document.SelectNodes('//div[@class="collection collection-my-Record"]/table/tbody/tr')
    if ($records.Count -ne 2) {
        throw "Expected two records, got $($records.Count)."
    }
    $headers = @($Document.SelectNodes('//div[@class="collection collection-my-Record"]/table/thead/tr/th') | ForEach-Object { $_.InnerText })
    $expectedHeaders = if ($DisplayNames) { @('Name', 'Notes', 'Details') } else { @('my:Name', 'my:Notes', 'my:Details') }
    if (($headers -join '|') -cne ($expectedHeaders -join '|')) {
        throw "Incorrect repeating-record columns: $($headers -join ', ')"
    }
    for ($index = 0; $index -lt $records.Count; $index++) {
        $prefix = if ($index -eq 0) { 'A' } else { 'B' }
        $other = if ($index -eq 0) { 'B' } else { 'A' }
        $record = $records[$index]
        if ($record.SelectNodes('td').Count -ne 3 -or
            $null -eq $record.SelectSingleNode("td[2]/p[strong='$prefix note']") -or
            $record.SelectNodes('.//div[@class="collection collection-my-Line"]/table/tbody/tr').Count -ne 2 -or
            -not $record.InnerText.Contains("$prefix line two") -or
            $record.InnerText.Contains("$other line")) {
            throw "Rich text or nested rows escaped record '$prefix'."
        }
    }
    if ($Document.SelectNodes('//div[@class="collection collection-my-RepeatedNotes"]/table/thead/tr/th').Count -ne 1 -or
        $Document.SelectNodes('//div[@class="collection collection-my-RepeatedNotes"]/table/tbody/tr/td/p').Count -ne 2) {
        throw 'Repeated rich-text values were treated as records containing formatting fields.'
    }
}

try {
    $source = Join-Path -Path $testRoot -ChildPath 'Source'
    New-Item -Path $source -ItemType Directory -Force | Out-Null
    $xmlPath = Join-Path -Path $source -ChildPath 'Sample.xml'
    @'
<?xml version="1.0" encoding="utf-8"?>
<?mso-infoPathSolution solutionVersion="1.0.0.1" href="test.xsn"?>
<my:FormData xmlns:my="urn:infopath-richtext-test" xmlns:h="http://www.w3.org/1999/xhtml">
  <my:RichNotes><div xmlns="http://www.w3.org/1999/xhtml"><p>First <span style="font-weight:bold">bold</span> tail</p><p><b>one</b> <i>two</i><br/>Next line</p><ul><li>List item</li></ul><table><tr><td colspan="2">Rich cell</td></tr></table><a href="https://example.invalid/help">Help</a><img src="images/photo.png" alt="Photo"/><script>Hidden script</script><a href="javascript:alert(1)" onclick="alert(2)">Not executable</a><iframe src="https://example.invalid/"/></div></my:RichNotes>
  <my:PrefixedNotes><h:div><h:p>Prefixed content</h:p></h:div></my:PrefixedNotes>
  <my:UnqualifiedNotes><div><p>Unqualified content</p></div></my:UnqualifiedNotes>
  <my:InlineNotes>Before <h:span>inline</h:span> after</my:InlineNotes>
  <my:ProseDate><h:p>2026-10-23</h:p></my:ProseDate>
  <my:ProseBoolean><h:span>true</h:span></my:ProseBoolean>
  <my:ActualDate>2026-10-23</my:ActualDate>
  <my:ActualBoolean>true</my:ActualBoolean>
  <my:InitiallyEmpty />
  <my:div><my:ActualValue>Genuine form group</my:ActualValue></my:div>
  <my:RepeatedNotes><h:p>Repeated rich one</h:p></my:RepeatedNotes>
  <my:RepeatedNotes><h:p>Repeated rich two</h:p></my:RepeatedNotes>
  <my:Record><my:Name>Alice</my:Name><my:Notes><h:p><h:strong>A note</h:strong></h:p></my:Notes><my:Details><my:Line><my:Text><h:p>A line one</h:p></my:Text></my:Line><my:Line><my:Text><h:p>A line two</h:p></my:Text></my:Line></my:Details></my:Record>
  <my:Record><my:Name>Bob</my:Name><my:Notes><h:p><h:strong>B note</h:strong></h:p></my:Notes><my:Details><my:Line><my:Text><h:p>B line one</h:p></my:Text></my:Line><my:Line><my:Text><h:p>B line two</h:p></my:Text></my:Line></my:Details></my:Record>
</my:FormData>
'@ | Set-Content -LiteralPath $xmlPath -Encoding UTF8
    $originalXml = [System.IO.File]::ReadAllBytes($xmlPath)

    foreach ($displayNames in @($true, $false)) {
        $output = Join-Path -Path $testRoot -ChildPath ("Output-" + $displayNames)
        & (Join-Path -Path $repoRoot -ChildPath 'RunInfoPathPipeline.ps1') `
            -SourceDirectory $source -OutputRootDirectory $output -StartStep 1
        if (-not $displayNames) {
            & (Join-Path $repoRoot 'ProcessInfopathToHTML.ps1') -InfoPathDirectory $source -MetadataCsvPath (Join-Path $output 'Reports\InfoPathVersionMap.csv') -GroupingCsvPath (Join-Path $output 'Reports\InfoPathGroupingDetailed.csv') -TemplateMappingCsvPath (Join-Path $output 'Templates\InfoPathXsltGroupMap.csv') -OutputDirectory (Join-Path $output 'Html') -UseDisplayName $false
        }
        $html = Get-Content -LiteralPath (Join-Path -Path $output -ChildPath 'Html\Sample.html') -Raw
        if ($html -notmatch '<p>2026-10-23</p>' -or $html -notmatch '<span>true</span>' -or
            $html -notmatch '10/23/2026' -or ([regex]::Matches($html, '<input\b')).Count -ne 1 -or
            -not $html.Contains('</b> <i>')) {
            throw 'HTML preprocessing changed rich-text prose or lost normal date/boolean rendering.'
        }
        $log = Get-Content -LiteralPath (Join-Path -Path $output -ChildPath 'Html\ProcessInfopathToHTML-errors.log') -Raw
        if (-not $log.Contains('No processing errors.')) {
            throw "Pipeline rendering errors: $log"
        }
        $map = @(Import-Csv -LiteralPath (Join-Path -Path $output -ChildPath 'Templates\InfoPathXsltGroupMap.csv'))
        if ($map.Count -ne 1 -or $map[0].GenerationStatus -ne 'Success') {
            throw 'Rich-text template generation failed.'
        }
        $rendered = Get-RenderedDocument -XsltPath $map[0].XsltPath -XmlPath $xmlPath -DisplayNames $displayNames
        Assert-RichText -Document $rendered -DisplayNames $displayNames

        $alternatePath = Join-Path -Path $testRoot -ChildPath 'Alternate.xml'
        $alternate = [System.Xml.XmlDocument]::new()
        $alternate.PreserveWhitespace = $true
        $alternate.Load($xmlPath)
        $alternate.SelectSingleNode('//*[local-name()="InitiallyEmpty"]').InnerXml = '<p xmlns="http://www.w3.org/1999/xhtml">New rich content</p>'
        $alternate.SelectSingleNode('//*[local-name()="PrefixedNotes"]').InnerText = 'Now plain text'
        $alternate.Save($alternatePath)
        $changed = Get-RenderedDocument -XsltPath $map[0].XsltPath -XmlPath $alternatePath -DisplayNames $displayNames
        if ($null -eq $changed.SelectSingleNode('//div[@class="field-value"]/p[text()="New rich content"]') -or
            -not $changed.DocumentElement.InnerText.Contains('Now plain text')) {
            throw 'Rich-text recognition was fixed to the sample instead of the current XML values.'
        }
    }
    if ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes($xmlPath)) -ne [Convert]::ToBase64String($originalXml)) {
        throw 'Source XML was modified.'
    }
    Write-Host 'PASS: Rich-text formatting, namespaces, whitespace, label modes, runtime values, nested sections, repeating records, and unchanged source XML.'
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
