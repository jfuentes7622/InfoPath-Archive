<#
.SYNOPSIS
    Verifies the customer's HTML field-rendering rules and unchanged source data.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path -Path $PSScriptRoot -ChildPath 'HtmlRendering.ps1')
$testRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("InfoPath Customer Fields-" + [guid]::NewGuid().ToString('N'))

function Assert-TextValue {
    param([string]$Html, [string]$Value)

    if ($Html -notmatch ('>\s*' + [regex]::Escape($Value) + '\s*<')) {
        throw "Missing displayed value '$Value'."
    }
}

try {
    $source = Join-Path -Path $testRoot -ChildPath 'Source Forms'
    $files = Join-Path -Path $source -ChildPath 'Files'
    New-Item -Path $files -ItemType Directory -Force | Out-Null
    [System.IO.File]::WriteAllText((Join-Path -Path $files -ChildPath 'Invoice.txt'), 'Referenced attachment content')
    $uploadBytes = [System.Text.Encoding]::ASCII.GetBytes("%PDF-1.7`nExample upload`n%%EOF")
    $xmlPath = Join-Path -Path $source -ChildPath 'Sample.xml'
    $xml = @'
<?xml version="1.0" encoding="utf-8"?>
<?mso-infoPathSolution solutionVersion="1.0.0.1" href="test.xsn"?>
<my:FormData xmlns:my="urn:example:infopath" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <my:CalculatedResult>true</my:CalculatedResult>
  <my:TextDate>2026-10-21</my:TextDate>
  <my:Identifier>001250</my:Identifier>
  <my:Checkbox>true</my:Checkbox>
  <my:Choices>Red;#Blue;#Green</my:Choices>
  <my:Amount>1500.5</my:Amount>
  <my:EventDateTime>2026-10-21T14:30:00-04:00</my:EventDateTime>
  <my:DateOnly>2026-10-21</my:DateOnly>
  <my:Status>1</my:Status>
  <my:EmptyField />
  <my:NullField xsi:nil="true" />
  <my:Marker>ATTACHMENT::docs/Invoice.pdf</my:Marker>
  <my:FileReference>Files\Invoice.txt</my:FileReference>
  <my:RawUpload>UPLOAD_BASE64</my:RawUpload>
  <my:DocumentLink><my:Url>https://example.invalid/document</my:Url><my:Description>View Document</my:Description></my:DocumentLink>
  <my:PlainLink>https://example.invalid/help</my:PlainLink>
  <my:DescribedLink>https://example.invalid/info, More information</my:DescribedLink>
  <my:Lookup>15;#New York</my:Lookup>
  <my:MultipleLookup>15;#New York;#42;#Los Angeles</my:MultipleLookup>
  <my:Multiline>First line
Second line</my:Multiline>
  <my:MultipleCheckboxes>Option1=true; Option2=false</my:MultipleCheckboxes>
  <my:Count>1250</my:Count>
  <my:Rate>0.75</my:Rate>
  <my:Person>DOMAIN\jsmith</my:Person>
  <my:People><my:Person><my:DisplayName>Stored Display Name</my:DisplayName><my:AccountId>DOMAIN\stored</my:AccountId></my:Person><my:Person><my:DisplayName /><my:AccountId>DOMAIN\unknown</my:AccountId></my:Person></my:People>
  <my:RadioChoice>Y</my:RadioChoice>
  <my:Record><my:Name>Alice</my:Name><my:Amount>1500.5</my:Amount></my:Record>
  <my:Record><my:Name>Bob</my:Name><my:Amount>20</my:Amount></my:Record>
  <my:RichText><p xmlns="http://www.w3.org/1999/xhtml">Formatted <b>content</b></p></my:RichText>
  <my:Signature><my:Signer>DOMAIN\jsmith</my:Signer><my:SignedDate>2026-10-21</my:SignedDate></my:Signature>
  <my:SingleLine>John Smith</my:SingleLine>
  <my:UnknownPerson>DOMAIN\unknown</my:UnknownPerson>
</my:FormData>
'@
    $xml.Replace('UPLOAD_BASE64', [Convert]::ToBase64String($uploadBytes)) | Set-Content -LiteralPath $xmlPath -Encoding UTF8
    $originalXml = [System.IO.File]::ReadAllBytes($xmlPath)
    $rulesPath = Join-Path -Path $testRoot -ChildPath 'Field Rules.json'
    $rules = Get-Content -LiteralPath (Join-Path -Path $repoRoot -ChildPath 'Examples\HtmlFieldRules.example.json') -Raw | ConvertFrom-Json
    $rules.Fields = @(
        $rules.Fields | Where-Object { $_.XPath -ne '//my:Person' }
        [pscustomobject]@{ XPath = '/my:FormData/my:Person'; Type = 'Person'; Values = [pscustomobject]@{ 'DOMAIN\jsmith' = 'John Smith' } }
        [pscustomobject]@{ XPath = '//my:People'; Type = 'Person' }
        [pscustomobject]@{ XPath = '//my:UnknownPerson'; Type = 'Person' }
        [pscustomobject]@{ XPath = '//my:TextDate'; Type = 'Text' }
    )
    $rules | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $rulesPath -Encoding UTF8
    $output = Join-Path -Path $testRoot -ChildPath 'Mapped Output'
    & (Join-Path -Path $repoRoot -ChildPath 'RunInfoPathPipeline.ps1') `
        -SourceDirectory $source -OutputRootDirectory $output -FieldRulesPath $rulesPath -PrintToPdf 'true'
    $htmlPath = Join-Path -Path $output -ChildPath 'Html\Sample.html'
    $html = Get-Content -LiteralPath $htmlPath -Raw
    foreach ($value in @('$1,500.50', '$20.00', '1,250', '75%', '10/21/2026', '10/21/2026 14:30',
        'Red, Blue, Green', 'New York', 'New York, Los Angeles', 'Approved', 'Yes',
        'John Smith', 'John Smith - 10/21/2026', 'Stored Display Name; DOMAIN\unknown',
        'DOMAIN\unknown', '2026-10-21', '001250', 'true', 'Invoice.pdf', 'Invoice.txt', 'View Document', 'More information')) {
        Assert-TextValue -Html $html -Value $value
    }
    if ($html -notmatch '<a href="https://example.invalid/document">View Document</a>' -or
        $html -notmatch '<a href="https://example.invalid/help">https://example.invalid/help</a>' -or
        $html -notmatch 'First line\r?\nSecond line' -or
        $html -notmatch '<b>content</b>' -or
        ([regex]::Matches($html, '<input\b')).Count -ne 3 -or
        ([regex]::Matches($html, '\bchecked(?:=|\s|>)')).Count -ne 2 -or
        $html -notmatch 'Option1' -or $html -notmatch 'Option2') {
        throw 'Hyperlinks, multiline/rich text, or checkbox selections were not rendered correctly.'
    }
    foreach ($field in @('Empty Field', 'Null Field')) {
        if ($html -notmatch ('<span class="field-label">' + $field + '</span>\s*<div class="field-value">\s*(?:<span class="plain-text">\s*</span>)?\s*</div>')) {
            throw "Empty field '$field' was not blank."
        }
    }
    $cssPath = @(Get-ChildItem -LiteralPath (Join-Path -Path $output -ChildPath 'Html') -Filter '*.css')[0].FullName
    if ((Get-Content -LiteralPath $cssPath -Raw) -notmatch '\.plain-text\s*\{\s*white-space:\s*pre-wrap;') {
        throw 'Multiline text does not preserve line breaks in the browser.'
    }
    $attachmentDirectory = Join-Path -Path $output -ChildPath 'Html\Sample_attachments'
    if ([System.IO.File]::ReadAllText((Join-Path -Path $attachmentDirectory -ChildPath 'Invoice.txt')) -ne 'Referenced attachment content' -or
        [Convert]::ToBase64String([System.IO.File]::ReadAllBytes((Join-Path -Path $attachmentDirectory -ChildPath 'Invoice.pdf'))) -ne [Convert]::ToBase64String($uploadBytes)) {
        throw 'File-reference or raw-base64 attachments were not archived intact.'
    }
    $pdfPath = Join-Path -Path $output -ChildPath 'Html\Sample.pdf'
    if (-not (Test-Path -LiteralPath $pdfPath) -or (Get-Item -LiteralPath $pdfPath).Length -lt 1000) {
        throw 'Customer-formatted HTML was not printed to PDF.'
    }
    $log = Get-Content -LiteralPath (Join-Path -Path $output -ChildPath 'Html\ProcessInfopathToHTML-errors.log') -Raw
    if (-not $log.Contains('No processing errors.')) { throw "Pipeline rendering errors: $log" }
    $main = @(Get-ChildItem -LiteralPath (Join-Path -Path $output -ChildPath 'Reports\Normalized') -Filter 'Main.csv' -Recurse)[0]
    $csv = Get-Content -LiteralPath $main.FullName -Raw
    foreach ($value in @('1500.5', '0.75', '15;#New York', '2026-10-21T14:30:00-04:00', 'DOMAIN\jsmith')) {
        if (-not $csv.Contains($value)) { throw "Normalized CSV value '$value' was changed." }
    }
    if ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes($xmlPath)) -ne [Convert]::ToBase64String($originalXml)) {
        throw 'Source XML was changed.'
    }

    $standaloneOutput = Join-Path -Path $testRoot -ChildPath 'Standalone'
    ($rules.Fields | Where-Object { $_.Type -eq 'Dropdown' }).Values.'1' = 'Accepted'
    $rules | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $rulesPath -Encoding UTF8
    & (Join-Path -Path $repoRoot -ChildPath 'ProcessInfopathToHTML.ps1') `
        -InfoPathDirectory $source -OutputDirectory $standaloneOutput `
        -MetadataCsvPath (Join-Path -Path $output -ChildPath 'Reports\InfoPathVersionMap.csv') `
        -GroupingCsvPath (Join-Path -Path $output -ChildPath 'Reports\InfoPathGroupingDetailed.csv') `
        -TemplateMappingCsvPath (Join-Path -Path $output -ChildPath 'Templates\InfoPathXsltGroupMap.csv') `
        -FieldRulesPath $rulesPath -DateFormat 'dd.MM.yyyy' -DateTimeFormat 'dd.MM.yyyy HH:mm:ss zzz'
    $standaloneHtml = Get-Content -LiteralPath (Join-Path -Path $standaloneOutput -ChildPath 'Sample.html') -Raw
    foreach ($value in @('Accepted', '21.10.2026', '21.10.2026 14:30:00 -04:00', 'John Smith - 21.10.2026')) {
        Assert-TextValue -Html $standaloneHtml -Value $value
    }

    $configuration = Import-HtmlFieldRules
    [xml]$unmapped = '<Form><Amount>1500.5</Amount><Rate>0.75</Rate><Account>DOMAIN\jsmith</Account><Choices>Red;#Blue;#Green</Choices><Lookup>15;#New York</Lookup><When>2026-10-21T14:30:00</When></Form>'
    Convert-HtmlFieldValues -Document $unmapped -Configuration $configuration
    foreach ($pair in @(@('Amount', '1500.5'), @('Rate', '0.75'), @('Account', 'DOMAIN\jsmith'), @('Choices', 'Red, Blue, Green'), @('Lookup', 'New York'), @('When', '10/21/2026 14:30'))) {
        if ($unmapped.SelectSingleNode('/Form/' + $pair[0]).InnerText -cne $pair[1]) {
            throw "Unexpected automatic rendering of '$($pair[0])'."
        }
    }
    foreach ($invalidJson in @(
        '{"Fields":[{"XPath":"//Amount","Type":"Unknown"}]}',
        '{"Fields":[{"XPath":"//Amount","Type":"Currency","Format":"Q"}]}',
        '{"Fields":[{"XPath":"//Amount","Type":"Currency","Formt":"C2"}]}',
        '{"Fields":[{"XPath":"//Signature","Type":"Signature"}]}',
        '{"Fields":[{"XPath":"//p:Amount","Type":"Number"}]}',
        '{"Fields":[{"XPath":"//Upload","Type":"Attachment","Encoding":"base64","FileName":"..\secret.txt"}]}'
    )) {
        $invalidPath = Join-Path -Path $testRoot -ChildPath 'Invalid.json'
        $invalidJson | Set-Content -LiteralPath $invalidPath -Encoding UTF8
        $rejected = $false
        try { $null = Import-HtmlFieldRules -Path $invalidPath }
        catch { $rejected = $true }
        if (-not $rejected) { throw "Invalid rule configuration was accepted: $invalidJson" }
    }
    $badRule = [pscustomobject]@{ XPath = '/Form/Amount'; Type = 'Currency' }
    $configuration.Fields = @($badRule)
    [xml]$badValue = '<Form><Amount>not money</Amount></Form>'
    $rejected = $false
    try { Convert-HtmlFieldValues -Document $badValue -Configuration $configuration }
    catch { $rejected = $_.Exception.Message.Contains('Invalid Currency value') }
    if (-not $rejected) { throw 'Invalid typed currency value did not produce an explicit error.' }
    $configuration.Culture = [System.Globalization.CultureInfo]::GetCultureInfo('fr-FR')
    [xml]$localized = '<Form><Amount>1500.5</Amount></Form>'
    Convert-HtmlFieldValues -Document $localized -Configuration $configuration
    if ($localized.SelectSingleNode('/Form/Amount').InnerText -cne ([decimal]1500.5).ToString('C2', $configuration.Culture)) {
        throw 'Configured numeric culture was not applied.'
    }
    $configuration.Fields = @([pscustomobject]@{ XPath = '/Form/Flag'; Type = 'Boolean' })
    [xml]$numericBoolean = '<Form><Flag>1</Flag></Form>'
    Convert-HtmlFieldValues -Document $numericBoolean -Configuration $configuration
    if ($numericBoolean.SelectSingleNode('/Form/Flag').InnerText -cne 'true') {
        throw 'Explicit Boolean field rule did not recognize numeric 1.'
    }
    $configuration.Fields = @(
        [pscustomobject]@{ XPath = '/Form/Flag'; Type = 'Text' }
        [pscustomobject]@{ XPath = '//Flag'; Type = 'Boolean' }
    )
    $rejected = $false
    try { Convert-HtmlFieldValues -Document $numericBoolean -Configuration $configuration }
    catch { $rejected = $_.Exception.Message.Contains('Overlapping HTML field rules') }
    if (-not $rejected) { throw 'Overlapping rules were not rejected.' }
    $configuration.Fields = @([pscustomobject]@{ XPath = '/Form/Upload'; Type = 'Attachment' })
    [System.IO.File]::WriteAllText((Join-Path -Path $testRoot -ChildPath 'Outside.txt'), 'Must not be copied')
    [xml]$outsideReference = '<Form><Upload>..\Outside.txt</Upload></Form>'
    $rejected = $false
    try {
        Export-HtmlFieldAttachments -Document $outsideReference -Configuration $configuration -SourceFileDirectory $source `
            -SourceRoot $source -AttachmentOutputDirectory $attachmentDirectory -RelativePrefix 'Sample_attachments'
    }
    catch { $rejected = $_.Exception.Message.Contains('must remain within the InfoPath source directory') }
    if (-not $rejected) { throw 'An attachment outside the source root was not rejected.' }
    [xml]$unsafeLink = '<Form><Link /></Form>'
    $rejected = $false
    try { Set-HtmlHyperlink -Node $unsafeLink.DocumentElement.FirstChild -Url 'javascript:alert(1)' -Label 'Invalid' }
    catch { $rejected = $_.Exception.Message.Contains('Unsupported or invalid HTML hyperlink') }
    if (-not $rejected) { throw 'An executable hyperlink scheme was not rejected.' }
    Set-HtmlHyperlink -Node $unsafeLink.DocumentElement.FirstChild -Url 'attachments/Invoice #1.pdf' -Label 'Invoice #1.pdf' -Attachment
    if ($unsafeLink.DocumentElement.FirstChild.GetAttribute('href', 'urn:infopath-archive:display') -cne 'attachments/Invoice%20%231.pdf') {
        throw 'Attachment URL-special characters were not escaped.'
    }
    $lookupValues = Get-HtmlChoiceValues -Value 'New York;#Los Angeles' -StripLookupIds
    if (($lookupValues -join ', ') -cne 'New York, Los Angeles') {
        throw 'Lookup values without numeric IDs lost a selected value.'
    }
    Write-Host 'PASS: Customer field formats, mappings, choices/lookups, links, attachments, signatures, multiline/rich text, date/time offsets, blanks, PDF output, standalone overrides, validation, and unchanged XML/CSV values.'
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
