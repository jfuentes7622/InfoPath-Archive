<#
.SYNOPSIS
    Verifies HTML label modes through the local pipeline and standalone transforms.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$testRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("InfoPathLabelTest-" + [guid]::NewGuid().ToString('N'))

function Assert-Labels {
    param(
        [string]$HtmlPath,
        [bool]$DisplayNames
    )

    $html = Get-Content -LiteralPath $HtmlPath -Raw
    $expected = if ($DisplayNames) {
        @('Form Data', 'Requestor Type', 'Contact Details', 'HTTP Response', 'Person Record', 'First Name', 'Alias Name', 'request status', 'HTTP Response Code')
    }
    else {
        @('my:FormData', 'my:RequestorType', 'my:Contact_Details', 'my:HTTPResponse', 'my:Person_Record', 'my:First_Name', 'my:Alias_Name', 'my:request_status', 'my:HTTP_ResponseCode')
    }
    $labels = @([regex]::Matches($html, '<(?:div|span) class="(?:section-title|collection-title|field-label)">([^<]*)</(?:div|span)>|<th>([^<]*)</th>') | ForEach-Object {
        $_.Groups[1].Value + $_.Groups[2].Value
    })
    foreach ($label in $expected) {
        if ($labels -cnotcontains $label) {
            throw "Missing label '$label' in '$HtmlPath'."
        }
    }
    $unexpected = if ($DisplayNames) { 'my:RequestorType' } else { 'Requestor Type' }
    if ($labels -ccontains $unexpected) {
        throw "Unexpected label '$unexpected' in '$HtmlPath'."
    }
    foreach ($value in @('Contractor', 'Accepted', 'Alice', 'Bob', 'Alpha', 'Beta', 'Pending', '200')) {
        if (-not $html.Contains(">$value<")) {
            throw "Field value '$value' was lost in '$HtmlPath'."
        }
    }
}

try {
    $source = Join-Path -Path $testRoot -ChildPath 'Source'
    New-Item -Path $source -ItemType Directory -Force | Out-Null
    $xmlPath = Join-Path -Path $source -ChildPath 'Sample.xml'
    @'
<?xml version="1.0" encoding="utf-8"?>
<?mso-infoPathSolution solutionVersion="1.0.0.1" href="test.xsn"?>
<my:FormData xmlns:my="urn:infopath-label-test">
  <my:RequestorType>Contractor</my:RequestorType>
  <my:Contact_Details><my:HTTPResponse>Accepted</my:HTTPResponse></my:Contact_Details>
  <my:Person_Record><my:First_Name>Alice</my:First_Name></my:Person_Record>
  <my:Person_Record><my:First_Name>Bob</my:First_Name></my:Person_Record>
  <my:Alias_Name>Alpha</my:Alias_Name>
  <my:Alias_Name>Beta</my:Alias_Name>
  <my:request_status>Pending</my:request_status>
  <my:HTTP_ResponseCode>200</my:HTTP_ResponseCode>
</my:FormData>
'@ | Set-Content -LiteralPath $xmlPath -Encoding UTF8

    $outputs = @{}
    foreach ($mode in @('Default', 'Internal')) {
        $output = Join-Path -Path $testRoot -ChildPath $mode
        $parameters = @{
            SourceDirectory = $source
            OutputRootDirectory = $output
        }
        & (Join-Path -Path $repoRoot -ChildPath 'RunInfoPathPipeline.ps1') @parameters -StartStep 1
        if ($mode -eq 'Internal') {
            & (Join-Path $repoRoot 'BuildXsltFromGroupingCsv.ps1') -GroupingCsvPath (Join-Path $output 'Reports\InfoPathGroupingDetailed.csv') -TemplateOutputDirectory (Join-Path $output 'Templates') -MappingCsvPath (Join-Path $output 'Templates\InfoPathXsltGroupMap.csv') -UseDisplayName $false
            & (Join-Path $repoRoot 'ProcessInfopathToHTML.ps1') -InfoPathDirectory $source -MetadataCsvPath (Join-Path $output 'Reports\InfoPathVersionMap.csv') -GroupingCsvPath (Join-Path $output 'Reports\InfoPathGroupingDetailed.csv') -TemplateMappingCsvPath (Join-Path $output 'Templates\InfoPathXsltGroupMap.csv') -OutputDirectory (Join-Path $output 'Html') -UseDisplayName $false
        }
        Assert-Labels -HtmlPath (Join-Path -Path $output -ChildPath 'Html\Sample.html') -DisplayNames ($mode -eq 'Default')
        $outputs[$mode] = $output

        $mapping = @(Import-Csv -LiteralPath (Join-Path -Path $output -ChildPath 'Templates\InfoPathXsltGroupMap.csv'))
        if ($mapping.Count -ne 1 -or $mapping[0].GenerationStatus -ne 'Success') {
            throw "Template generation failed for '$mode'."
        }
        $transform = [System.Xml.Xsl.XslCompiledTransform]::new()
        $transform.Load($mapping[0].XsltPath)
        $directHtml = Join-Path -Path $output -ChildPath 'Direct.html'
        $transform.Transform($xmlPath, $directHtml)
        Assert-Labels -HtmlPath $directHtml -DisplayNames ($mode -eq 'Default')
    }

    $displayMap = @(Import-Csv -LiteralPath (Join-Path -Path $outputs.Default -ChildPath 'Templates\InfoPathXsltGroupMap.csv'))[0]
    $internalMap = @(Import-Csv -LiteralPath (Join-Path -Path $outputs.Internal -ChildPath 'Templates\InfoPathXsltGroupMap.csv'))[0]
    [xml]$displayXslt = Get-Content -LiteralPath $displayMap.XsltPath -Raw
    [xml]$internalXslt = Get-Content -LiteralPath $internalMap.XsltPath -Raw
    $structuralXPath = '//@select[not(../@name="UseDisplayName")] | //@class'
    $displayStructure = @($displayXslt.SelectNodes($structuralXPath) | ForEach-Object { $_.Value })
    $internalStructure = @($internalXslt.SelectNodes($structuralXPath) | ForEach-Object { $_.Value })
    if (Compare-Object -ReferenceObject $displayStructure -DifferenceObject $internalStructure) {
        throw 'Label mode changed XPath bindings or CSS classes.'
    }

    foreach ($displayNames in @($true, $false)) {
        $output = $outputs.Internal
        $htmlOutput = Join-Path -Path $testRoot -ChildPath ("Override-" + $displayNames)
        & (Join-Path -Path $repoRoot -ChildPath 'ProcessInfopathToHTML.ps1') `
            -InfoPathDirectory $source `
            -MetadataCsvPath (Join-Path -Path $output -ChildPath 'Reports\InfoPathVersionMap.csv') `
            -GroupingCsvPath (Join-Path -Path $output -ChildPath 'Reports\InfoPathGroupingDetailed.csv') `
            -TemplateMappingCsvPath (Join-Path -Path $output -ChildPath 'Templates\InfoPathXsltGroupMap.csv') `
            -OutputDirectory $htmlOutput -UseDisplayName $displayNames
        Assert-Labels -HtmlPath (Join-Path -Path $htmlOutput -ChildPath 'Sample.html') -DisplayNames $displayNames
    }

    Write-Host 'PASS: Pipeline defaults, internal mode, template defaults, transform overrides, field values, and structural bindings.'
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
