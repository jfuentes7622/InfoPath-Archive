<#
.SYNOPSIS
	Generates XSLT and starter CSS from a sample InfoPath XML file.

.DESCRIPTION
	Reads a single InfoPath XML file, inspects the document structure, and builds an
	HTML-focused XSLT template plus a starter CSS file. Also writes a generation log file.

.PARAMETER InfoPathXmlPath
	Full path to the source InfoPath XML file used as the template sample.
	This file must exist and be valid XML.

.PARAMETER OutputName
	Base name used for generated output files.
	The script creates:
	- <OutputName>.xslt
	- <OutputName>.css
	- <OutputName>.log

.PARAMETER OutputDirectory
	Destination folder for generated files.
	Default: current working directory.

.EXAMPLE
	.\CreateXSLTfromInfopathFile.ps1 `
	  -InfoPathXmlPath "C:\Temp\Samples\Notice_Example.xml" `
	  -OutputName "notice_template_v1"

.EXAMPLE
	.\CreateXSLTfromInfopathFile.ps1 `
	  -InfoPathXmlPath "C:\Temp\Samples\Notice_Example.xml" `
	  -OutputName "notice_template_v1" `
	  -OutputDirectory "C:\Temp\Templates"

.NOTES
	The generated XSLT is a baseline template and may need manual refinement for
	complex nested/repeating structures.
#>

[CmdletBinding()]
param(
	[Parameter(Mandatory = $true)]
	[ValidateNotNullOrEmpty()]
	[string]$InfoPathXmlPath,

	[Parameter(Mandatory = $true)]
	[ValidateNotNullOrEmpty()]
	[string]$OutputName,

	[string]$OutputDirectory = (Get-Location).Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function New-Indent {
	param([int]$Level)
	return (" " * ($Level * 2))
}

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
		[ValidateSet("Init", "Input", "Build", "Write", "Done", "Warn")]
		[string]$Group = "Build"
	)

	$color = switch ($Group) {
		"Init" { "Cyan" }
		"Input" { "DarkCyan" }
		"Build" { "Yellow" }
		"Write" { "Green" }
		"Done" { "Magenta" }
		"Warn" { "DarkYellow" }
		default { "White" }
	}

	Write-Host "[$Group] $Message" -ForegroundColor $color
}

function Convert-ToSafeClassName {
	param([string]$Name)

	$safe = ($Name -replace "[^A-Za-z0-9_-]", "-")
	if ([string]::IsNullOrWhiteSpace($safe)) {
		return "unnamed"
	}

	return $safe
}

function Convert-ToHumanReadableLabel {
	param(
		[Parameter(Mandatory = $true)]
		[string]$ElementName
	)

	# Remove namespace prefixes (my:, d:, etc.)
	$label = $ElementName -replace '^[a-z]+:', ''
	
	# Use [regex]::Replace with full regex engine support for lookahead/lookbehind patterns
	$pattern = [regex]'(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])'
	$label = $pattern.Replace($label, ' ')
	
	# Trim any excess whitespace
	$label = $label.Trim()
	
	# Return the result (already properly capitalized by camelCase)
	return $label
}

function Convert-ToXPathLiteral {
	param([string]$Name)

	if ($Name -match "^[A-Za-z_][A-Za-z0-9_.-]*$") {
		return $Name
	}

	return "*[name()='$Name']"
}

function Get-ElementChildren {
	param([System.Xml.XmlNode]$Node)

	# Return as a single array object so callers can always use .Count safely under StrictMode.
	return ,@($Node.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })
}

function Test-LeafNode {
	param([System.Xml.XmlNode]$Node)

	$children = Get-ElementChildren -Node $Node
	return ($children.Count -eq 0)
}

function Get-RepeatedChildNames {
	param([System.Xml.XmlNode]$Node)

	$counts = @{}
	foreach ($child in (Get-ElementChildren -Node $Node)) {
		if ($counts.ContainsKey($child.Name)) {
			$counts[$child.Name]++
		}
		else {
			$counts[$child.Name] = 1
		}
	}

	$repeated = @{}
	foreach ($key in $counts.Keys) {
		if ($counts[$key] -gt 1) {
			$repeated[$key] = $true
		}
	}

	return $repeated
}

function Get-FirstChildByName {
	param(
		[System.Xml.XmlNode]$Node,
		[string]$Name
	)

	foreach ($child in (Get-ElementChildren -Node $Node)) {
		if ($child.Name -eq $Name) {
			return $child
		}
	}

	return $null
}

function Add-Line {
	param(
		[Parameter(Mandatory = $true)]
		[AllowEmptyCollection()]
		[System.Collections.Generic.List[string]]$Lines,
		[Parameter(Mandatory = $true)]
		[int]$IndentLevel,
		[Parameter(Mandatory = $true)]
		[string]$Text
	)

	$Lines.Add("$(New-Indent -Level $IndentLevel)$Text") | Out-Null
}

function Add-ValueRenderer {
	param(
		[Parameter(Mandatory = $true)]
		[AllowEmptyCollection()]
		[System.Collections.Generic.List[string]]$Lines,
		[Parameter(Mandatory = $true)]
		[int]$IndentLevel,
		[Parameter(Mandatory = $true)]
		[string]$SelectExpression
	)

	Add-Line -Lines $Lines -IndentLevel $IndentLevel -Text "<xsl:choose>"
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text ('<xsl:when test="starts-with({0}, ''ATTACHMENT::'')">' -f $SelectExpression)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text '<a target="_blank">'
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text ('<xsl:attribute name="href"><xsl:value-of select="substring-after({0}, ''ATTACHMENT::'')" /></xsl:attribute>' -f $SelectExpression)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text ('<xsl:value-of select="substring-after({0}, ''ATTACHMENT::'')" />' -f $SelectExpression)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text '</a>'
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "</xsl:when>"
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "<xsl:otherwise>"
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text ('<xsl:value-of select="{0}" />' -f $SelectExpression)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "</xsl:otherwise>"
	Add-Line -Lines $Lines -IndentLevel $IndentLevel -Text "</xsl:choose>"
}

function Add-ElementPanel {
	param(
		[Parameter(Mandatory = $true)]
		[System.Xml.XmlNode]$Node,
		[Parameter(Mandatory = $true)]
		[string]$CurrentXPath,
		[Parameter(Mandatory = $true)]
		[AllowEmptyCollection()]
		[System.Collections.Generic.List[string]]$Lines,
		[Parameter(Mandatory = $true)]
		[AllowEmptyCollection()]
		[System.Collections.Generic.List[string]]$LogEntries,
		[Parameter(Mandatory = $true)]
		[int]$IndentLevel
	)

	Write-Stage -Group Build -Message ("Rendering element '{0}' at XPath '{1}'" -f $Node.Name, $CurrentXPath)

	$safeClass = Convert-ToSafeClassName -Name $Node.Name
	$humanLabel = Convert-ToHumanReadableLabel -ElementName $Node.Name
	Add-Line -Lines $Lines -IndentLevel $IndentLevel -Text "<div class=`"section section-$safeClass`">"
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "<div class=`"section-title`">$humanLabel</div>"

	$children = Get-ElementChildren -Node $Node
	if ($children.Count -eq 0) {
		Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "<div class=`"field-row`">"
		$humanLabel = Convert-ToHumanReadableLabel -ElementName $Node.Name
		Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "<span class=`"field-label`">$humanLabel</span>"
		Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text '<span class="field-value">'
		Add-ValueRenderer -Lines $Lines -IndentLevel ($IndentLevel + 3) -SelectExpression $CurrentXPath
		Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "</span>"
		Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "</div>"
		Add-Line -Lines $Lines -IndentLevel $IndentLevel -Text "</div>"
		return
	}

	$repeatedNames = Get-RepeatedChildNames -Node $Node
	$renderedRepeated = @{}

	foreach ($child in $children) {
		$childXPathPart = Convert-ToXPathLiteral -Name $child.Name
		$childXPath = "$CurrentXPath/$childXPathPart"

		if ($repeatedNames.ContainsKey($child.Name)) {
			if ($renderedRepeated.ContainsKey($child.Name)) {
				continue
			}

			$renderedRepeated[$child.Name] = $true
			$sampleNode = Get-FirstChildByName -Node $Node -Name $child.Name

			if ($null -eq $sampleNode) {
				Write-Stage -Group Warn -Message ("Skipping repeated node '{0}' under '{1}' because no representative sample was found." -f $child.Name, $Node.Name)
				Write-LogEntry -LogEntries $LogEntries -Message "Unable to find representative sample for repeated node '$($child.Name)' under '$($Node.Name)'."
				continue
			}

			$sampleChildren = Get-ElementChildren -Node $sampleNode
			$humanLabel = Convert-ToHumanReadableLabel -ElementName $child.Name
			Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "<div class=`"collection collection-$(Convert-ToSafeClassName -Name $child.Name)`">"
			Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "<div class=`"collection-title`">$humanLabel</div>"

			if ($sampleChildren.Count -eq 0) {
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "<table class=`"data-table`">"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text "<thead><tr><th>Value</th></tr></thead>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text "<tbody>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 4) -Text ('<xsl:for-each select="{0}">' -f $childXPath)
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 5) -Text '<tr><td><xsl:value-of select="." /></td></tr>'
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 4) -Text "</xsl:for-each>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text "</tbody>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "</table>"
			}
			else {
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "<table class=`"data-table`">"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text "<thead>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 4) -Text "<tr>"

				foreach ($headerNode in $sampleChildren) {
					$headerLabel = Convert-ToHumanReadableLabel -ElementName $headerNode.Name
					Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 5) -Text "<th>$headerLabel</th>"
				}

				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 4) -Text "</tr>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text "</thead>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text "<tbody>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 4) -Text ('<xsl:for-each select="{0}">' -f $childXPath)
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 5) -Text "<tr>"

				foreach ($valueNode in $sampleChildren) {
					$valuePath = Convert-ToXPathLiteral -Name $valueNode.Name
					Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 6) -Text "<td>"
					Add-ValueRenderer -Lines $Lines -IndentLevel ($IndentLevel + 7) -SelectExpression $valuePath
					Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 6) -Text "</td>"
				}

				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 5) -Text "</tr>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 4) -Text "</xsl:for-each>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text "</tbody>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "</table>"

				foreach ($nestedCandidate in $sampleChildren) {
					if (-not (Test-LeafNode -Node $nestedCandidate)) {
						Write-Stage -Group Warn -Message ("Nested structure '{0}/{1}' rendered as table text; consider manual refinement." -f $child.Name, $nestedCandidate.Name)
						Write-LogEntry -LogEntries $LogEntries -Message "Nested complex structure '$($child.Name)/$($nestedCandidate.Name)' rendered as table cell text. Consider manual refinement."
						Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "<!-- TODO: Nested structure '$($child.Name)/$($nestedCandidate.Name)' may need custom rendering. -->"
					}
				}
			}

			Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "</div>"
			continue
		}

		if (Test-LeafNode -Node $child) {
			Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "<div class=`"field-row`">"
			$childLabel = Convert-ToHumanReadableLabel -ElementName $child.Name
			Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "<span class=`"field-label`">$childLabel</span>"
			Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text '<span class="field-value">'
			Add-ValueRenderer -Lines $Lines -IndentLevel ($IndentLevel + 3) -SelectExpression $childXPath
			Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "</span>"
			Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "</div>"
		}
		else {
			Add-ElementPanel -Node $child -CurrentXPath $childXPath -Lines $Lines -LogEntries $LogEntries -IndentLevel ($IndentLevel + 1)
		}
	}

	Add-Line -Lines $Lines -IndentLevel $IndentLevel -Text "</div>"
}

if (-not (Test-Path -LiteralPath $InfoPathXmlPath -PathType Leaf)) {
	throw "InfoPathXmlPath does not exist: $InfoPathXmlPath"
}

Write-Stage -Group Init -Message "Starting InfoPath to XSLT generation."
Write-Stage -Group Input -Message "Input document: $InfoPathXmlPath"
Write-Stage -Group Input -Message "Output base name: $OutputName"
Write-Stage -Group Init -Message "Validating output location."
if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
	Write-Stage -Group Write -Message "Creating output directory: $OutputDirectory"
	New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
}

Write-Stage -Group Input -Message "Loading InfoPath XML from disk."
$streamReader = [System.IO.StreamReader]::new($InfoPathXmlPath, $true)
$rawXml = $streamReader.ReadToEnd()
$streamReader.Close()
$streamReader.Dispose()

$xmlDocument = [System.Xml.XmlDocument]::new()
$xmlDocument.XmlResolver = $null
$xmlSettings = [System.Xml.XmlReaderSettings]::new()
$xmlSettings.IgnoreComments = $true
$xmlSettings.IgnoreWhitespace = $true
$xmlSettings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
$xmlReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($rawXml), $xmlSettings)
try {
	$xmlDocument.Load($xmlReader)
}
finally {
	$xmlReader.Dispose()
}

if ($null -eq $xmlDocument.DocumentElement) {
	throw "The XML file does not contain a document element: $InfoPathXmlPath"
}

$baseOutputName = [System.IO.Path]::GetFileName($OutputName)
if ([string]::IsNullOrWhiteSpace($baseOutputName)) {
	throw "OutputName must include at least one valid character."
}

$xsltPath = Join-Path -Path $OutputDirectory -ChildPath "$baseOutputName.xslt"
$cssPath = Join-Path -Path $OutputDirectory -ChildPath "$baseOutputName.css"
$logPath = Join-Path -Path $OutputDirectory -ChildPath "$baseOutputName.log"

$logEntries = New-Object System.Collections.Generic.List[string]
Write-LogEntry -LogEntries $logEntries -Message "Started template generation."
Write-LogEntry -LogEntries $logEntries -Message "Input XML: $InfoPathXmlPath"
Write-LogEntry -LogEntries $logEntries -Message "Output XSLT: $xsltPath"
Write-LogEntry -LogEntries $logEntries -Message "Output CSS: $cssPath"
Write-Stage -Group Write -Message "XSLT output file: $xsltPath"
Write-Stage -Group Write -Message "CSS output file: $cssPath"
Write-Stage -Group Build -Message "Preparing XSLT template structure."

$xsltLines = New-Object System.Collections.Generic.List[string]
$cssRelativePath = [System.IO.Path]::GetFileName($cssPath)
$rootName = $xmlDocument.DocumentElement.Name
$rootXPath = Convert-ToXPathLiteral -Name $rootName
Write-Stage -Group Build -Message "Detected root element: $rootName"

Add-Line -Lines $xsltLines -IndentLevel 0 -Text "<?xml version=`"1.0`" encoding=`"utf-8`"?>"
Add-Line -Lines $xsltLines -IndentLevel 0 -Text "<xsl:stylesheet version=`"1.0`" xmlns:xsl=`"http://www.w3.org/1999/XSL/Transform`">"
Add-Line -Lines $xsltLines -IndentLevel 1 -Text "<xsl:output method=`"html`" indent=`"yes`" encoding=`"utf-8`" />"
Add-Line -Lines $xsltLines -IndentLevel 1 -Text "<xsl:param name=`"sourceFilePath`" />"
Add-Line -Lines $xsltLines -IndentLevel 1 -Text "<xsl:param name=`"sourceFileName`" />"
Add-Line -Lines $xsltLines -IndentLevel 1 -Text "<xsl:template match=`"/`">"
Add-Line -Lines $xsltLines -IndentLevel 2 -Text "<html>"
Add-Line -Lines $xsltLines -IndentLevel 3 -Text "<head>"
Add-Line -Lines $xsltLines -IndentLevel 4 -Text "<meta charset=`"utf-8`" />"
Add-Line -Lines $xsltLines -IndentLevel 4 -Text "<title>InfoPath Form Template</title>"
Add-Line -Lines $xsltLines -IndentLevel 4 -Text "<link rel=`"stylesheet`" type=`"text/css`" href=`"$cssRelativePath`" />"
Add-Line -Lines $xsltLines -IndentLevel 3 -Text "</head>"
Add-Line -Lines $xsltLines -IndentLevel 3 -Text "<body>"
Add-Line -Lines $xsltLines -IndentLevel 4 -Text "<div class=`"page`">"
Add-Line -Lines $xsltLines -IndentLevel 5 -Text "<div class=`"page-header`">"
Add-Line -Lines $xsltLines -IndentLevel 6 -Text "<h1>$rootName</h1>"
Add-Line -Lines $xsltLines -IndentLevel 6 -Text "<div class=`"source-meta`">"
Add-Line -Lines $xsltLines -IndentLevel 7 -Text "<span>Source File:</span> <xsl:value-of select=`"`$sourceFileName`" />"
Add-Line -Lines $xsltLines -IndentLevel 7 -Text "<br />"
Add-Line -Lines $xsltLines -IndentLevel 7 -Text "<span>Source Path:</span> <xsl:value-of select=`"`$sourceFilePath`" />"
Add-Line -Lines $xsltLines -IndentLevel 6 -Text "</div>"
Add-Line -Lines $xsltLines -IndentLevel 5 -Text "</div>"

Add-ElementPanel -Node $xmlDocument.DocumentElement -CurrentXPath $rootXPath -Lines $xsltLines -LogEntries $logEntries -IndentLevel 5

Add-Line -Lines $xsltLines -IndentLevel 4 -Text "</div>"
Add-Line -Lines $xsltLines -IndentLevel 3 -Text "</body>"
Add-Line -Lines $xsltLines -IndentLevel 2 -Text "</html>"
Add-Line -Lines $xsltLines -IndentLevel 1 -Text "</xsl:template>"
Add-Line -Lines $xsltLines -IndentLevel 0 -Text "</xsl:stylesheet>"

Write-Stage -Group Build -Message "Generating starter CSS."
$cssContent = @'
body {
  margin: 0;
  padding: 0;
  background: #f3f5f7;
  color: #1f2937;
  font-family: "Segoe UI", Tahoma, sans-serif;
}

.page {
  max-width: 1200px;
  margin: 24px auto;
  padding: 20px;
  background: #ffffff;
  border: 1px solid #d6dde5;
  border-radius: 8px;
}

.page-header {
  border-bottom: 2px solid #d6dde5;
  margin-bottom: 16px;
  padding-bottom: 12px;
}

.page-header h1 {
  margin: 0 0 8px 0;
  font-size: 28px;
}

.source-meta {
  color: #4b5563;
  font-size: 12px;
}

.section {
  margin: 14px 0;
  padding: 12px;
  border: 1px solid #d1d9e0;
  border-radius: 6px;
  background: #f9fbfc;
}

.section-title,
.collection-title {
  font-weight: 600;
  color: #0f172a;
  margin-bottom: 8px;
}

.field-row {
  display: grid;
  grid-template-columns: 280px 1fr;
  gap: 8px;
  margin-bottom: 6px;
}

.field-label {
  font-weight: 600;
}

.field-value {
  word-break: break-word;
}

.collection {
  margin-top: 10px;
}

.data-table {
  width: 100%;
  border-collapse: collapse;
}

.data-table th,
.data-table td {
  border: 1px solid #cfd8e3;
  padding: 6px 8px;
  text-align: left;
  vertical-align: top;
}

.data-table th {
  background: #e8eef5;
}
'@

Write-Stage -Group Write -Message "Writing XSLT and CSS files."
$xsltLines | Set-Content -LiteralPath $xsltPath -Encoding UTF8
$cssContent | Set-Content -LiteralPath $cssPath -Encoding UTF8

Write-LogEntry -LogEntries $logEntries -Message "Template generation complete."
Write-Stage -Group Write -Message "Writing generation log file."
$logEntries | Set-Content -LiteralPath $logPath -Encoding UTF8

Write-Stage -Group Done -Message "Template generation completed successfully."

Write-Host "Created XSLT: $xsltPath" -ForegroundColor Green
Write-Host "Created CSS: $cssPath" -ForegroundColor Green
Write-Host "Created log: $logPath" -ForegroundColor Green

# Output actual created file paths for validation by caller
$result = @{
    XsltPath = $xsltPath
    CssPath  = $cssPath
    LogPath  = $logPath
    Success  = $true
}

# Write result as a special marker line for caller to parse
Write-Host "##TEMPLATE_GENERATION_RESULT##$(ConvertTo-Json -InputObject $result -Compress)##END_RESULT##" -ForegroundColor Magenta

