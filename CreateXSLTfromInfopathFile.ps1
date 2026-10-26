<#
.SYNOPSIS
	Generates XSLT and starter CSS from a sample InfoPath XML file.

.DESCRIPTION
	Reads a single InfoPath XML file, inspects the document structure, and builds an
	HTML-focused XSLT template plus a starter CSS file. Also writes a generation log file.
	Boolean values are displayed as disabled checkboxes, including in repeating tables.
	XHTML formatting is preserved as rich-text field content rather than form structure.

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

.PARAMETER UseDisplayName
	Prefer matched XSN captions, falling back to readable XML-name labels.
	Set to $false for exact internal XML names, including namespace prefixes.
	Only visible labels change; XPath and CSS identifiers remain internal.
	Default: $true.

.PARAMETER FieldRulesPath
	Optional JSON HTML field definitions, used to recognize structured display fields.
	Supply the same file to ProcessInfopathToHTML.ps1 when rendering.

.PARAMETER XsnInventoryCsvPath
	Processed XSN inventory for sample captions. Matching uses the form identity
	and exact solutionVersion, or the nearest higher version of the same form.
	Generated templates also accept per-file captions from the HTML renderer.

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

	[string]$OutputDirectory = (Get-Location).Path,

	[bool]$UseDisplayName = $true,

	[string]$FieldRulesPath,

	[string]$XsnInventoryCsvPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path -Path $PSScriptRoot -ChildPath 'Helper\HtmlRendering.ps1')
. (Join-Path -Path $PSScriptRoot -ChildPath 'Helper\XsnTemplates.ps1')
$fieldRules = Import-HtmlFieldRules -Path $FieldRulesPath

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
	$label = $label.Replace('_', ' ')
	
	# Use [regex]::Replace with full regex engine support for lookahead/lookbehind patterns
	$pattern = [regex]'(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])'
	$label = $pattern.Replace($label, ' ')
	
	# Trim any excess whitespace
	$label = $label.Trim()
	
	# Return the result (already properly capitalized by camelCase)
	return $label
}

function Convert-ToFieldLabel {
	param(
		[Parameter(Mandatory = $true)]
		[string]$ElementName,

		[System.Xml.XmlElement]$Node,

		[string]$SelectExpression
	)

	$displayLabel = [System.Security.SecurityElement]::Escape((Convert-ToHumanReadableLabel -ElementName $ElementName))
	$embeddedLabel = [System.Security.SecurityElement]::Escape($Node.GetAttribute('label', 'urn:infopath-archive:display'))
	$labelXPath = [System.Security.SecurityElement]::Escape("($SelectExpression)/@archive:label")
	$internalLabel = [System.Security.SecurityElement]::Escape($ElementName)
	$embeddedChoice = if ($embeddedLabel) {
		'<xsl:when test="$UseEmbeddedXsnLabels"><xsl:text>{0}</xsl:text></xsl:when>' -f $embeddedLabel
	} else { '' }
	return '<xsl:choose><xsl:when test="$UseDisplayName"><xsl:choose><xsl:when test="{0}"><xsl:value-of select="{0}" /></xsl:when>{1}<xsl:otherwise><xsl:text>{2}</xsl:text></xsl:otherwise></xsl:choose></xsl:when><xsl:otherwise><xsl:text>{3}</xsl:text></xsl:otherwise></xsl:choose>' -f $labelXPath, $embeddedChoice, $displayLabel, $internalLabel
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

	if ($Node -is [System.Xml.XmlElement] -and $Node.HasAttribute('kind', 'urn:infopath-archive:display')) {
		return ,@()
	}

	# Return as a single array object so callers can always use .Count safely under StrictMode.
	return ,@($Node.ChildNodes | Where-Object {
		$_.NodeType -eq [System.Xml.XmlNodeType]::Element -and -not (Test-HtmlFormattingNode -Node $_)
	})
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
		$childName = $child.get_Name()
		if ($counts.ContainsKey($childName)) {
			$counts[$childName]++
		}
		else {
			$counts[$childName] = 1
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
		if ($child.get_Name() -eq $Name) {
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
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text ('<xsl:when test="{0}/@archive:kind = ''link''">' -f $SelectExpression)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text ('<a><xsl:attribute name="href"><xsl:value-of select="{0}/@archive:href" /></xsl:attribute><xsl:value-of select="{0}" /></a>' -f $SelectExpression)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text '</xsl:when>'
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text ('<xsl:when test="{0}/@archive:kind = ''selections''">' -f $SelectExpression)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text ('<xsl:for-each select="{0}/archive:option"><label class="checkbox-selection"><input type="checkbox" disabled="disabled"><xsl:if test="@selected = ''true''"><xsl:attribute name="checked">checked</xsl:attribute></xsl:if></input><xsl:text> </xsl:text><xsl:value-of select="." /></label></xsl:for-each>' -f $SelectExpression)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text '</xsl:when>'
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text ('<xsl:when test="{0}/@archive:kind = ''text''"><span class="plain-text"><xsl:value-of select="{0}" /></span></xsl:when>' -f $SelectExpression)
	$formattingXPath = Get-HtmlFormattingXPath
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text ('<xsl:when test="{0}/@archive:kind = ''rich-text'' or {0}/*[{1}]">' -f $SelectExpression, $formattingXPath)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text ('<xsl:apply-templates select="{0}/node()" mode="rich-text" />' -f $SelectExpression)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "</xsl:when>"
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text ('<xsl:when test="starts-with({0}, ''ATTACHMENT::'')">' -f $SelectExpression)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text '<a target="_blank">'
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text ('<xsl:attribute name="href"><xsl:value-of select="substring-after({0}, ''ATTACHMENT::'')" /></xsl:attribute>' -f $SelectExpression)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text ('<xsl:call-template name="file-name"><xsl:with-param name="path" select="translate(substring-after({0}, ''ATTACHMENT::''), ''\'', ''/'')" /></xsl:call-template>' -f $SelectExpression)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text '</a>'
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "</xsl:when>"
	$booleanValue = "translate(normalize-space($SelectExpression), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')"
	$booleanType = "$SelectExpression/@xsi:type = 'boolean' or substring-after($SelectExpression/@xsi:type, ':') = 'boolean'"
	$booleanTest = "not($SelectExpression/*) and not($SelectExpression/@xsi:nil = 'true' or $SelectExpression/@xsi:nil = '1') and ($booleanValue = 'true' or $booleanValue = 'false' or (($booleanType) and ($booleanValue = '1' or $booleanValue = '0')))"
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text ('<xsl:when test="{0}">' -f $booleanTest)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text '<input type="checkbox" disabled="disabled">'
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text ('<xsl:if test="{0} = ''true'' or {0} = ''1''"><xsl:attribute name="checked">checked</xsl:attribute></xsl:if>' -f $booleanValue)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text ('<xsl:attribute name="aria-label"><xsl:choose><xsl:when test="{0} = ''true'' or {0} = ''1''">True</xsl:when><xsl:otherwise>False</xsl:otherwise></xsl:choose></xsl:attribute>' -f $booleanValue)
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text '</input>'
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "</xsl:when>"
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "<xsl:otherwise>"
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text ('<span class="plain-text"><xsl:value-of select="{0}" /></span>' -f $SelectExpression)
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

	# The XML adapter can shadow .Name when a record contains a field named Name.
	$nodeName = $Node.get_Name()
	Write-Stage -Group Build -Message ("Rendering element '{0}' at XPath '{1}'" -f $nodeName, $CurrentXPath)

	$safeClass = Convert-ToSafeClassName -Name $nodeName
	$humanLabel = Convert-ToFieldLabel -ElementName $nodeName -Node $Node -SelectExpression $CurrentXPath
	Add-Line -Lines $Lines -IndentLevel $IndentLevel -Text "<div class=`"section section-$safeClass`">"
	Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "<div class=`"section-title`">$humanLabel</div>"

	$children = Get-ElementChildren -Node $Node
	if ($children.Count -eq 0) {
		Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "<div class=`"field-row`">"
		$humanLabel = Convert-ToFieldLabel -ElementName $nodeName -Node $Node -SelectExpression $CurrentXPath
		Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "<span class=`"field-label`">$humanLabel</span>"
		Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text '<div class="field-value">'
		Add-ValueRenderer -Lines $Lines -IndentLevel ($IndentLevel + 3) -SelectExpression $CurrentXPath
		Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "</div>"
		Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "</div>"
		Add-Line -Lines $Lines -IndentLevel $IndentLevel -Text "</div>"
		return
	}

	$repeatedNames = Get-RepeatedChildNames -Node $Node
	$renderedRepeated = @{}

	foreach ($child in $children) {
		$childName = $child.get_Name()
		$childXPathPart = Convert-ToXPathLiteral -Name $childName
		$childXPath = "$CurrentXPath/$childXPathPart"

		if ($repeatedNames.ContainsKey($childName)) {
			if ($renderedRepeated.ContainsKey($childName)) {
				continue
			}

			$renderedRepeated[$childName] = $true
			$sampleNode = Get-FirstChildByName -Node $Node -Name $childName

			if ($null -eq $sampleNode) {
				Write-Stage -Group Warn -Message ("Skipping repeated node '{0}' under '{1}' because no representative sample was found." -f $childName, $nodeName)
				Write-LogEntry -LogEntries $LogEntries -Message "Unable to find representative sample for repeated node '$childName' under '$nodeName'."
				continue
			}

			$sampleChildren = Get-ElementChildren -Node $sampleNode
			$humanLabel = Convert-ToFieldLabel -ElementName $childName -Node $sampleNode -SelectExpression $childXPath
			Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "<div class=`"collection collection-$(Convert-ToSafeClassName -Name $childName)`">"
			Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "<div class=`"collection-title`">$humanLabel</div>"

			if ($sampleChildren.Count -eq 0) {
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "<table class=`"data-table`">"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text "<thead><tr><th>Value</th></tr></thead>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text "<tbody>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 4) -Text ('<xsl:for-each select="{0}">' -f $childXPath)
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 5) -Text '<tr><td>'
				Add-ValueRenderer -Lines $Lines -IndentLevel ($IndentLevel + 6) -SelectExpression '.'
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 5) -Text '</td></tr>'
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 4) -Text "</xsl:for-each>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text "</tbody>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "</table>"
			}
			else {
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "<table class=`"data-table`">"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text "<thead>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 4) -Text "<tr>"

				foreach ($headerNode in $sampleChildren) {
					$headerPath = "$childXPath/$(Convert-ToXPathLiteral -Name $headerNode.get_Name())"
					$headerLabel = Convert-ToFieldLabel -ElementName $headerNode.get_Name() -Node $headerNode -SelectExpression $headerPath
					Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 5) -Text "<th>$headerLabel</th>"
				}

				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 4) -Text "</tr>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text "</thead>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text "<tbody>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 4) -Text ('<xsl:for-each select="{0}">' -f $childXPath)
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 5) -Text "<tr>"

				foreach ($valueNode in $sampleChildren) {
					$valuePath = Convert-ToXPathLiteral -Name $valueNode.get_Name()
					Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 6) -Text "<td>"
					if (Test-LeafNode -Node $valueNode) {
						Add-ValueRenderer -Lines $Lines -IndentLevel ($IndentLevel + 7) -SelectExpression $valuePath
					}
					else {
						Add-ElementPanel -Node $valueNode -CurrentXPath $valuePath -Lines $Lines -LogEntries $LogEntries -IndentLevel ($IndentLevel + 7)
					}
					Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 6) -Text "</td>"
				}

				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 5) -Text "</tr>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 4) -Text "</xsl:for-each>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 3) -Text "</tbody>"
				Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "</table>"

			}

			Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "</div>"
			continue
		}

		if (Test-LeafNode -Node $child) {
			Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 1) -Text "<div class=`"field-row`">"
			$childLabel = Convert-ToFieldLabel -ElementName $childName -Node $child -SelectExpression $childXPath
			Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "<span class=`"field-label`">$childLabel</span>"
			Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text '<div class="field-value">'
			Add-ValueRenderer -Lines $Lines -IndentLevel ($IndentLevel + 3) -SelectExpression $childXPath
			Add-Line -Lines $Lines -IndentLevel ($IndentLevel + 2) -Text "</div>"
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
$xmlDocument.PreserveWhitespace = $true
$xmlSettings = [System.Xml.XmlReaderSettings]::new()
$xmlSettings.IgnoreComments = $true
$xmlSettings.IgnoreWhitespace = $false
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
Convert-HtmlFieldValues -Document $xmlDocument -Configuration $fieldRules

$baseOutputName = [System.IO.Path]::GetFileName($OutputName)
if ([string]::IsNullOrWhiteSpace($baseOutputName)) {
	throw "OutputName must include at least one valid character."
}

$xsltPath = Join-Path -Path $OutputDirectory -ChildPath "$baseOutputName.xslt"
$cssPath = Join-Path -Path $OutputDirectory -ChildPath "$baseOutputName.css"
$logPath = Join-Path -Path $OutputDirectory -ChildPath "$baseOutputName.log"

$logEntries = New-Object System.Collections.Generic.List[string]
$xsnInventory = Import-XsnInventory -Path $XsnInventoryCsvPath
$xsnMatch = Find-XsnTemplate -Document $xmlDocument -Inventory $xsnInventory
$labelCount = Set-XsnDisplayLabels -Document $xmlDocument -Match $xsnMatch
Write-Stage -Group Input -Message "$($xsnMatch.Message) Applied labels: $labelCount."
Write-LogEntry -LogEntries $logEntries -Message "$($xsnMatch.Message) Applied labels: $labelCount."
Write-LogEntry -LogEntries $logEntries -Message "Started template generation."
Write-LogEntry -LogEntries $logEntries -Message "Input XML: $InfoPathXmlPath"
Write-LogEntry -LogEntries $logEntries -Message "Output XSLT: $xsltPath"
Write-LogEntry -LogEntries $logEntries -Message "Output CSS: $cssPath"
Write-Stage -Group Write -Message "XSLT output file: $xsltPath"
Write-Stage -Group Write -Message "CSS output file: $cssPath"
Write-Stage -Group Build -Message "Preparing XSLT template structure."

$xsltLines = New-Object System.Collections.Generic.List[string]
$cssRelativePath = [System.IO.Path]::GetFileName($cssPath)
$rootName = $xmlDocument.DocumentElement.get_Name()
$rootXPath = Convert-ToXPathLiteral -Name $rootName
Write-Stage -Group Build -Message "Detected root element: $rootName"

Add-Line -Lines $xsltLines -IndentLevel 0 -Text "<?xml version=`"1.0`" encoding=`"utf-8`"?>"
Add-Line -Lines $xsltLines -IndentLevel 0 -Text '<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:archive="urn:infopath-archive:display" exclude-result-prefixes="xsi archive">'
Add-Line -Lines $xsltLines -IndentLevel 1 -Text "<xsl:output method=`"html`" indent=`"yes`" encoding=`"utf-8`" />"
Add-Line -Lines $xsltLines -IndentLevel 1 -Text "<xsl:param name=`"sourceFilePath`" />"
Add-Line -Lines $xsltLines -IndentLevel 1 -Text "<xsl:param name=`"sourceFileName`" />"
Add-Line -Lines $xsltLines -IndentLevel 1 -Text ('<xsl:param name="UseDisplayName" select="{0}()" />' -f $UseDisplayName.ToString().ToLowerInvariant())
Add-Line -Lines $xsltLines -IndentLevel 1 -Text '<xsl:param name="UseEmbeddedXsnLabels" select="true()" />'
$richTextTemplates = @'
<xsl:template name="file-name">
  <xsl:param name="path" />
  <xsl:choose>
    <xsl:when test="contains($path, '/')"><xsl:call-template name="file-name"><xsl:with-param name="path" select="substring-after($path, '/')" /></xsl:call-template></xsl:when>
    <xsl:otherwise><xsl:value-of select="$path" /></xsl:otherwise>
  </xsl:choose>
</xsl:template>
<xsl:template match="text()" mode="rich-text"><xsl:value-of select="." /></xsl:template>
<xsl:template match="*" mode="rich-text">
  <xsl:variable name="tag" select="translate(local-name(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
  <xsl:choose>
    <xsl:when test="contains('|script|style|iframe|object|embed|', concat('|', $tag, '|'))" />
    <xsl:when test="namespace-uri() = 'http://www.w3.org/1999/xhtml' or namespace-uri() = ''">
      <xsl:choose>
        <xsl:when test="contains('|a|abbr|acronym|address|article|aside|b|bdi|bdo|big|blockquote|br|caption|center|cite|code|col|colgroup|dd|del|details|dfn|div|dl|dt|em|figcaption|figure|font|footer|h1|h2|h3|h4|h5|h6|header|hr|i|img|ins|kbd|li|main|mark|nav|ol|p|pre|q|s|samp|section|small|span|strike|strong|sub|summary|sup|table|tbody|td|tfoot|th|thead|time|tr|tt|u|ul|var|wbr|', concat('|', $tag, '|'))">
          <xsl:element name="{$tag}">
            <xsl:copy-of select="@*[namespace-uri() = '' and contains('|style|class|title|lang|dir|align|valign|width|height|colspan|rowspan|border|cellpadding|cellspacing|face|color|size|alt|start|type|', concat('|', local-name(), '|'))]" />
            <xsl:for-each select="@href[parent::*[local-name() = 'a']] | @src[parent::*[local-name() = 'img']]">
              <xsl:variable name="url" select="translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ &#x9;&#xA;&#xD;', 'abcdefghijklmnopqrstuvwxyz')" />
              <xsl:if test="not(contains($url, ':')) or starts-with($url, 'http:') or starts-with($url, 'https:') or (local-name() = 'href' and (starts-with($url, 'mailto:') or starts-with($url, 'tel:'))) or (local-name() = 'src' and (starts-with($url, 'data:image/png;') or starts-with($url, 'data:image/jpeg;') or starts-with($url, 'data:image/gif;') or starts-with($url, 'data:image/webp;')))">
                <xsl:copy />
              </xsl:if>
            </xsl:for-each>
            <xsl:apply-templates select="node()" mode="rich-text" />
          </xsl:element>
        </xsl:when>
        <xsl:otherwise><xsl:apply-templates select="node()" mode="rich-text" /></xsl:otherwise>
      </xsl:choose>
    </xsl:when>
    <xsl:otherwise><xsl:apply-templates select="node()" mode="rich-text" /></xsl:otherwise>
  </xsl:choose>
</xsl:template>
'@
Add-Line -Lines $xsltLines -IndentLevel 1 -Text $richTextTemplates
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

.plain-text {
  white-space: pre-wrap;
}

.checkbox-selection {
  display: inline-block;
  margin-right: 12px;
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
    XsnMatchStatus = $xsnMatch.Status
    XsnSolutionVersion = if ($null -ne $xsnMatch.Template) { $xsnMatch.Template.Row.SolutionVersion } else { '' }
    XsnTemplateUrl = if ($null -ne $xsnMatch.Template) { $xsnMatch.Template.Row.TemplateUrl } else { '' }
    XsnLabelsPath = if ($null -ne $xsnMatch.Template) { $xsnMatch.Template.LabelsPath } else { '' }
    XsltPath = $xsltPath
    CssPath  = $cssPath
    LogPath  = $logPath
    Success  = $true
}

# Write result as a special marker line for caller to parse
Write-Host "##TEMPLATE_GENERATION_RESULT##$(ConvertTo-Json -InputObject $result -Compress)##END_RESULT##" -ForegroundColor Magenta
