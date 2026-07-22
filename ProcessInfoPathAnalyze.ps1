<#
.SYNOPSIS
	Analyzes InfoPath XML documents to discover generic structure without assumptions, grouped by version.

.DESCRIPTION
	Reads InfoPath XML files and generates structure reports identifying:
	  - Non-repeating fields (single occurrence)
	  - Repeating groups (elements appearing multiple times at same level)
	  - Element nesting hierarchy and depth
	  - Field types (leaf nodes, attributes, branches)
	  - Schema metadata to skip (namespaces, xsi attributes)

	Outputs are organized per version if a versioning CSV is provided:
	  - StructureReport_[VersionName].json: Detected structure with group metadata (per version)
	  - StructureRecommendations_[VersionName].csv: User-reviewable extraction recommendations (per version)
	  - AnalysisSummary.csv: Overview of all versions and sample files analyzed

	Without versioning CSV, generates single StructureReport.json for all files.
	For each version, selects sample file(s) with the highest field count.

.PARAMETER SourceDirectory
	Root folder containing representative InfoPath XML files to analyze.

.PARAMETER OutputDirectory
	Directory where structure reports and recommendations will be written.

.PARAMETER VersioningCsvPath
	Optional path to versioning CSV (output of ClassifyInfoPathVersions.ps1).
	Must have columns: SourceFilePath or FileName, and TemplateVersion or TemplateName.
	If provided, analysis is grouped by version with one report per version.
	If omitted, creates single StructureReport.json for all files.

.PARAMETER MaxFilesToAnalyze
	Maximum number of XML files to sample per version. Use 2-5 representative files for faster analysis.
	Default: 5

.PARAMETER Recurse
	Include subfolders when discovering XML files.

.EXAMPLE
	.\ProcessInfoPathAnalyze.ps1 `
	  -SourceDirectory "C:\Temp\RedApps" `
	  -OutputDirectory "C:\Temp\Analysis" `
	  -VersioningCsvPath "C:\Temp\Reports\InfoPathVersionMap.csv" `
	  -MaxFilesToAnalyze 3

.EXAMPLE
	.\ProcessInfoPathAnalyze.ps1 `
	  -SourceDirectory "C:\Temp\RedApps" `
	  -OutputDirectory "C:\Temp\Analysis"

.NOTES
	With versioning CSV: Generates StructureReport_[VersionName].json per version, selecting best sample per version.
	Without versioning CSV: Generates single StructureReport.json from all sampled files.
	Users should review StructureRecommendations_[Version].csv before using reports with ProcessInfoPathProperties.ps1.
#>

[CmdletBinding()]
param(
	# Folder containing representative InfoPath XML files.
	[Parameter(Mandatory = $true)]
	[ValidateNotNullOrEmpty()]
	[string]$SourceDirectory,

	# Output directory for structure report and recommendations.
	[Parameter(Mandatory = $true)]
	[ValidateNotNullOrEmpty()]
	[string]$OutputDirectory,

	# Optional versioning CSV from ClassifyInfoPathVersions.ps1 to group analysis by version.
	[Parameter()]
	[string]$VersioningCsvPath,

	# Maximum sample files to analyze per version (default 5).
	[int]$MaxFilesToAnalyze = 5,

	# Include subfolders when searching for XML files.
	[switch]$Recurse
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Stage {
	param(
		[Parameter(Mandatory = $true)]
		[string]$Message,

		[ValidateSet("Init", "Input", "Scan", "Analyze", "Report", "Done", "Warn", "Error")]
		[string]$Group = "Analyze"
	)

	$color = switch ($Group) {
		"Init" { "Cyan" }
		"Input" { "DarkCyan" }
		"Scan" { "Yellow" }
		"Analyze" { "Gray" }
		"Report" { "Green" }
		"Done" { "Magenta" }
		"Warn" { "DarkYellow" }
		"Error" { "Red" }
		default { "White" }
	}

	Write-Host "[$Group] $Message" -ForegroundColor $color
}

# Generic structure analyzer - no assumptions about element names or purposes
function Analyze-XmlStructure {
	param(
		[Parameter(Mandatory = $true)]
		[System.Xml.XmlNode]$Node,

		[Parameter(Mandatory = $true)]
		[string]$Path,

		[Parameter(Mandatory = $true)]
		[hashtable]$StructureMap,

		[int]$CurrentDepth = 0
	)

	# Track max depth
	if ($CurrentDepth -gt $StructureMap["MaxDepth"]) {
		$StructureMap["MaxDepth"] = $CurrentDepth
	}

	# Analyze attributes (namespace declarations and schema metadata)
	if ($null -ne $Node.Attributes) {
		foreach ($attribute in $Node.Attributes) {
			$attrName = $attribute.Name
			
			# Categorize attributes
			$isNamespace = $attrName -like "xmlns*"
			$isXsiAttribute = $attrName -like "xsi:*" -or $attrName -like "xml:*"
			
			$attrPath = "$Path/@$attrName"
			
			if (-not $StructureMap["AllAttributes"].ContainsKey($attrPath)) {
				$StructureMap["AllAttributes"][$attrPath] = @{
					Path        = $attrPath
					Name        = $attrName
					IsNamespace = $isNamespace
					IsXsiMeta   = $isXsiAttribute
					SampleValue = $attribute.Value
					Occurrences = 0
				}
			}
			$StructureMap["AllAttributes"][$attrPath]["Occurrences"]++
		}
	}

	# Get child elements
	$elementChildren = @($Node.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })
	
	if ($elementChildren.Count -eq 0) {
		# Leaf node - store as field
		$leafPath = $Path
		if (-not $StructureMap["AllFields"].ContainsKey($leafPath)) {
			$StructureMap["AllFields"][$leafPath] = @{
				Path             = $leafPath
				Type             = "leaf"
				SampleValue      = $Node.InnerText.Trim() | Select-Object -First 50  # First 50 chars
				Occurrences      = 0
				MaxOccurrences   = 1
				IsRepeating      = $false
				ParentPath       = $Path.Substring(0, $Path.LastIndexOf('/'))
			}
		}
		$StructureMap["AllFields"][$leafPath]["Occurrences"]++
		return
	}

	# Analyze sibling counts to detect repeating groups
	$elementCounts = @{}
	foreach ($child in $elementChildren) {
		$childName = $child.Name
		if (-not $elementCounts.ContainsKey($childName)) {
			$elementCounts[$childName] = 0
		}
		$elementCounts[$childName]++
	}

	# For each unique element name at this level, check if it repeats
	foreach ($childName in $elementCounts.Keys) {
		$childPath = "$Path/$childName"
		$isRepeating = $elementCounts[$childName] -gt 1
		
		if (-not $StructureMap["AllElements"].ContainsKey($childPath)) {
			$StructureMap["AllElements"][$childPath] = @{
				Path            = $childPath
				ElementName     = $childName
				Type            = "branch"
				Occurrences     = 0
				MaxOccurrences  = 0
				IsRepeating     = $isRepeating
				ChildElements   = @()
				FieldCount      = 0
				ParentPath      = $Path
				SampleChildren  = @()
			}
		}
		
		$element = $StructureMap["AllElements"][$childPath]
		if ($elementCounts[$childName] -gt $element["MaxOccurrences"]) {
			$element["MaxOccurrences"] = $elementCounts[$childName]
		}
		$element["Occurrences"]++
	}

	# Recursively analyze each child
	foreach ($child in $elementChildren) {
		$childPath = "$Path/$($child.Name)"
		Analyze-XmlStructure -Node $child -Path $childPath -StructureMap $StructureMap -CurrentDepth ($CurrentDepth + 1)
	}
}

# Converts analysis hashtable to structured report
function New-StructureReport {
	param(
		[Parameter(Mandatory = $true)]
		[hashtable]$StructureMap,

		[Parameter(Mandatory = $true)]
		[int]$FilesAnalyzed
	)

	$nonRepeatingFields = @()
	$repeatingGroups = @()

	# Identify non-repeating fields (leaf nodes that don't repeat)
	foreach ($fieldPath in $StructureMap["AllFields"].Keys) {
		$field = $StructureMap["AllFields"][$fieldPath]
		if (-not $field["IsRepeating"]) {
			$nonRepeatingFields += @{
				XPath    = $fieldPath
				Type     = "leaf"
				Sample   = $field["SampleValue"]
				Action   = "Include"
			}
		}
	}

	# Identify repeating groups (elements appearing multiple times)
	foreach ($elemPath in $StructureMap["AllElements"].Keys) {
		$elem = $StructureMap["AllElements"][$elemPath]
		if ($elem["IsRepeating"] -and $elem["MaxOccurrences"] -gt 1) {
			# Find all leaf fields under this repeating element
			$groupFields = @()
			foreach ($fieldPath in $StructureMap["AllFields"].Keys) {
				if ($fieldPath -like "$elemPath/*") {
					$field = $StructureMap["AllFields"][$fieldPath]
					$groupFields += @{
						XPath  = $fieldPath
						Sample = $field["SampleValue"]
					}
				}
			}
			
			# Auto-generate generic group name
			$lastSegment = $elemPath.Split('/')[-1]
			$groupName = "Group_$lastSegment"
			
			$repeatingGroups += @{
				GroupPath        = $elemPath
				GroupName        = $groupName
				MaxOccurrences   = $elem["MaxOccurrences"]
				FieldCount       = @($groupFields).Count
				ChildFields      = $groupFields
				Action           = "Create_Separate_Table"
			}
		}
	}

	# Filter out namespace and schema metadata
	$metadata = @{
		ElementsAnalyzed       = @($StructureMap["AllElements"].Keys).Count
		FieldsDiscovered       = @($StructureMap["AllFields"].Keys).Count
		AttributesDiscovered   = @($StructureMap["AllAttributes"].Keys).Count
		FilesAnalyzed          = $FilesAnalyzed
		MaxDepth               = $StructureMap["MaxDepth"]
		RepeatingGroupsFound   = @($repeatingGroups).Count
		NamespaceAttributeCount = @($StructureMap["AllAttributes"].Keys | Where-Object { $StructureMap["AllAttributes"][$_]["IsNamespace"] }).Count
	}

	return @{
		Metadata          = $metadata
		NonRepeatingFields = $nonRepeatingFields
		RepeatingGroups   = $repeatingGroups
	}
}

function Write-StructureOutputs {
	param(
		[Parameter(Mandatory = $true)]
		[hashtable]$StructureMap,

		[Parameter(Mandatory = $true)]
		[int]$FilesAnalyzed,

		[Parameter(Mandatory = $true)]
		[string]$ReportJsonPath,

		[Parameter(Mandatory = $true)]
		[string]$RecommendationsPath
	)

	$report = New-StructureReport -StructureMap $StructureMap -FilesAnalyzed $FilesAnalyzed
	$report | ConvertTo-Json -Depth 10 | Set-Content -Path $ReportJsonPath -Encoding UTF8

	$recommendations = New-Object System.Collections.Generic.List[object]

	foreach ($field in $report.NonRepeatingFields) {
		$recommendations.Add([pscustomobject]@{
			FieldPath      = $field.XPath
			Type           = $field.Type
			Category       = "NonRepeating"
			Recommendation = $field.Action
			SampleValue    = $field.Sample
		})
	}

	foreach ($group in $report.RepeatingGroups) {
		$recommendations.Add([pscustomobject]@{
			FieldPath      = $group.GroupPath
			Type           = "repeating_group"
			Category       = $group.GroupName
			Recommendation = $group.Action
			SampleValue    = "Max: $($group.MaxOccurrences), Fields: $($group.FieldCount)"
		})
	}

	$recommendations | Export-Csv -Path $RecommendationsPath -NoTypeInformation -Encoding UTF8
	return $report
}

if (-not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) {
	throw "SourceDirectory does not exist or is not a folder: $SourceDirectory"
}

Write-Stage -Group Init -Message "Starting InfoPath structure analysis."
Write-Stage -Group Input -Message "Source directory: $SourceDirectory"
Write-Stage -Group Input -Message "Max files to analyze: $MaxFilesToAnalyze"

$searchParameters = @{
	LiteralPath = $SourceDirectory
	Filter      = "*.xml"
	File        = $true
}

if ($Recurse.IsPresent) {
	$searchParameters.Recurse = $true
}

$allXmlFiles = @(Get-ChildItem @searchParameters)

if ($allXmlFiles.Count -eq 0) {
	throw "No XML files were found in: $SourceDirectory"
}

Write-Stage -Group Input -Message "Total XML files found: $($allXmlFiles.Count)"

# Helper function to count leaf nodes in an XML file
function Get-XmlFieldCount {
	param(
		[Parameter(Mandatory = $true)]
		[System.Xml.XmlNode]$Node
	)

	$count = 0
	$elementChildren = @($Node.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })
	
	if ($elementChildren.Count -eq 0) {
		return 1  # Leaf node
	}

	foreach ($child in $elementChildren) {
		$count += Get-XmlFieldCount -Node $child
	}

	return $count
}

# Build version grouping if versioning CSV is provided
$versionGroups = @{}

if (-not [string]::IsNullOrWhiteSpace($VersioningCsvPath) -and (Test-Path -LiteralPath $VersioningCsvPath -PathType Leaf)) {
	Write-Stage -Group Input -Message "Loading versioning CSV: $VersioningCsvPath"
	
	try {
		$versionData = Import-Csv -LiteralPath $VersioningCsvPath -Encoding UTF8
		
		# Build index of files by version
		foreach ($row in $versionData) {
			$version = $null
			$filePath = $null
			
			# Find version column
			foreach ($col in @("TemplateVersion", "TemplateName", "Version", "Group")) {
				if ($null -ne $row.$col -and -not [string]::IsNullOrWhiteSpace($row.$col)) {
					$version = $row.$col
					break
				}
			}
			
			# Find filepath column
			foreach ($col in @("SourceFilePath", "FilePath", "FileName", "Name")) {
				if ($null -ne $row.$col -and -not [string]::IsNullOrWhiteSpace($row.$col)) {
					$filePath = $row.$col
					break
				}
			}
			
			if ($null -ne $version -and -not [string]::IsNullOrWhiteSpace($filePath)) {
				if (-not $versionGroups.ContainsKey($version)) {
					$versionGroups[$version] = @()
				}
				$versionGroups[$version] += @{ FilePath = $filePath; Version = $version }
			}
		}
		
		Write-Stage -Group Input -Message "Found $($versionGroups.Count) versions in versioning CSV"
		foreach ($v in $versionGroups.Keys) {
			Write-Stage -Group Input -Message "  Version '$v': $($versionGroups[$v].Count) files"
		}
	}
	catch {
		Write-Stage -Group Warn -Message "Could not load versioning CSV: $($_.Exception.Message)"
		$versionGroups = @{}
	}
}

# Create output directory if needed
if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
	Write-Stage -Group Report -Message "Creating output directory: $OutputDirectory"
	New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
}

$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path

# Determine analysis strategy
if ($versionGroups.Count -gt 0) {
	# Version-based analysis
	Write-Stage -Group Analyze -Message "Performing version-based analysis ($($versionGroups.Count) versions)..."
	
	$analysisSummary = New-Object System.Collections.Generic.List[object]
	$summaryPath = Join-Path -Path $OutputDirectory -ChildPath "AnalysisSummary.csv"
	if (Test-Path -LiteralPath $summaryPath -PathType Leaf) {
		Remove-Item -LiteralPath $summaryPath -Force
	}
	
	foreach ($version in $versionGroups.Keys) {
		Write-Stage -Group Scan -Message "Processing version: $version"
		
		# Match files from XML directory to version group
		$versionFiles = @()
		foreach ($xmlFile in $allXmlFiles) {
			$fileName = $xmlFile.Name
			foreach ($versionEntry in $versionGroups[$version]) {
				$entryName = [System.IO.Path]::GetFileName($versionEntry.FilePath)
				if ($fileName -eq $entryName) {
					$versionFiles += $xmlFile
					break
				}
			}
		}
		
		if ($versionFiles.Count -eq 0) {
			Write-Stage -Group Warn -Message "No XML files found for version '$version' in directory"
			continue
		}
		
		Write-Stage -Group Scan -Message "  Found $($versionFiles.Count) files for version '$version', analyzing..."
		
		# Select files with highest field count
		$filesWithCounts = @()
		foreach ($xmlFile in $versionFiles) {
			try {
				$sr = [System.IO.StreamReader]::new($xmlFile.FullName, $true)
				$rawXml = $sr.ReadToEnd()
				$sr.Close()
				$sr.Dispose()

				$xmlSettings = [System.Xml.XmlReaderSettings]::new()
				$xmlSettings.IgnoreComments = $true
				$xmlSettings.IgnoreWhitespace = $true
				$xmlSettings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit

				$xmlDocument = [System.Xml.XmlDocument]::new()
				$xmlDocument.XmlResolver = $null
				$xr = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($rawXml), $xmlSettings)
				try {
					$xmlDocument.Load($xr)
				}
				finally {
					$xr.Dispose()
				}

				if ($null -ne $xmlDocument.DocumentElement) {
					$fieldCount = Get-XmlFieldCount -Node $xmlDocument.DocumentElement
					$filesWithCounts += @{ File = $xmlFile; FieldCount = $fieldCount }
				}
			}
			catch {
				Write-Stage -Group Warn -Message "  Could not count fields in $($xmlFile.Name): $($_.Exception.Message)"
			}
		}
		
		# Select top files by field count
		$selectedFiles = @($filesWithCounts | Sort-Object -Property FieldCount -Descending | Select-Object -First ([Math]::Min($MaxFilesToAnalyze, $filesWithCounts.Count)) | ForEach-Object { $_.File })
		
		Write-Stage -Group Scan -Message "  Selected $($selectedFiles.Count) file(s) with highest field counts for version '$version'"
		
		# Analyze selected files for this version
		$structureMap = @{
			AllElements    = @{}
			AllFields      = @{}
			AllAttributes  = @{}
			MaxDepth       = 0
		}

		# Pre-resolve output paths so they get updated while files are still processing
		$safeName = ($version -replace "[^A-Za-z0-9_-]", "_")
		$reportJsonPath = Join-Path -Path $OutputDirectory -ChildPath "StructureReport_$safeName.json"
		$recommendationsPath = Join-Path -Path $OutputDirectory -ChildPath "StructureRecommendations_$safeName.csv"
		$filesAnalyzedForVersion = 0
		
		foreach ($xmlFile in $selectedFiles) {
			Write-Stage -Group Scan -Message "    Analyzing: $($xmlFile.Name)"
			
			try {
				$sr = [System.IO.StreamReader]::new($xmlFile.FullName, $true)
				$rawXml = $sr.ReadToEnd()
				$sr.Close()
				$sr.Dispose()

				$xmlSettings = [System.Xml.XmlReaderSettings]::new()
				$xmlSettings.IgnoreComments = $true
				$xmlSettings.IgnoreWhitespace = $true
				$xmlSettings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit

				$xmlDocument = [System.Xml.XmlDocument]::new()
				$xmlDocument.XmlResolver = $null
				$xr = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($rawXml), $xmlSettings)
				try {
					$xmlDocument.Load($xr)
				}
				finally {
					$xr.Dispose()
				}

				if ($null -ne $xmlDocument.DocumentElement) {
					Analyze-XmlStructure -Node $xmlDocument.DocumentElement -Path $xmlDocument.DocumentElement.Name -StructureMap $structureMap
					$filesAnalyzedForVersion++

					# Stream report outputs as each sample file is analyzed
					$null = Write-StructureOutputs -StructureMap $structureMap -FilesAnalyzed $filesAnalyzedForVersion -ReportJsonPath $reportJsonPath -RecommendationsPath $recommendationsPath
					Write-Stage -Group Report -Message "    Updated outputs for version '$version' ($filesAnalyzedForVersion/$($selectedFiles.Count) files):"
					Write-Stage -Group Report -Message "      Report: $reportJsonPath"
					Write-Stage -Group Report -Message "      Recommendations: $recommendationsPath"
				}
			}
			catch {
				Write-Stage -Group Warn -Message "    Skipping file: $($_.Exception.Message)"
			}
		}

		if ($filesAnalyzedForVersion -eq 0) {
			Write-Stage -Group Warn -Message "  No valid sample files analyzed for version '$version'"
			continue
		}

		$report = New-StructureReport -StructureMap $structureMap -FilesAnalyzed $filesAnalyzedForVersion
		
		# Add to summary
		$summaryRow = [pscustomobject]@{
			Version           = $version
			FilesAnalyzed     = $filesAnalyzedForVersion
			ElementsDiscovered = $report.Metadata.ElementsAnalyzed
			FieldsDiscovered  = $report.Metadata.FieldsDiscovered
			RepeatingGroups   = $report.Metadata.RepeatingGroupsFound
			MaxDepth          = $report.Metadata.MaxDepth
			ReportPath        = $reportJsonPath
		}

		$analysisSummary.Add($summaryRow)
		if ($analysisSummary.Count -eq 1) {
			$summaryRow | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
		}
		else {
			$summaryRow | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8 -Append
		}
		Write-Stage -Group Report -Message "  Updated analysis summary: $summaryPath"
	}
	
	Write-Stage -Group Done -Message "Version-based analysis complete ($($analysisSummary.Count) versions analyzed)"
	Write-Stage -Group Done -Message "Output:"
	$analysisSummary | ForEach-Object { Write-Stage -Group Done -Message "  Version '$($_.Version)': $($_.FieldsDiscovered) fields, $($_.RepeatingGroups) repeating groups" }
	Write-Stage -Group Done -Message "  Summary: $summaryPath"
}
else {
	# Single unified analysis (no versioning)
	Write-Stage -Group Scan -Message "Performing unified analysis (no versioning)..."
	
	# Random sampling for unified analysis
	$filesToAnalyze = @($allXmlFiles | Get-Random -Count ([Math]::Min($MaxFilesToAnalyze, $allXmlFiles.Count)))
	Write-Stage -Group Scan -Message "Analyzing $($filesToAnalyze.Count) random sample files..."
	
	$structureMap = @{
		AllElements    = @{}
		AllFields      = @{}
		AllAttributes  = @{}
		MaxDepth       = 0
	}

	$reportJsonPath = Join-Path -Path $OutputDirectory -ChildPath "StructureReport.json"
	$recommendationsPath = Join-Path -Path $OutputDirectory -ChildPath "StructureRecommendations.csv"
	$filesAnalyzedUnified = 0
	
	foreach ($xmlFile in $filesToAnalyze) {
		Write-Stage -Group Scan -Message "Scanning: $($xmlFile.Name)"
		
		try {
			$sr = [System.IO.StreamReader]::new($xmlFile.FullName, $true)
			$rawXml = $sr.ReadToEnd()
			$sr.Close()
			$sr.Dispose()

			$xmlSettings = [System.Xml.XmlReaderSettings]::new()
			$xmlSettings.IgnoreComments = $true
			$xmlSettings.IgnoreWhitespace = $true
			$xmlSettings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit

			$xmlDocument = [System.Xml.XmlDocument]::new()
			$xmlDocument.XmlResolver = $null
			$xr = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($rawXml), $xmlSettings)
			try {
				$xmlDocument.Load($xr)
			}
			finally {
				$xr.Dispose()
			}

			if ($null -ne $xmlDocument.DocumentElement) {
				Analyze-XmlStructure -Node $xmlDocument.DocumentElement -Path $xmlDocument.DocumentElement.Name -StructureMap $structureMap
				$filesAnalyzedUnified++

				# Stream outputs as each file is analyzed
				$null = Write-StructureOutputs -StructureMap $structureMap -FilesAnalyzed $filesAnalyzedUnified -ReportJsonPath $reportJsonPath -RecommendationsPath $recommendationsPath
				Write-Stage -Group Report -Message "  Updated unified outputs ($filesAnalyzedUnified/$($filesToAnalyze.Count) files):"
				Write-Stage -Group Report -Message "    Report: $reportJsonPath"
				Write-Stage -Group Report -Message "    Recommendations: $recommendationsPath"
			}
		}
		catch {
			Write-Stage -Group Warn -Message "Skipping file: $($_.Exception.Message)"
		}
	}

	if ($filesAnalyzedUnified -eq 0) {
		throw "No valid XML sample files could be analyzed from: $SourceDirectory"
	}

	$report = New-StructureReport -StructureMap $structureMap -FilesAnalyzed $filesAnalyzedUnified
	
	# Console summary
	Write-Stage -Group Done -Message "Analysis complete."
	Write-Stage -Group Done -Message "  Elements discovered: $($report.Metadata.ElementsAnalyzed)"
	Write-Stage -Group Done -Message "  Fields discovered: $($report.Metadata.FieldsDiscovered)"
	Write-Stage -Group Done -Message "  Repeating groups: $($report.Metadata.RepeatingGroupsFound)"
	Write-Stage -Group Done -Message "  Max nesting depth: $($report.Metadata.MaxDepth)"
	Write-Stage -Group Done -Message "  Namespace metadata found: $($report.Metadata.NamespaceAttributeCount)"
	Write-Stage -Group Done -Message ""
	Write-Stage -Group Done -Message "Output:"
	Write-Stage -Group Done -Message "  Report: $reportJsonPath"
	Write-Stage -Group Done -Message "  Recommendations: $recommendationsPath"
}
