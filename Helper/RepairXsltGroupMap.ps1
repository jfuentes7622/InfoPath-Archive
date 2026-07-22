#
# Repairs InfoPathXsltGroupMap.csv by matching actual XSLT files in the templates directory
# and updating the CSV paths to match reality
#

param(
	[Parameter(Mandatory = $true)]
	[ValidateNotNullOrEmpty()]
	[string]$TemplateDirectory,

	[Parameter(Mandatory = $true)]
	[ValidateNotNullOrEmpty()]
	[string]$MapCsvPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Load the CSV
Write-Host "Loading CSV: $MapCsvPath" -ForegroundColor Cyan
$mapData = Import-Csv -LiteralPath $MapCsvPath -Encoding UTF8

# Get actual files in templates directory
Write-Host "Scanning template directory: $TemplateDirectory" -ForegroundColor Cyan
$xsltFiles = @(Get-ChildItem -LiteralPath $TemplateDirectory -Filter "*.xslt" -File)
$cssFiles = @(Get-ChildItem -LiteralPath $TemplateDirectory -Filter "*.css" -File)
$logFiles = @(Get-ChildItem -LiteralPath $TemplateDirectory -Filter "*.log" -File)

Write-Host "Found $($xsltFiles.Count) XSLT files, $($cssFiles.Count) CSS files, $($logFiles.Count) log files" -ForegroundColor Yellow

# Build lookup of files by base name (without extension)
$xsltLookup = @{}
$cssLookup = @{}
$logLookup = @{}

foreach ($file in $xsltFiles) {
	$baseName = $file.BaseName
	$xsltLookup[$baseName] = $file.FullName
}

foreach ($file in $cssFiles) {
	$baseName = $file.BaseName
	$cssLookup[$baseName] = $file.FullName
}

foreach ($file in $logFiles) {
	$baseName = $file.BaseName
	$logLookup[$baseName] = $file.FullName
}

# Repair each row
$repaired = 0
$skipped = 0

foreach ($row in $mapData) {
	$suggestedGroup = $row.SuggestedGroup
	$currentXsltPath = $row.XsltPath
	
	# Try to find the correct file by extracting the base name from the current path
	# and looking for matching files in the template directory
	
	# Pattern: tmpl__VERSION__SUGGESTEDGROUP.xslt
	# We need to find files that match the SuggestedGroup
	
	$candidateBaseName = $null
	
	# Try to find a file in the xslt directory that corresponds to this group
	# Look for files starting with "tmpl__" and containing part of the group
	foreach ($baseName in $xsltLookup.Keys) {
		# Check if this file might match the suggested group
		# The pattern should be: tmpl__VERSION__SOMETHING
		if ($baseName -like "tmpl__*__*") {
			# Extract what comes after the second double underscore
			$parts = $baseName -split '__'
			if ($parts.Count -ge 3) {
				# Try to match against SuggestedGroup or other hints
				# For now, we'll check if the group appears anywhere in the filename
				if ($baseName -match [regex]::Escape($suggestedGroup)) {
					$candidateBaseName = $baseName
					break
				}
			}
		}
	}
	
	if ($null -ne $candidateBaseName -and $xsltLookup.ContainsKey($candidateBaseName)) {
		$newXsltPath = $xsltLookup[$candidateBaseName]
		$newCssPath = if ($cssLookup.ContainsKey($candidateBaseName)) { $cssLookup[$candidateBaseName] } else { $null }
		$newLogPath = if ($logLookup.ContainsKey($candidateBaseName)) { $logLookup[$candidateBaseName] } else { $null }
		
		if ($newXsltPath -ne $currentXsltPath) {
			Write-Host "Repairing group '$suggestedGroup':" -ForegroundColor Green
			Write-Host "  Old XSLT: $currentXsltPath" -ForegroundColor DarkRed
			Write-Host "  New XSLT: $newXsltPath" -ForegroundColor Green
			
			$row.XsltPath = $newXsltPath
			$row.CssPath = if ([string]::IsNullOrWhiteSpace($newCssPath)) { $null } else { $newCssPath }
			$row.GeneratorLogPath = if ([string]::IsNullOrWhiteSpace($newLogPath)) { $null } else { $newLogPath }
			$repaired++
		}
	}
	else {
		Write-Host "Could not find template file for group '$suggestedGroup'" -ForegroundColor Yellow
		Write-Host "  Current path: $currentXsltPath" -ForegroundColor DarkYellow
		$skipped++
	}
}

Write-Host ""
Write-Host "Summary: $repaired rows repaired, $skipped rows skipped" -ForegroundColor Cyan

# Write corrected CSV
$backupPath = "$MapCsvPath.backup"
Copy-Item -LiteralPath $MapCsvPath -Destination $backupPath -Force
Write-Host "Backup created: $backupPath" -ForegroundColor Green

$mapData | Export-Csv -LiteralPath $MapCsvPath -NoTypeInformation -Encoding UTF8
Write-Host "Updated CSV written: $MapCsvPath" -ForegroundColor Green
