[CmdletBinding()]
param(
	[string]$SiteUrl = "https://contoso.sharepoint.com/sites/YourSite",

	[string]$LibraryName = "Shared Documents",

	[string]$OutputCsvPath = ".\SharePointLibraryMetadata.csv",

	[int]$PageSize = 500,

	[switch]$IncludeFolders,

	[string]$ClientId,

	[ValidateSet("Interactive", "DeviceLogin")]
	[string]$AuthMethod = "Interactive"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Convert-FieldValueToString {
	param(
		[Parameter(ValueFromPipeline = $true)]
		$Value
	)

	if ($null -eq $Value) {
		return $null
	}

	if ($Value -is [Array]) {
		return (($Value | ForEach-Object { Convert-FieldValueToString -Value $_ }) -join "; ")
	}

	if ($Value -is [DateTime]) {
		return $Value.ToString("o")
	}

	if ($Value.PSObject.Properties.Name -contains "Email") {
		if ([string]::IsNullOrWhiteSpace($Value.Email)) {
			return $Value.LookupValue
		}

		return $Value.Email
	}

	if ($Value.PSObject.Properties.Name -contains "LookupValue") {
		return $Value.LookupValue
	}

	if ($Value.PSObject.Properties.Name -contains "Label") {
		return $Value.Label
	}

	if ($Value.PSObject.Properties.Name -contains "Url" -and $Value.PSObject.Properties.Name -contains "Description") {
		return "$($Value.Url) ($($Value.Description))"
	}

	if ($Value -is [Guid]) {
		return $Value.Guid
	}

	return [string]$Value
}

function Get-ExpandedSharePointUser {
	param(
		[Parameter(Mandatory = $true)]
		$UserValue,

		[Parameter(Mandatory = $true)]
		[hashtable]$Cache
	)

	if ($null -eq $UserValue) {
		return $null
	}

	$lookupId = $UserValue.LookupId
	if ($null -eq $lookupId) {
		return [pscustomobject]@{
			Id          = $null
			DisplayName = $UserValue.LookupValue
			Email       = $UserValue.Email
			LoginName   = $null
			Title       = $null
			IsSiteAdmin = $null
		}
	}

	if (-not $Cache.ContainsKey($lookupId)) {
		$resolvedUser = Get-PnPUser -Identity $lookupId -ErrorAction SilentlyContinue

		$Cache[$lookupId] = [pscustomobject]@{
			Id          = $lookupId
			DisplayName = if ($null -ne $resolvedUser) { $resolvedUser.Title } else { $UserValue.LookupValue }
			Email       = if ($null -ne $resolvedUser) { $resolvedUser.Email } else { $UserValue.Email }
			LoginName   = if ($null -ne $resolvedUser) { $resolvedUser.LoginName } else { $null }
			Title       = if ($null -ne $resolvedUser) { $resolvedUser.Title } else { $UserValue.LookupValue }
			IsSiteAdmin = if ($null -ne $resolvedUser) { $resolvedUser.IsSiteAdmin } else { $null }
		}
	}

	return $Cache[$lookupId]
}

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
	throw "PnP.PowerShell is required. Install it with: Install-Module PnP.PowerShell -Scope CurrentUser"
}

Import-Module PnP.PowerShell

$connectParameters = @{
	Url = $SiteUrl
}

if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
	$connectParameters.ClientId = $ClientId
}

switch ($AuthMethod) {
	"Interactive" {
		$connectParameters.Interactive = $true
	}
	"DeviceLogin" {
		$connectParameters.DeviceLogin = $true
	}
}

Connect-PnPOnline @connectParameters

$list = Get-PnPList -Identity $LibraryName
$fieldDefinitions = Get-PnPField -List $LibraryName | Sort-Object InternalName
$fieldNames = $fieldDefinitions.InternalName
$userCache = @{}
$rows = New-Object System.Collections.Generic.List[object]

Get-PnPListItem -List $LibraryName -PageSize $PageSize -Fields $fieldNames | ForEach-Object {
	$item = $_
	$fileSystemObjectType = [int]$item["FSObjType"]

	if (-not $IncludeFolders.IsPresent -and $fileSystemObjectType -eq 1) {
		return
	}

	$author = $item["Author"]
	$editor = $item["Editor"]
	$expandedAuthor = Get-ExpandedSharePointUser -UserValue $author -Cache $userCache
	$expandedEditor = Get-ExpandedSharePointUser -UserValue $editor -Cache $userCache

	$row = [ordered]@{
		SiteUrl                 = $SiteUrl
		LibraryName             = $LibraryName
		ItemId                  = $item.Id
		UniqueId                = Convert-FieldValueToString -Value $item["UniqueId"]
		FileSystemObjectType    = if ($fileSystemObjectType -eq 1) { "Folder" } else { "File" }
		FileName                = Convert-FieldValueToString -Value $item["FileLeafRef"]
		FileRef                 = Convert-FieldValueToString -Value $item["FileRef"]
		FileDirRef              = Convert-FieldValueToString -Value $item["FileDirRef"]
		ContentTypeId           = Convert-FieldValueToString -Value $item["ContentTypeId"]
		Created                 = Convert-FieldValueToString -Value $item["Created"]
		Modified                = Convert-FieldValueToString -Value $item["Modified"]
		AuthorId                = if ($null -ne $expandedAuthor) { $expandedAuthor.Id } else { $null }
		AuthorDisplayName       = if ($null -ne $expandedAuthor) { $expandedAuthor.DisplayName } else { $null }
		AuthorEmail             = if ($null -ne $expandedAuthor) { $expandedAuthor.Email } else { $null }
		AuthorLoginName         = if ($null -ne $expandedAuthor) { $expandedAuthor.LoginName } else { $null }
		AuthorTitle             = if ($null -ne $expandedAuthor) { $expandedAuthor.Title } else { $null }
		AuthorIsSiteAdmin       = if ($null -ne $expandedAuthor) { $expandedAuthor.IsSiteAdmin } else { $null }
		EditorId                = if ($null -ne $expandedEditor) { $expandedEditor.Id } else { $null }
		EditorDisplayName       = if ($null -ne $expandedEditor) { $expandedEditor.DisplayName } else { $null }
		EditorEmail             = if ($null -ne $expandedEditor) { $expandedEditor.Email } else { $null }
		EditorLoginName         = if ($null -ne $expandedEditor) { $expandedEditor.LoginName } else { $null }
		EditorTitle             = if ($null -ne $expandedEditor) { $expandedEditor.Title } else { $null }
		EditorIsSiteAdmin       = if ($null -ne $expandedEditor) { $expandedEditor.IsSiteAdmin } else { $null }
	}

	foreach ($fieldName in $fieldNames) {
		if ($row.Contains($fieldName)) {
			continue
		}

		$row[$fieldName] = Convert-FieldValueToString -Value $item[$fieldName]
	}

	$rows.Add([pscustomobject]$row)
}

$outputDirectory = Split-Path -Path $OutputCsvPath -Parent
if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -Path $outputDirectory)) {
	New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

$rows |
	Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8

Disconnect-PnPOnline

Write-Host "Export complete: $OutputCsvPath"
Write-Host "Items exported: $($rows.Count)"
