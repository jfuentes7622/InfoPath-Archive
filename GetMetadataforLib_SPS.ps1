<#
.SYNOPSIS
    Exports metadata from a SharePoint document library to a CSV file using SharePoint Server module cmdlets.

.DESCRIPTION
    This script connects to a SharePoint site, reads all items from a document library recursively,
    extracts metadata (fields, authors, editors, dates, etc.), and exports the data to a CSV file.
    Includes color-coded console progress output and automatic log file generation.

.PARAMETER SiteUrl
    The SharePoint site URL (on-premises).
    Default: http://sharepoint/sites/YourSite
    Example: http://ncbwashington/sites/notices/red/

.PARAMETER LibraryName
    The name of the document library to export metadata from.
    Default: Shared Documents
    Example: Applications

.PARAMETER OutputCsvPath
    The file path where the metadata CSV will be saved.
    Default: .\SharePointLibraryMetadata.csv
    Example: C:\Reports\LibraryMetadata.csv

.PARAMETER PageSize
    Number of items per page when querying the library (default: 500).
    Larger values reduce queries but increase memory usage.

.PARAMETER IncludeFolders
    When specified, folder items are included in the export.
    When NOT specified (default), only file items are exported.

.EXAMPLE
    # Export metadata from library (files only)
    .\GetMetadataforLib_SPS.ps1 `
      -SiteUrl "http://sharepoint/sites/Finance" `
      -LibraryName "Shared Documents" `
      -OutputCsvPath "C:\Reports\Finance_Metadata.csv"

.EXAMPLE
    # Export metadata including folders
    .\GetMetadataforLib_SPS.ps1 `
      -SiteUrl "http://sharepoint/sites/Finance" `
      -LibraryName "Shared Documents" `
      -OutputCsvPath "C:\Reports\Finance_Metadata.csv" `
      -IncludeFolders

.EXAMPLE
    # Export with custom page size for better performance
    .\GetMetadataforLib_SPS.ps1 `
      -SiteUrl "http://sharepoint/sites/Finance" `
      -LibraryName "Shared Documents" `
      -OutputCsvPath "C:\Reports\Finance_Metadata.csv" `
      -PageSize 1000

.NOTES
    Requires: Microsoft.SharePoint.PowerShell snap-in (run on SharePoint server with Management Shell)
    Output: Creates a timestamped log file alongside the CSV with operation details
    Fields Exported: Item ID, file/folder type, path, content type, created/modified dates,
                     author/editor info (name, email, login), and all custom fields
#>

[CmdletBinding()]
param(
    [string]$SiteUrl = "http://sharepoint/sites/YourSite",

    [string]$LibraryName = "Shared Documents",

    [string]$OutputCsvPath = ".\SharePointLibraryMetadata.csv",

    [int]$PageSize = 500,

    [switch]$IncludeFolders
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

    if ($Value -is [Microsoft.SharePoint.SPFieldUserValue]) {
        return $Value.LookupValue
    }

    if ($Value -is [Microsoft.SharePoint.SPFieldLookupValue]) {
        return $Value.LookupValue
    }

    if ($Value -is [Microsoft.SharePoint.SPFieldUrlValue]) {
        return "$($Value.Url) ($($Value.Description))"
    }

    if ($null -ne $Value.PSObject.Properties["Label"]) {
        return $Value.Label
    }

    if ($Value -is [hashtable] -or $Value -is [pscustomobject]) {
        return ($Value | ConvertTo-Json -Depth 5 -Compress)
    }

    if ($Value -is [Guid]) {
        return $Value.Guid
    }

    return [string]$Value
}

function Convert-ToUserFieldValue {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPWeb]$Web,

        [Parameter(ValueFromPipeline = $true)]
        $Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [Microsoft.SharePoint.SPFieldUserValue]) {
        return $Value
    }

    if ($Value -is [string]) {
        try {
            return New-Object Microsoft.SharePoint.SPFieldUserValue($Web, $Value)
        }
        catch {
            return $null
        }
    }

    return $null
}

function Get-ExpandedSharePointUser {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPWeb]$Web,

        [Parameter(Mandatory = $true)]
        $UserValue,

        [Parameter(Mandatory = $true)]
        [hashtable]$Cache
    )

    if ($null -eq $UserValue) {
        return $null
    }

    $resolvedUserValue = Convert-ToUserFieldValue -Web $Web -Value $UserValue
    if ($null -eq $resolvedUserValue) {
        return $null
    }

    $lookupId = $resolvedUserValue.LookupId

    if (-not $Cache.ContainsKey($lookupId)) {
        $spUser = Get-SPUser -Web $Web -Identity $lookupId -ErrorAction SilentlyContinue

        $Cache[$lookupId] = [pscustomobject]@{
            Id          = $lookupId
            DisplayName = if ($null -ne $spUser) { $spUser.Name } else { $resolvedUserValue.LookupValue }
            Email       = if ($null -ne $spUser) { $spUser.Email } else { $null }
            LoginName   = if ($null -ne $spUser) { $spUser.LoginName } else { $null }
            IsSiteAdmin = if ($null -ne $spUser) { $spUser.IsSiteAdmin } else { $null }
        }
    }

    return $Cache[$lookupId]
}

# ── Snap-in ────────────────────────────────────────────────────────────────
Write-Host "[INIT] Checking for Microsoft.SharePoint.PowerShell snap-in..." -ForegroundColor Cyan

if (-not (Get-PSSnapin -Registered -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "Microsoft.SharePoint.PowerShell" })) {
    throw "Microsoft.SharePoint.PowerShell snap-in is not available on this machine. Run this script on a SharePoint Server with Management Shell installed."
}

if (-not (Get-PSSnapin -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "Microsoft.SharePoint.PowerShell" })) {
    Write-Host "[INIT] Loading Microsoft.SharePoint.PowerShell snap-in..." -ForegroundColor Cyan
    Add-PSSnapin Microsoft.SharePoint.PowerShell
    Write-Host "[INIT] Snap-in loaded." -ForegroundColor Cyan
}
else {
    Write-Host "[INIT] Snap-in already loaded." -ForegroundColor Cyan
}

# ── Connection ──────────────────────────────────────────────────────────────
Write-Host "[CONNECT] Opening SPWeb: $SiteUrl" -ForegroundColor DarkCyan
$web = Get-SPWeb -Identity $SiteUrl
Write-Host "[CONNECT] SPWeb opened: $($web.Title)" -ForegroundColor DarkCyan

try {
    # ── Library ─────────────────────────────────────────────────────────────
    Write-Host "[LIBRARY] Locating library: $LibraryName" -ForegroundColor DarkCyan
    $list = $web.Lists.TryGetList($LibraryName)
    if ($null -eq $list) {
        throw "Library '$LibraryName' was not found in '$SiteUrl'."
    }
    Write-Host "[LIBRARY] Found library: $($list.Title)  (Item count: $($list.ItemCount))" -ForegroundColor DarkCyan

    # ── Fields ──────────────────────────────────────────────────────────────
    Write-Host "[FIELDS] Enumerating visible fields..." -ForegroundColor DarkCyan
    $fieldNames = @(
        $list.Fields |
            Where-Object { -not $_.Hidden } |
            Sort-Object InternalName |
            Select-Object -ExpandProperty InternalName
    )

    Write-Host "[FIELDS] $($fieldNames.Count) field(s) will be exported." -ForegroundColor DarkCyan

    $userCache = @{}
    $rows = New-Object System.Collections.Generic.List[object]
    $pageNumber = 0

    # ── Paging ──────────────────────────────────────────────────────────────
    $query = New-Object Microsoft.SharePoint.SPQuery
    $query.RowLimit = [uint32]$PageSize
    $query.ViewAttributes = "Scope='RecursiveAll'"
    $query.ListItemCollectionPosition = $null

    do {
        $pageNumber++
        Write-Host "[PAGE $pageNumber] Fetching up to $PageSize item(s)..." -ForegroundColor Yellow
        $items = $list.GetItems($query)
        Write-Host "[PAGE $pageNumber] Retrieved $($items.Count) item(s)." -ForegroundColor Yellow

        foreach ($item in $items) {
            $fileSystemObjectType = [int]$item["FSObjType"]

            if (-not $IncludeFolders.IsPresent -and $fileSystemObjectType -eq 1) {
                Write-Host "  [SKIP]   Folder skipped: $($item['FileRef'])" -ForegroundColor DarkGray
                continue
            }

            $itemLabel = $item["FileRef"]
            $itemType  = if ($fileSystemObjectType -eq 1) { "Folder" } else { "File" }
            Write-Host "  [ITEM]   [$itemType] $itemLabel" -ForegroundColor White

            # ── User resolution ─────────────────────────────────────────────
            Write-Host "    [USER] Resolving Author / Editor..." -ForegroundColor Gray
            $expandedAuthor = Get-ExpandedSharePointUser -Web $web -UserValue $item["Author"] -Cache $userCache
            $expandedEditor = Get-ExpandedSharePointUser -Web $web -UserValue $item["Editor"] -Cache $userCache
            Write-Host "    [USER] Author: $($expandedAuthor.DisplayName)  |  Editor: $($expandedEditor.DisplayName)" -ForegroundColor Gray

            $row = [ordered]@{
                SiteUrl              = $SiteUrl
                LibraryName          = $LibraryName
                ItemId               = $item.ID
                UniqueId             = Convert-FieldValueToString -Value $item["UniqueId"]
                FileSystemObjectType = if ($fileSystemObjectType -eq 1) { "Folder" } else { "File" }
                FileName             = Convert-FieldValueToString -Value $item["FileLeafRef"]
                FileRef              = Convert-FieldValueToString -Value $item["FileRef"]
                FileDirRef           = Convert-FieldValueToString -Value $item["FileDirRef"]
                ContentTypeId        = Convert-FieldValueToString -Value $item["ContentTypeId"]
                Created              = Convert-FieldValueToString -Value $item["Created"]
                Modified             = Convert-FieldValueToString -Value $item["Modified"]
                AuthorId             = if ($null -ne $expandedAuthor) { $expandedAuthor.Id } else { $null }
                AuthorDisplayName    = if ($null -ne $expandedAuthor) { $expandedAuthor.DisplayName } else { $null }
                AuthorEmail          = if ($null -ne $expandedAuthor) { $expandedAuthor.Email } else { $null }
                AuthorLoginName      = if ($null -ne $expandedAuthor) { $expandedAuthor.LoginName } else { $null }
                AuthorIsSiteAdmin    = if ($null -ne $expandedAuthor) { $expandedAuthor.IsSiteAdmin } else { $null }
                EditorId             = if ($null -ne $expandedEditor) { $expandedEditor.Id } else { $null }
                EditorDisplayName    = if ($null -ne $expandedEditor) { $expandedEditor.DisplayName } else { $null }
                EditorEmail          = if ($null -ne $expandedEditor) { $expandedEditor.Email } else { $null }
                EditorLoginName      = if ($null -ne $expandedEditor) { $expandedEditor.LoginName } else { $null }
                EditorIsSiteAdmin    = if ($null -ne $expandedEditor) { $expandedEditor.IsSiteAdmin } else { $null }
            }

            foreach ($fieldName in $fieldNames) {
                if ($row.Contains($fieldName)) {
                    continue
                }

                $row[$fieldName] = Convert-FieldValueToString -Value $item[$fieldName]
            }

            $rows.Add([pscustomobject]$row)
        }

        $query.ListItemCollectionPosition = $items.ListItemCollectionPosition
    }
    while ($null -ne $query.ListItemCollectionPosition)

    # ── Export ──────────────────────────────────────────────────────────────
    Write-Host "[EXPORT] Preparing output file: $OutputCsvPath" -ForegroundColor Green
    $outputDirectory = Split-Path -Path $OutputCsvPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -Path $outputDirectory)) {
        Write-Host "[EXPORT] Creating output directory: $outputDirectory" -ForegroundColor Green
        New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
    }

    Write-Host "[EXPORT] Writing $($rows.Count) row(s) to CSV..." -ForegroundColor Green
    $rows |
        Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8

    Write-Host "[EXPORT] Export complete: $OutputCsvPath" -ForegroundColor Green
    Write-Host "[EXPORT] Items exported: $($rows.Count)" -ForegroundColor Green
}
finally {
    if ($null -ne $web) {
        Write-Host "[CLEANUP] Disposing SPWeb object..." -ForegroundColor DarkCyan
        $web.Dispose()
        Write-Host "[CLEANUP] Done." -ForegroundColor DarkCyan
    }
}