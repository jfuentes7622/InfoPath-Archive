<#
.SYNOPSIS
    Downloads files from a SharePoint library to a local folder using SharePoint Server module cmdlets.

.DESCRIPTION
    This script connects to a SharePoint site, reads items from a document library recursively,
    and saves files to a local destination while preserving folder structure. Includes colored
    console progress output and automatic log file generation.

.PARAMETER SiteUrl
    The SharePoint site URL (on-premises).
    Example: http://sharepoint/sites/Finance

.PARAMETER LibraryName
    The name of the document library to download from.
    Example: Shared Documents

.PARAMETER LocalPath
    The local folder path where files will be saved.
    Folder structure from SharePoint is preserved.
    Example: C:\Exports\SharePointFiles

.PARAMETER PageSize
    Number of items per page when querying the library (default: 500).

.PARAMETER Overwrite
    Controls how existing local files are handled:
    - NOT specified (default): Skip existing files, log as skipped
    - -Overwrite specified: Replace existing files with SharePoint version

.EXAMPLE
    # Download files without overwriting existing local files
    .\MoveLibraryFilesToLocal_SPS.ps1 `
      -SiteUrl "http://sharepoint/sites/Finance" `
      -LibraryName "Shared Documents" `
      -LocalPath "D:\SPExports\Finance"

.EXAMPLE
    # Download files and overwrite existing local files
    .\MoveLibraryFilesToLocal_SPS.ps1 `
      -SiteUrl "http://sharepoint/sites/Finance" `
      -LibraryName "Shared Documents" `
      -LocalPath "D:\SPExports\Finance" `
      -Overwrite

.NOTES
    Requires: Microsoft.SharePoint.PowerShell snap-in (run on SharePoint server with Management Shell)
    Output: Creates a timestamped log file in the local destination folder
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl = "http://sharepoint/sites/YourSite",

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$LibraryName = "Shared Documents",

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$LocalPath = "C:\Exports\SharePointFiles",

    [int]$PageSize = 500,

    [switch]$Overwrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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

        [ValidateSet("Init", "Connect", "Scan", "File", "Write", "Skip", "Done", "Warn", "Error")]
        [string]$Group = "Scan"
    )

    $color = switch ($Group) {
        "Init" { "Cyan" }
        "Connect" { "DarkCyan" }
        "Scan" { "Yellow" }
        "File" { "White" }
        "Write" { "Green" }
        "Skip" { "DarkGray" }
        "Done" { "Magenta" }
        "Warn" { "DarkYellow" }
        "Error" { "Red" }
        default { "White" }
    }

    Write-Host "[$Group] $Message" -ForegroundColor $color
}

function Ensure-SharePointSnapIn {
    if (-not (Get-PSSnapin -Registered -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "Microsoft.SharePoint.PowerShell" })) {
        throw "Microsoft.SharePoint.PowerShell snap-in is not available on this machine. Run on a SharePoint server with Management Shell installed."
    }

    if (-not (Get-PSSnapin -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "Microsoft.SharePoint.PowerShell" })) {
        Write-Stage -Group Init -Message "Loading Microsoft.SharePoint.PowerShell snap-in..."
        Add-PSSnapin Microsoft.SharePoint.PowerShell
    }
}

function Get-RelativePathFromLibrary {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LibraryServerRelativeUrl,

        [Parameter(Mandatory = $true)]
        [string]$FileServerRelativeUrl
    )

    $libraryRoot = $LibraryServerRelativeUrl.TrimEnd("/")
    $filePath = $FileServerRelativeUrl

    if ($filePath.StartsWith($libraryRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        $relative = $filePath.Substring($libraryRoot.Length).TrimStart("/")
        return $relative
    }

    return [System.IO.Path]::GetFileName($FileServerRelativeUrl)
}

function Save-SPFileToDisk {
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.SPFile]$SPFile,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [switch]$Overwrite
    )

    $destinationDirectory = Split-Path -Path $DestinationPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($destinationDirectory) -and -not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
        New-Item -Path $destinationDirectory -ItemType Directory -Force | Out-Null
    }

    if ((Test-Path -LiteralPath $DestinationPath -PathType Leaf) -and -not $Overwrite.IsPresent) {
        return $false
    }

    [System.IO.File]::WriteAllBytes($DestinationPath, $SPFile.OpenBinary())
    return $true
}

$logEntries = New-Object System.Collections.Generic.List[string]
$scriptStart = Get-Date

Write-Stage -Group Init -Message "Starting SharePoint library download."
Write-Stage -Group Init -Message "Site: $SiteUrl"
Write-Stage -Group Init -Message "Library: $LibraryName"
Write-Stage -Group Init -Message "Local path: $LocalPath"

Write-LogEntry -LogEntries $logEntries -Message "Started SharePoint library download."
Write-LogEntry -LogEntries $logEntries -Message "SiteUrl=$SiteUrl"
Write-LogEntry -LogEntries $logEntries -Message "LibraryName=$LibraryName"
Write-LogEntry -LogEntries $logEntries -Message "LocalPath=$LocalPath"
Write-LogEntry -LogEntries $logEntries -Message "PageSize=$PageSize"
Write-LogEntry -LogEntries $logEntries -Message "Overwrite=$($Overwrite.IsPresent)"

if (-not (Test-Path -LiteralPath $LocalPath -PathType Container)) {
    Write-Stage -Group Write -Message "Creating local destination root: $LocalPath"
    New-Item -Path $LocalPath -ItemType Directory -Force | Out-Null
}

$web = $null
$totalItemsScanned = 0
$totalFilesSaved = 0
$totalFilesSkipped = 0
$totalErrors = 0

try {
    Ensure-SharePointSnapIn

    Write-Stage -Group Connect -Message "Opening SPWeb: $SiteUrl"
    $web = Get-SPWeb -Identity $SiteUrl
    Write-Stage -Group Connect -Message "Connected to web: $($web.Title)"

    $list = $web.Lists.TryGetList($LibraryName)
    if ($null -eq $list) {
        throw "Library '$LibraryName' was not found in '$SiteUrl'."
    }

    $libraryRootUrl = $list.RootFolder.ServerRelativeUrl
    Write-Stage -Group Scan -Message "Library root URL: $libraryRootUrl"
    Write-Stage -Group Scan -Message "Scanning items with recursive paging (PageSize=$PageSize)."

    $query = New-Object Microsoft.SharePoint.SPQuery
    $query.RowLimit = [uint32]$PageSize
    $query.ViewAttributes = "Scope='RecursiveAll'"
    $query.ListItemCollectionPosition = $null

    $page = 0
    do {
        $page++
        Write-Stage -Group Scan -Message "Reading page $page..."
        $items = $list.GetItems($query)
        Write-Stage -Group Scan -Message "Page $page returned $($items.Count) item(s)."

        foreach ($item in $items) {
            $totalItemsScanned++
            $fsType = [int]$item["FSObjType"]

            if ($fsType -ne 0) {
                continue
            }

            $spFile = $item.File
            if ($null -eq $spFile) {
                Write-Stage -Group Warn -Message "Item $($item.ID) has no SPFile object; skipping."
                Write-LogEntry -LogEntries $logEntries -Message "WARN: Item $($item.ID) missing SPFile object."
                $totalErrors++
                continue
            }

            $relativePath = Get-RelativePathFromLibrary -LibraryServerRelativeUrl $libraryRootUrl -FileServerRelativeUrl $spFile.ServerRelativeUrl
            $relativePathLocal = ($relativePath -replace "/", "\")
            $destinationPath = Join-Path -Path $LocalPath -ChildPath $relativePathLocal

            Write-Stage -Group File -Message "Processing file: $($spFile.ServerRelativeUrl)"
            Write-LogEntry -LogEntries $logEntries -Message "FILE: $($spFile.ServerRelativeUrl) => $destinationPath"

            try {
                $saved = Save-SPFileToDisk -SPFile $spFile -DestinationPath $destinationPath -Overwrite:$Overwrite.IsPresent
                if ($saved) {
                    $totalFilesSaved++
                    Write-Stage -Group Write -Message "Saved: $destinationPath"
                }
                else {
                    $totalFilesSkipped++
                    Write-Stage -Group Skip -Message "Exists (use -Overwrite to replace): $destinationPath"
                }
            }
            catch {
                $totalErrors++
                Write-Stage -Group Error -Message "Failed to save file: $($spFile.ServerRelativeUrl)"
                Write-LogEntry -LogEntries $logEntries -Message "ERROR: Save failed for '$($spFile.ServerRelativeUrl)' :: $($_.Exception.Message)"
            }
        }

        $query.ListItemCollectionPosition = $items.ListItemCollectionPosition
    }
    while ($null -ne $query.ListItemCollectionPosition)
}
catch {
    $totalErrors++
    Write-Stage -Group Error -Message $_.Exception.Message
    Write-LogEntry -LogEntries $logEntries -Message "FATAL: $($_.Exception.Message)"
    throw
}
finally {
    if ($null -ne $web) {
        Write-Stage -Group Connect -Message "Disposing SPWeb object."
        $web.Dispose()
    }

    $duration = (Get-Date) - $scriptStart
    $logFileName = "MoveLibraryFiles_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss")
    $logPath = Join-Path -Path $LocalPath -ChildPath $logFileName

    Write-LogEntry -LogEntries $logEntries -Message "Completed. Duration=$($duration.ToString())"
    Write-LogEntry -LogEntries $logEntries -Message "Summary: ItemsScanned=$totalItemsScanned, FilesSaved=$totalFilesSaved, FilesSkipped=$totalFilesSkipped, Errors=$totalErrors"

    $logEntries | Set-Content -LiteralPath $logPath -Encoding UTF8

    Write-Stage -Group Done -Message "Completed library download."
    Write-Host "Items scanned : $totalItemsScanned" -ForegroundColor Magenta
    Write-Host "Files saved   : $totalFilesSaved" -ForegroundColor Green
    Write-Host "Files skipped : $totalFilesSkipped" -ForegroundColor DarkGray
    Write-Host "Errors        : $totalErrors" -ForegroundColor Red
    Write-Host "Log file      : $logPath" -ForegroundColor Magenta
}
