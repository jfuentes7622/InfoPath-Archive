<#
.SYNOPSIS
    Downloads configured InfoPath templates and their available SharePoint history.

.DESCRIPTION
    Run locally on a SharePoint Server using Windows PowerShell and the
    Microsoft.SharePoint.PowerShell snap-in. Discovers the library's document
    template and its content-type templates, deduplicates XSN URLs, and exports
    the current file plus all available historical file versions.

    Downloads bytes and SharePoint history metadata only; no XSN extraction,
    content hashing, or manifest processing runs on the server. Copy the entire export
    folder to the processing machine and pass XsnTemplateDownloads.csv to the
    local pipeline's XsnExportCsvPath parameter. Relative XSN paths remain valid
    after copying. Versions deleted from history cannot be recovered.

    Each run creates a new timestamped export folder. Does not change SharePoint
    files, extract templates into the HTML pipeline, or scan data XML for URLs.

.PARAMETER SiteUrl
    URL of the SharePoint Server web containing the library.

.PARAMETER LibraryName
    Display name of the form/document library.

.PARAMETER LocalPath
    Local parent folder for the timestamped export, CSV inventory, and log.

.EXAMPLE
    .\GetLibraryXsnVersions_SPS.ps1 `
      -SiteUrl "http://sharepoint/sites/Finance" `
      -LibraryName "Applications" `
      -LocalPath "D:\SPExports\Templates"
#>

#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$LibraryName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$LocalPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-ExportMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('Init', 'Scan', 'File', 'Done', 'Warn', 'Error')]
        [string]$Group = 'Scan'
    )

    $color = switch ($Group) {
        'Init' { 'Cyan' }
        'Scan' { 'Yellow' }
        'File' { 'White' }
        'Done' { 'Green' }
        'Warn' { 'DarkYellow' }
        'Error' { 'Red' }
    }
    Write-Host "[$Group] $Message" -ForegroundColor $color
    $logEntries.Add(("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Group, $Message))
}

function Add-ConfiguredTemplate {
    param(
        [AllowEmptyString()]
        [string]$TemplateUrl,

        [string]$Source,

        [uri]$WebUri,

        [hashtable]$Templates
    )

    if ([string]::IsNullOrWhiteSpace($TemplateUrl)) {
        return
    }
    $uri = [uri]::new($WebUri, $TemplateUrl)
    if ($uri.Scheme -notin @('http', 'https')) {
        throw "Unsupported template URL for ${Source}: $TemplateUrl"
    }
    if (-not $uri.AbsolutePath.EndsWith('.xsn', [System.StringComparison]::OrdinalIgnoreCase)) {
        Write-ExportMessage -Group Warn -Message "Not an XSN; skipping ${Source}: $uri"
        return
    }
    if ($uri.Query -or $uri.Fragment) {
        throw "Expected a direct XSN file URL without a query or fragment: $uri"
    }
    $key = $uri.AbsoluteUri
    if (-not $Templates.ContainsKey($key)) {
        $Templates[$key] = [pscustomobject]@{
            Url = $key
            Sources = [System.Collections.Generic.List[string]]::new()
        }
    }
    if (-not $Templates[$key].Sources.Contains($Source)) {
        $Templates[$key].Sources.Add($Source)
    }
}

function New-InventoryRow {
    param($Template)

    return [pscustomobject][ordered]@{
        SiteUrl = $SiteUrl
        LibraryName = $LibraryName
        TemplateUrl = $Template.Url
        ConfiguredBy = $Template.Sources -join '; '
        SharePointVersion = ''
        IsCurrentVersion = $false
        Modified = ''
        XsnRelativePath = ''
        XsnPath = ''
        Status = 'Failed'
        Error = ''
    }
}

$runName = 'XsnTemplates_{0}_{1}' -f (Get-Date -Format 'yyyyMMdd_HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8))
$exportRoot = Join-Path -Path $LocalPath -ChildPath $runName
New-Item -Path $exportRoot -ItemType Directory -Force | Out-Null
$exportRoot = (Resolve-Path -LiteralPath $exportRoot).Path
$csvPath = Join-Path -Path $exportRoot -ChildPath 'XsnTemplateDownloads.csv'
$logPath = Join-Path -Path $exportRoot -ChildPath 'GetLibraryXsnVersions.log'
$logEntries = [System.Collections.Generic.List[string]]::new()
$rows = [System.Collections.Generic.List[object]]::new()
$web = $null
$failedCount = 0
$completed = $false

try {
    Write-ExportMessage -Group Init -Message "Exporting XSN versions. Site: $SiteUrl; library: $LibraryName; destination: $exportRoot"
    if (-not (Get-Command Get-PSSnapin -ErrorAction SilentlyContinue)) {
        throw 'Run in Windows PowerShell on a SharePoint Server, not PowerShell 7.'
    }
    if (-not (Get-PSSnapin -Registered | Where-Object { $_.Name -eq 'Microsoft.SharePoint.PowerShell' })) {
        throw 'Microsoft.SharePoint.PowerShell is not installed. Run on a SharePoint Server with Management Shell.'
    }
    if (-not (Get-PSSnapin | Where-Object { $_.Name -eq 'Microsoft.SharePoint.PowerShell' })) {
        Add-PSSnapin Microsoft.SharePoint.PowerShell
    }
    $web = Get-SPWeb -Identity $SiteUrl
    $list = $web.Lists.TryGetList($LibraryName)
    if ($null -eq $list) {
        throw "Library '$LibraryName' was not found in '$SiteUrl'."
    }
    $webUri = [uri]($web.Url.TrimEnd('/') + '/')
    $templates = @{}
    Add-ConfiguredTemplate -TemplateUrl $list.DocumentTemplateUrl -Source 'Library' -WebUri $webUri -Templates $templates
    foreach ($contentType in $list.ContentTypes) {
        Add-ConfiguredTemplate -TemplateUrl $contentType.DocumentTemplateUrl -Source "ContentType: $($contentType.Name) [$($contentType.Id)]" -WebUri $webUri -Templates $templates
    }
    if ($templates.Count -eq 0) {
        throw 'No XSN templates are configured on the library or its content types.'
    }
    Write-ExportMessage -Message "Found $($templates.Count) distinct configured XSN template(s)."
    Write-ExportMessage -Group Warn -Message 'Only current files and retained SharePoint history can be exported. Deleted versions and former template URLs are not recoverable by this script.'

    foreach ($template in ($templates.Values | Sort-Object Url)) {
        $templateSite = $null
        $templateWeb = $null
        try {
            Write-ExportMessage -Message "Reading current file and version history: $($template.Url)"
            # Open the template's owning web, which can differ from the library's web.
            $templateSite = New-Object Microsoft.SharePoint.SPSite -ArgumentList $template.Url
            $templateWeb = $templateSite.OpenWeb()
            $file = $templateWeb.GetFile($template.Url)
            if ($null -eq $file -or -not $file.Exists) {
                throw "Configured XSN file does not exist: $($template.Url)"
            }
            $hasher = [System.Security.Cryptography.SHA256]::Create()
            try {
                $urlHash = [System.BitConverter]::ToString($hasher.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($template.Url))).Replace('-', '').Substring(0, 16)
            }
            finally {
                $hasher.Dispose()
            }
            $templateDirectory = Join-Path -Path $exportRoot -ChildPath $urlHash
            New-Item -Path $templateDirectory -ItemType Directory | Out-Null
            $versions = @(
                [pscustomobject]@{
                    File = $file
                    Label = $file.UIVersionLabel
                    IsCurrent = $true
                    Modified = $file.TimeLastModified
                }
                foreach ($version in $file.Versions) {
                    if (-not $version.IsCurrentVersion -and $version.VersionLabel -ne $file.UIVersionLabel) {
                        [pscustomobject]@{
                            File = $version
                            Label = $version.VersionLabel
                            IsCurrent = $false
                            Modified = $version.Created
                        }
                    }
                }
            )
            foreach ($version in $versions) {
                $row = New-InventoryRow -Template $template
                $row.SharePointVersion = $version.Label
                $row.IsCurrentVersion = $version.IsCurrent
                $row.Modified = $version.Modified.ToString('o')
                $rows.Add($row)
                try {
                    if ($version.Label -notmatch '^\d+\.\d+$') {
                        throw "Unexpected SharePoint file version label: '$($version.Label)'"
                    }
                    $xsnPath = Join-Path -Path $templateDirectory -ChildPath ("template_sp-{0}.xsn" -f $version.Label)
                    Write-ExportMessage -Group File -Message "Exporting SharePoint version $($version.Label) (current=$($version.IsCurrent)): $($template.Url)"
                    [System.IO.File]::WriteAllBytes($xsnPath, $version.File.OpenBinary())
                    $row.XsnPath = $xsnPath
                    $row.XsnRelativePath = Join-Path -Path $urlHash -ChildPath ("template_sp-{0}.xsn" -f $version.Label)
                    $row.Status = 'Success'
                    Write-ExportMessage -Group File -Message "Downloaded: $($row.XsnRelativePath)"
                }
                catch {
                    $failedCount++
                    $row.Error = $_.Exception.Message
                    Write-ExportMessage -Group Error -Message "Version $($version.Label) failed for $($template.Url): $($row.Error)"
                }
            }
        }
        catch {
            $failedCount++
            $row = New-InventoryRow -Template $template
            $row.Error = $_.Exception.Message
            $rows.Add($row)
            Write-ExportMessage -Group Error -Message "Template failed: $($template.Url): $($row.Error)"
        }
        finally {
            if ($null -ne $templateWeb) {
                $templateWeb.Dispose()
            }
            if ($null -ne $templateSite) {
                $templateSite.Dispose()
            }
        }
    }
    if ($failedCount -gt 0) {
        throw "XSN export incomplete: $failedCount failure(s). See the inventory and log in '$exportRoot'."
    }
    $completed = $true
}
catch {
    Write-ExportMessage -Group Error -Message $_.Exception.Message
    throw
}
finally {
    if ($null -ne $web) {
        $web.Dispose()
    }
    if ($rows.Count -gt 0) {
        $rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
        Write-ExportMessage -Message "Inventory: $csvPath"
    }
    if ($completed) {
        Write-ExportMessage -Group Done -Message "Exported $($rows.Count) XSN version(s) successfully."
    }
    else {
        Write-ExportMessage -Group Warn -Message 'Export did not complete successfully. Review the errors before using the inventory.'
    }
    Write-ExportMessage -Message "Log: $logPath"
    $logEntries | Set-Content -LiteralPath $logPath -Encoding UTF8
}
