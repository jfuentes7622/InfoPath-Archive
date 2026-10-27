<#
.SYNOPSIS
    Exports current lookup lists and site-collection user information read-only.
.DESCRIPTION
    Run on SharePoint Server with Windows PowerShell 5.1. Reads a dependency
    manifest produced by GetXsnReferenceDependencies.ps1 on this server (directly
    from the downloaded XSN export), or from previously extracted templates.
    Each run creates
    a new timestamped folder containing Manifest.json, CSV data and Export.log.
    No SharePoint changes, EnsureUser, directory/AD calls or credentials are used.
    User Information List reads include retained rows where SharePoint exposes
    them; deleted/purged rows cannot be recovered. User IDs are site-collection
    scoped; exact stored login/claims identities are preserved.
    Referenced URLs may identify subwebs. Opens each exact requested web, then
    reads its owning site-collection URL rather than treating it as a collection.

    This is a CURRENT snapshot, not historical labels as of an archived form.
    CSVs contain personal information: restrict access and retention, and copy
    them only into an appropriately protected offline archive. A failed export
    persists diagnostics and partial results, then throws. Failed CSVs must not
    be used as complete snapshots. No external modules are required.
.PARAMETER LocalPath
    Parent folder for a new timestamped export folder.
.EXAMPLE
    .\GetXsnReferenceDependencies.ps1 `
      -XsnExportCsvPath "D:\SPExports\Templates\XsnTemplates_20261027_101341_179ba7ca" `
      -OutputPath "D:\SPExports\ReferenceDependencies.json"
    .\GetXsnReferenceData_SPS.ps1 `
      -DependencyManifestPath "D:\SPExports\ReferenceDependencies.json" `
      -LocalPath "D:\SPExports\ReferenceData"

    Run both commands on SPS after GetLibraryXsnVersions_SPS.ps1. Discovery
    temporarily extracts manifests/views; no offline pipeline run is needed.
    Copy the original XSN export and the complete reference-data export offline.
#>
#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$DependencyManifestPath,
    [Parameter(Mandatory = $true)][string]$LocalPath,
    [ValidateRange(1, 10000)][int]$PageSize = 1000
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function ConvertTo-XsnExportUrl {
    param([string]$Value)
    $uri = $null
    if (-not [uri]::TryCreate($Value, [UriKind]::Absolute, [ref]$uri) -or
        $uri.Scheme -notin @('http', 'https') -or $uri.Query -or $uri.Fragment -or $uri.UserInfo) {
        throw "Invalid absolute site/web URL '$Value'."
    }
    return $uri.AbsoluteUri.TrimEnd('/')
}

function ConvertTo-XsnStoredValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }
    $typeName = $Value.GetType().FullName
    $valueType = $Value.GetType()
    while ($null -ne $valueType) {
        if ($valueType.FullName -in @('Microsoft.SharePoint.SPFieldLookupValue', 'Microsoft.SharePoint.SPFieldUserValue')) {
            return [string]$Value.LookupId
        }
        $valueType = $valueType.BaseType
    }
    if ($Value -is [datetime]) { return $Value.ToString('o', [Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [System.Collections.IEnumerable]) {
        return (@($Value | ForEach-Object { ConvertTo-XsnStoredValue $_ }) -join ';#')
    }
    if ($typeName -eq 'Microsoft.SharePoint.SPFieldUrlValue') { return "$($Value.Url), $($Value.Description)" }
    if ($Value -is [System.IFormattable]) { return $Value.ToString($null, [Globalization.CultureInfo]::InvariantCulture) }
    $text = $Value.ToString()
    if ($text -eq $typeName) { throw "No meaningful stored-value conversion for '$typeName'." }
    return $text
}

function Export-XsnListRows {
    param($List, [string[]]$Fields, [string]$Path, [switch]$Users)
    $query = New-Object Microsoft.SharePoint.SPQuery
    $query.RowLimit = [uint32]$PageSize
    $query.ViewAttributes = "Scope='RecursiveAll'"
    $query.ViewFieldsOnly = $true
    $query.IncludeMandatoryColumns = $false
    $query.Query = '<OrderBy><FieldRef Name="ID" Ascending="TRUE"/></OrderBy>'
    $query.ViewFields = (@($Fields | ForEach-Object { '<FieldRef Name="' + [System.Security.SecurityElement]::Escape($_) + '"/>' }) -join '')
    # Validate actual internal names rather than allowing SharePoint to resolve display names.
    $fieldMap = @{}
    foreach ($name in $Fields) {
        $field = $List.Fields.GetFieldByInternalName($name)
        if ($field.InternalName -cne $name) { throw "Field '$name' is not an exact internal field name." }
        $fieldMap[$name] = $field
    }
    $headers = if ($Users) { @('UserId', 'LoginName', 'DisplayName', 'Email') } else { $Fields }
    $header = (@($headers | ForEach-Object { '"' + $_.Replace('"', '""') + '"' }) -join ',')
    $writer = [IO.StreamWriter]::new($Path, $false, [Text.UTF8Encoding]::new($true))
    $count = 0
    try {
        $writer.WriteLine($header)
        do {
            $page = $List.GetItems($query)
            foreach ($item in $page) {
                $row = [ordered]@{}
                if ($Users) {
                    $row.UserId = [string]$item['ID']
                    $row.LoginName = [string]$item['Name']
                    $row.DisplayName = [string]$item['Title']
                    $row.Email = [string]$item['EMail']
                }
                else {
                    foreach ($name in $Fields) {
                        $value = $item[$name]
                        $field = $fieldMap[$name]
                        # Some versions expose stored lookup values as strings, others as typed values.
                        if ($value -is [string] -and $field.TypeAsString -in @('Lookup', 'LookupMulti', 'User', 'UserMulti') -and $value) {
                            $value = $field.GetFieldValue($value)
                        }
                        $row[$name] = ConvertTo-XsnStoredValue $value
                    }
                }
                $csv = @([pscustomobject]$row | ConvertTo-Csv -NoTypeInformation)
                $writer.WriteLine($csv[1])
                $count++
            }
            $query.ListItemCollectionPosition = $page.ListItemCollectionPosition
        } while ($null -ne $query.ListItemCollectionPosition)
    }
    finally { $writer.Dispose() }
    return $count
}

$runName = 'XsnReferences_{0}_{1}' -f (Get-Date -Format 'yyyyMMdd_HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8))
$root = Join-Path ([IO.Path]::GetFullPath($LocalPath)) $runName
New-Item -ItemType Directory -Path $root -Force | Out-Null
$log = [System.Collections.Generic.List[string]]::new()
$listResults = [System.Collections.Generic.List[object]]::new()
$userResults = [System.Collections.Generic.List[object]]::new()
$dependencies = $null
$failed = $false
try {
    $dependencies = Get-Content -LiteralPath $DependencyManifestPath -Raw | ConvertFrom-Json
    if ($dependencies.SchemaVersion -ne 1 -or $dependencies.Lists -isnot [array] -or $dependencies.UserSites -isnot [array] -or
        $dependencies.Templates -isnot [array] -or $dependencies.Warnings -isnot [array]) { throw 'Invalid dependency manifest: expected schema 1 with array collections.' }
    foreach ($warning in $dependencies.Warnings) { $log.Add("Discovery warning: $warning") }
    if (-not (Get-Command Get-SPWeb -ErrorAction SilentlyContinue)) {
        Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
    }
    $requests = @{}
    foreach ($request in $dependencies.Lists) {
        try {
            $url = ConvertTo-XsnExportUrl $request.SiteUrl
            $id = ([guid]$request.ListId).ToString('D')
            if ($request.Fields -isnot [array]) { throw 'List Fields must be an array of internal names.' }
            $fields = @($request.Fields + 'ID' | Sort-Object -Unique)
            if (@($fields | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count) { throw 'Empty internal field name.' }
            $key = $url + '|' + $id
            if ($requests.ContainsKey($key)) { $requests[$key].Fields = @($requests[$key].Fields + $fields | Sort-Object -Unique) }
            else { $requests[$key] = [pscustomobject]@{ SiteUrl = $url; ListId = $id; Fields = $fields } }
        }
        catch {
            $failed = $true
            $listResults.Add([pscustomobject]@{ SiteUrl = $request.SiteUrl; ListId = $request.ListId; Status = 'Failed'; Error = $_.Exception.Message; DataPath = ''; RowCount = 0; SiteCollectionUrl = '' })
            $log.Add("Invalid list request: $($_.Exception.Message)")
        }
    }
    foreach ($request in @($requests.Values | Sort-Object SiteUrl, ListId)) {
        $web = $null
        $result = [pscustomobject][ordered]@{ SiteUrl = $request.SiteUrl; ListId = $request.ListId; Status = 'Failed'; Error = ''; DataPath = ''; RowCount = 0; SiteCollectionUrl = '' }
        $listResults.Add($result)
        try {
            $web = Get-SPWeb -Identity $request.SiteUrl -ErrorAction Stop
            if ((ConvertTo-XsnExportUrl $web.Url) -ine $request.SiteUrl) { throw "Requested web does not resolve exactly: '$($request.SiteUrl)'." }
            $result.SiteCollectionUrl = ConvertTo-XsnExportUrl $web.Site.Url
            $list = $web.Lists[[guid]$request.ListId]
            $result.DataPath = 'List_{0:D4}.csv' -f $listResults.Count
            $result.RowCount = Export-XsnListRows $list $request.Fields (Join-Path $root $result.DataPath)
            $result.Status = 'Success'
        }
        catch { $result.Error = $_.Exception.Message; $failed = $true }
        finally {
            if ($null -ne $web) { $web.Dispose() }
            $log.Add("List $($result.SiteUrl) $($result.ListId): $($result.Status) $($result.Error)")
        }
    }
    $collections = @{}
    foreach ($requested in @($dependencies.UserSites | Sort-Object -Unique)) {
        $web = $null
        try {
            $url = ConvertTo-XsnExportUrl $requested
            $web = Get-SPWeb -Identity $url -ErrorAction Stop
            if ((ConvertTo-XsnExportUrl $web.Url) -ine $url) { throw "Requested user-source web does not resolve exactly: '$url'." }
            $collectionUrl = ConvertTo-XsnExportUrl $web.Site.Url
            if (-not $collections.ContainsKey($collectionUrl)) { $collections[$collectionUrl] = [System.Collections.Generic.List[string]]::new() }
            if (-not $collections[$collectionUrl].Contains($url)) { $collections[$collectionUrl].Add($url) }
        }
        catch {
            $failed = $true
            $userResults.Add([pscustomobject]@{ SiteCollectionUrl = ''; SourceSiteUrls = @($requested); Status = 'Failed'; Error = $_.Exception.Message; DataPath = ''; RowCount = 0 })
            $log.Add("User source '$requested': Failed $($_.Exception.Message)")
        }
        finally { if ($null -ne $web) { $web.Dispose() } }
    }
    foreach ($collectionUrl in @($collections.Keys | Sort-Object)) {
        $site = $null; $web = $null
        $result = [pscustomobject][ordered]@{ SiteCollectionUrl = $collectionUrl; SourceSiteUrls = @($collections[$collectionUrl].ToArray()); Status = 'Failed'; Error = ''; DataPath = ''; RowCount = 0 }
        $userResults.Add($result)
        try {
            $site = Get-SPSite -Identity $collectionUrl -ErrorAction Stop
            $web = $site.OpenWeb()
            $result.DataPath = 'Users_{0:D4}.csv' -f $userResults.Count
            # Not SiteUsers: the hidden User Information List also retains inactive entries.
            $result.RowCount = Export-XsnListRows $web.SiteUserInfoList @('ID', 'Name', 'Title', 'EMail') (Join-Path $root $result.DataPath) -Users
            $result.Status = 'Success'
        }
        catch { $result.Error = $_.Exception.Message; $failed = $true }
        finally {
            try { if ($null -ne $web) { $web.Dispose() } }
            finally { if ($null -ne $site) { $site.Dispose() } }
            $log.Add("Users $collectionUrl : $($result.Status) $($result.Error)")
        }
    }
}
catch { $failed = $true; $log.Add("Export failure: $($_.Exception.Message)") }
finally {
    [ordered]@{
        SchemaVersion = 1; ExportedAtUtc = [datetime]::UtcNow.ToString('o')
        Dependencies = $dependencies; Lists = @($listResults.ToArray()); Users = @($userResults.ToArray())
    } | ConvertTo-Json -Depth 40 | Set-Content -LiteralPath (Join-Path $root 'Manifest.json') -Encoding UTF8
    @('CURRENT snapshot, not historical labels. Contains personal data; protect access and retention. No credentials stored.') + $log.ToArray() |
        Set-Content -LiteralPath (Join-Path $root 'Export.log') -Encoding UTF8
}
if ($failed) { throw "Reference export had failures. See '$root\Manifest.json' and Export.log." }
Write-Output $root
