<#
.SYNOPSIS
    Tests download-only XSN exports with mocked SharePoint objects.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$exportScript = Join-Path -Path $repoRoot -ChildPath 'GetLibraryXsnVersions_SPS.ps1'
$testRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("InfoPathXsnTest-" + [guid]::NewGuid().ToString('N'))

function New-TestFile {
    param([byte[]]$Bytes, [string]$Label, [bool]$Current)

    $file = [pscustomobject]@{
        Bytes = $Bytes
        Exists = $true
        UIVersionLabel = $Label
        VersionLabel = $Label
        IsCurrentVersion = $Current
        TimeLastModified = [datetime]'2026-10-26T12:00:00'
        Created = [datetime]'2026-10-25T12:00:00'
        Versions = @()
        OpenCount = 0
    }
    $file | Add-Member -MemberType ScriptMethod -Name OpenBinary -Value {
        $this.OpenCount++
        return ,$this.Bytes
    }
    return $file
}

function New-TestSite {
    param([string]$Url, $File)

    $web = [pscustomobject]@{ ExpectedUrl = $Url; File = $File; DisposeCount = 0 }
    $web | Add-Member -MemberType ScriptMethod -Name GetFile -Value {
        param($Url)
        if ($Url -ne $this.ExpectedUrl) {
            throw "Unexpected template URL: $Url"
        }
        return $this.File
    }
    $web | Add-Member -MemberType ScriptMethod -Name Dispose -Value { $this.DisposeCount++ }
    $site = [pscustomobject]@{ Web = $web; DisposeCount = 0 }
    $site | Add-Member -MemberType ScriptMethod -Name OpenWeb -Value { return $this.Web }
    $site | Add-Member -MemberType ScriptMethod -Name Dispose -Value { $this.DisposeCount++ }
    return $site
}

function Get-PSSnapin {
    param([switch]$Registered)
    if ($testState.SnapInRegistered) {
        return [pscustomobject]@{ Name = 'Microsoft.SharePoint.PowerShell' }
    }
}

function Get-SPWeb {
    param([string]$Identity)
    if ($Identity -ne 'http://sharepoint/sites/Finance') {
        throw "Unexpected library web: $Identity"
    }
    return $testState.LibraryWeb
}

function New-Object {
    param([string]$TypeName, [object[]]$ArgumentList)
    if ($TypeName -ne 'Microsoft.SharePoint.SPSite') {
        throw "Unexpected New-Object type: $TypeName"
    }
    $url = [string]$ArgumentList[0]
    if (-not $testState.Sites.ContainsKey($url)) {
        throw "Unexpected SPSite URL: $url"
    }
    return $testState.Sites[$url]
}

function Invoke-TestExport {
    param([string]$Case, [string]$ExpectedError = '')

    $destination = Join-Path -Path $testRoot -ChildPath $Case
    $caughtMessage = ''
    try {
        & $exportScript -SiteUrl 'http://sharepoint/sites/Finance' -LibraryName 'Applications' -LocalPath $destination 6>$null
    }
    catch {
        $caughtMessage = $_.Exception.Message
    }
    if ($ExpectedError) {
        if ($caughtMessage -notlike $ExpectedError) {
            throw "Case '$Case' expected '$ExpectedError', got '$caughtMessage'."
        }
    }
    elseif ($caughtMessage) {
        $details = @(Get-ChildItem -LiteralPath $destination -Filter 'XsnTemplateDownloads.csv' -Recurse |
            ForEach-Object { Import-Csv -LiteralPath $_.FullName } |
            Where-Object Error |
            ForEach-Object { $_.Error }) -join '; '
        throw "Case '$Case' failed: $caughtMessage $details"
    }
    $directories = @(Get-ChildItem -LiteralPath $destination -Directory)
    if ($directories.Count -ne 1) {
        throw "Case '$Case' did not create exactly one export directory."
    }
    if (-not (Test-Path -LiteralPath (Join-Path -Path $directories[0].FullName -ChildPath 'GetLibraryXsnVersions.log'))) {
        throw "Case '$Case' did not preserve a log."
    }
    return $directories[0].FullName
}

try {
    New-Item -Path $testRoot -ItemType Directory | Out-Null
    $bytes1 = [System.Text.Encoding]::UTF8.GetBytes('unprocessed XSN bytes 1')
    $bytes2 = [System.Text.Encoding]::UTF8.GetBytes('unprocessed XSN bytes 2')
    $bytes3 = [System.Text.Encoding]::UTF8.GetBytes('unprocessed XSN bytes 3')
    $bytesB = [System.Text.Encoding]::UTF8.GetBytes('unprocessed XSN bytes B')
    $current = New-TestFile -Bytes $bytes3 -Label '3.0' -Current $true
    $old1 = New-TestFile -Bytes $bytes1 -Label '1.0' -Current $false
    $old2 = New-TestFile -Bytes $bytes2 -Label '2.0' -Current $false
    $current.Versions = @($old1, $old2, $current)
    $other = New-TestFile -Bytes $bytesB -Label '1.0' -Current $true
    $urlA = 'http://sharepoint/sites/Finance/Applications/Forms/template.xsn'
    $urlB = 'http://sharepoint/sites/Templates/Form%20B.XSN'
    $list = [pscustomobject]@{
        DocumentTemplateUrl = '/sites/Finance/Applications/Forms/template.xsn'
        ContentTypes = @(
            [pscustomobject]@{ Name = 'Form A'; Id = '0x0101A'; DocumentTemplateUrl = $urlA }
            [pscustomobject]@{ Name = 'Form B'; Id = '0x0101B'; DocumentTemplateUrl = $urlB }
            [pscustomobject]@{ Name = 'Document'; Id = '0x0101'; DocumentTemplateUrl = '/templates/document.dotx' }
            [pscustomobject]@{ Name = 'Empty'; Id = '0x01'; DocumentTemplateUrl = '' }
        )
    }
    $lists = [pscustomobject]@{ List = $list }
    $lists | Add-Member -MemberType ScriptMethod -Name TryGetList -Value {
        param($Name)
        if ($Name -ne 'Applications') { throw "Unexpected library: $Name" }
        return $this.List
    }
    $libraryWeb = [pscustomobject]@{ Url = 'http://sharepoint/sites/Finance'; Lists = $lists; DisposeCount = 0 }
    $libraryWeb | Add-Member -MemberType ScriptMethod -Name Dispose -Value { $this.DisposeCount++ }
    $testState = @{
        SnapInRegistered = $true
        LibraryWeb = $libraryWeb
        Sites = @{
            $urlA = (New-TestSite -Url $urlA -File $current)
            $urlB = (New-TestSite -Url $urlB -File $other)
        }
    }

    $result = Invoke-TestExport -Case 'Success With Spaces'
    $rows = @(Import-Csv -LiteralPath (Join-Path -Path $result -ChildPath 'XsnTemplateDownloads.csv'))
    if ($rows.Count -ne 4 -or @($rows | Where-Object Status -ne 'Success').Count -ne 0) {
        throw 'Expected four successfully inventoried XSN files.'
    }
    $formA = @($rows | Where-Object TemplateUrl -eq $urlA)
    if (($formA.SharePointVersion -join ',') -ne '3.0,1.0,2.0') {
        throw 'SharePoint history labels were not preserved.'
    }
    foreach ($column in @('SolutionVersion', 'SolutionName', 'Sha256', 'ManifestPath')) {
        if ($rows[0].PSObject.Properties.Name -contains $column) {
            throw "Server export unexpectedly contains processed column '$column'."
        }
    }
    if ($formA[0].IsCurrentVersion -ne 'True' -or $formA[1].IsCurrentVersion -ne 'False' -or $formA[0].ConfiguredBy -notlike '*Library; ContentType: Form A*') {
        throw 'Current/history flags or deduplicated configuration sources are incorrect.'
    }
    foreach ($row in $rows) {
        $relativeFile = Join-Path -Path $result -ChildPath $row.XsnRelativePath
        if (-not (Test-Path -LiteralPath $relativeFile) -or $relativeFile -ne $row.XsnPath) {
            throw 'An exported XSN is missing or has an incorrect portable path.'
        }
        $expectedFile = if ($row.TemplateUrl -eq $urlB) { $other } else { @($current, $old1, $old2) | Where-Object VersionLabel -eq $row.SharePointVersion }
        if ([System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($relativeFile)) -ne [System.Convert]::ToBase64String($expectedFile.Bytes)) {
            throw 'Download changed the raw XSN bytes.'
        }
    }
    foreach ($file in @($current, $old1, $old2, $other)) {
        if ($file.OpenCount -ne 1) {
            throw 'A version was skipped or downloaded more than once.'
        }
    }
    foreach ($site in $testState.Sites.Values) {
        if ($site.DisposeCount -ne 1 -or $site.Web.DisposeCount -ne 1) {
            throw 'Template site/web objects were not disposed.'
        }
    }
    if ($libraryWeb.DisposeCount -ne 1) { throw 'Library web was not disposed.' }

    $other.Exists = $false
    $result = Invoke-TestExport -Case 'MissingTemplate' -ExpectedError '*export incomplete: 1 failure*'
    $rows = @(Import-Csv -LiteralPath (Join-Path -Path $result -ChildPath 'XsnTemplateDownloads.csv'))
    if (@($rows | Where-Object Error -like '*does not exist*').Count -ne 1) {
        throw 'A missing configured template was not inventoried as a failure.'
    }
    $other.Exists = $true
    $lists.List = $null
    $null = Invoke-TestExport -Case 'MissingLibrary' -ExpectedError "*Library 'Applications' was not found*"
    $lists.List = $list
    $list.DocumentTemplateUrl = ''
    $savedContentTypes = $list.ContentTypes
    $list.ContentTypes = @()
    $null = Invoke-TestExport -Case 'NoTemplates' -ExpectedError '*No XSN templates are configured*'
    $list.DocumentTemplateUrl = 'Applications/Forms/template.xsn'
    $list.ContentTypes = $savedContentTypes
    $null = Invoke-TestExport -Case 'WebRelativeUrl'
    $testState.SnapInRegistered = $false
    $null = Invoke-TestExport -Case 'NoSnapIn' -ExpectedError '*is not installed*'

    Write-Host 'PASS: Download-only XSN history exports, portable paths, unchanged bytes, URL resolution, deduplication, failures, and disposal.'
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
