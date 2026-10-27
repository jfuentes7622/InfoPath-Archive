<#
.SYNOPSIS
    Discovers offline lookup dependencies without contacting SharePoint.
.DESCRIPTION
    Accepts an XSN download export folder or XsnTemplateDownloads.csv and
    temporarily extracts its manifests and views using ProcessXsnTemplates.ps1.
    Run on SPS before the offline pipeline; no manual extraction or round trip
    to the processing machine is required. Windows expand.exe is required for
    this input mode; keep ProcessXsnTemplates.ps1 and the Helper folder alongside
    this script. Temporary extraction is removed on success and failure.
    Original downloads are not changed and no package code is executed.

    Also accepts a processed XsnTemplateVersions.csv inventory and
    all its declared views, or a directly extracted XSN. Writes dependency JSON
    and an adjacent .log, including rejected/ambiguous bindings as Warnings.
    Simple namespace-qualified MAIN fields and relative repeating contexts,
    generated GetDOM dropdowns (including uniqueness projections), and declared
    SharePoint list adapters are supported. Computed fields, arbitrary XSLT,
    named-template contexts and non-list adapters are explicitly not inferred.
    Mode-specific apply-templates chains are supported only when their MAIN
    context is unique and the root is declared by the manifest's schema metadata.
    No credentials are needed. Discovery does not execute template code.
    The eventual export is a CURRENT snapshot, not historical lookup labels.
.PARAMETER XsnExportCsvPath
    XsnTemplateDownloads.csv from GetLibraryXsnVersions_SPS.ps1, its containing
    folder, or a parent containing exactly one download inventory recursively.
    Automatically prepares the downloaded XSN versions in a temporary folder.
    Keep the original export's relative XSN paths intact.
.PARAMETER XsnInventoryCsvPath
    Existing processed XsnTemplateVersions.csv. Uses its previously extracted
    manifests and views; does not perform extraction in this input mode.
.PARAMETER ExtractedXsnDirectory
    An already-extracted XSN folder containing manifest.xsf and its view files.
.PARAMETER TemplateUrl
    Original absolute SharePoint template URL for an extracted XSN. Supplies
    context for relative list URLs and owning-web inference; never contacted.
.PARAMETER OutputPath
    Full dependency JSON filename, not a folder. Creates missing parent folders
    and writes an adjacent .log, preserving preparation diagnostics on failure.
.PARAMETER UserSiteUrl
    Explicit owning web for users. Otherwise derived only from the standard
    /web/Library/Forms/template.xsn location; never from a lookup connection.
.EXAMPLE
    .\GetXsnReferenceDependencies.ps1 `
      -XsnExportCsvPath "D:\SPExports\Templates\XsnTemplates_20261027_101341_179ba7ca" `
      -OutputPath "D:\SPExports\ReferenceDependencies.json"

    Run on SPS after downloading templates. Finds the download inventory and
    temporarily extracts all downloaded versions; no pipeline run is needed.
.EXAMPLE
    .\GetXsnReferenceDependencies.ps1 `
      -XsnExportCsvPath "D:\SPExports\Templates\XsnTemplateDownloads.csv" `
      -OutputPath "D:\SPExports\ReferenceDependencies.json"
    .\GetXsnReferenceData_SPS.ps1 `
      -DependencyManifestPath "D:\SPExports\ReferenceDependencies.json" `
      -LocalPath "D:\SPExports\ReferenceData"

    Discover dependencies and export lists/users on SPS before copying the
    original XSN export and complete reference-data export folder offline.
.EXAMPLE
    .\GetXsnReferenceDependencies.ps1 `
      -XsnInventoryCsvPath "D:\Data\PipelineOutput\XSNTemplates\XsnTemplateVersions.csv" `
      -OutputPath "D:\Data\ReferenceDependencies.json"
.EXAMPLE
    .\GetXsnReferenceDependencies.ps1 `
      -ExtractedXsnDirectory "D:\Data\InfopathData\XSNSample" `
      -TemplateUrl "http://sharepoint/sites/Finance/Applications/Forms/template.xsn" `
      -OutputPath "D:\Data\ReferenceDependencies.json"
#>
#requires -Version 5.1
[CmdletBinding(DefaultParameterSetName = 'Inventory')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Export')][ValidateNotNullOrEmpty()][string]$XsnExportCsvPath,
    [Parameter(Mandatory = $true, ParameterSetName = 'Inventory')][string]$XsnInventoryCsvPath,
    [Parameter(Mandatory = $true, ParameterSetName = 'Extracted')][string]$ExtractedXsnDirectory,
    [Parameter(ParameterSetName = 'Extracted')][string]$TemplateUrl = '',
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$OutputPath,
    [string]$UserSiteUrl = ''
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Helper\XsnTemplates.ps1')
. (Join-Path $PSScriptRoot 'Helper\XsnReferenceDiscovery.ps1')
$warnings = [System.Collections.Generic.List[string]]::new()
$templates = [System.Collections.Generic.List[object]]::new()
$lists = @{}
$failed = $false
$outputFile = [IO.Path]::GetFullPath($OutputPath)
if ($OutputPath -match '[\\/]$' -or (Test-Path -LiteralPath $outputFile -PathType Container)) {
    throw 'OutputPath must include a JSON filename, not just a folder (for example D:\SPExports\ReferenceDependencies.json).'
}
New-Item -ItemType Directory -Path (Split-Path $outputFile -Parent) -Force | Out-Null
$temporaryDirectory = $null
$preparationLog = @()
try {
    $inputs = @()
    if ($PSCmdlet.ParameterSetName -eq 'Export') {
        $processorPath = Join-Path $PSScriptRoot 'ProcessXsnTemplates.ps1'
        if (-not (Test-Path -LiteralPath $processorPath -PathType Leaf)) {
            throw "Required XSN processor not found: $processorPath. Copy ProcessXsnTemplates.ps1 and the Helper folder alongside this script."
        }
        $temporaryDirectory = Join-Path ([IO.Path]::GetTempPath()) ('InfoPathReferenceDependencies-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $temporaryDirectory | Out-Null
        $XsnInventoryCsvPath = Join-Path $temporaryDirectory 'XsnTemplateVersions.csv'
        try {
            $null = & $processorPath -InputCsvPath $XsnExportCsvPath -OutputDirectory $temporaryDirectory
        }
        catch {
            $warnings.Add("XSN preparation incomplete: $($_.Exception.Message)")
            $failed = $true
        }
        finally {
            $preparationLogPath = Join-Path $temporaryDirectory 'ProcessXsnTemplates.log'
            if (Test-Path -LiteralPath $preparationLogPath -PathType Leaf) {
                $preparationLog = @(Get-Content -LiteralPath $preparationLogPath)
            }
        }
    }
    if ($PSCmdlet.ParameterSetName -in @('Inventory', 'Export')) {
        $inventory = Import-XsnInventory $XsnInventoryCsvPath
        foreach ($record in $inventory) {
            $inputs += [pscustomobject]@{ Directory = (Split-Path $record.LabelsPath -Parent); Url = $record.Row.TemplateUrl; Metadata = $record.Row }
        }
        foreach ($row in @(Import-Csv -LiteralPath $XsnInventoryCsvPath | Where-Object Status -ne 'Success')) {
            $detail = if ($row.PSObject.Properties.Name -contains 'Error') { [string]$row.Error } else { [string]$row.Status }
            $warnings.Add("Inventory template '$($row.TemplateUrl)' was not successfully processed and was skipped: $detail")
            $failed = $true
        }
    }
    else { $inputs = @([pscustomobject]@{ Directory = $ExtractedXsnDirectory; Url = $TemplateUrl; Metadata = $null }) }
    foreach ($inputRecord in $inputs) {
        try {
            $result = Get-XsnReferenceTemplate $inputRecord.Directory $inputRecord.Url $UserSiteUrl $inputRecord.Metadata $warnings
            $templates.Add($result.Template)
            if ($result.HasFailures) { $failed = $true }
            foreach ($list in $result.Lists) {
                $key = $list.SiteUrl + '|' + $list.ListId
                if ($lists.ContainsKey($key)) { $lists[$key].Fields = @($lists[$key].Fields + $list.Fields | Sort-Object -Unique) }
                else { $lists[$key] = $list }
            }
        }
        catch { $warnings.Add("Template '$($inputRecord.Url)' failed: $($_.Exception.Message)"); $failed = $true }
    }
}
catch { $warnings.Add($_.Exception.Message); $failed = $true }
finally {
    try {
        $dependency = [ordered]@{
            SchemaVersion = 1; Templates = @($templates.ToArray()); Lists = @($lists.Values | Sort-Object SiteUrl, ListId)
            UserSites = @($templates.ToArray() | ForEach-Object { $_.UserSiteUrl } | Where-Object { $_ } | Sort-Object -Unique)
            Warnings = @($warnings.ToArray())
        }
        $dependency | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $outputFile -Encoding UTF8
        @('Discovery does not contact SharePoint. Exports represent current data, not historical labels.') + $warnings.ToArray() + $preparationLog |
            Set-Content -LiteralPath ($outputFile + '.log') -Encoding UTF8
    }
    finally {
        if ($null -ne $temporaryDirectory -and (Test-Path -LiteralPath $temporaryDirectory)) {
            Remove-Item -LiteralPath $temporaryDirectory -Recurse -Force
        }
    }
}
if ($failed) { throw "Dependency discovery had failures. See '$outputFile' and '$outputFile.log'." }
Write-Output $outputFile
