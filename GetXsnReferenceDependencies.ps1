<#
.SYNOPSIS
    Discovers offline lookup dependencies without contacting SharePoint.
.DESCRIPTION
    Reads a processed XsnTemplateVersions.csv inventory and
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
.PARAMETER UserSiteUrl
    Explicit owning web for users. Otherwise derived only from the standard
    /web/Library/Forms/template.xsn location; never from a lookup connection.
#>
#requires -Version 5.1
[CmdletBinding(DefaultParameterSetName = 'Inventory')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Inventory')][string]$XsnInventoryCsvPath,
    [Parameter(Mandatory = $true, ParameterSetName = 'Extracted')][string]$ExtractedXsnDirectory,
    [Parameter(ParameterSetName = 'Extracted')][string]$TemplateUrl = '',
    [Parameter(Mandatory = $true)][string]$OutputPath,
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
New-Item -ItemType Directory -Path (Split-Path $outputFile -Parent) -Force | Out-Null
try {
    $inputs = @()
    if ($PSCmdlet.ParameterSetName -eq 'Inventory') {
        $inventory = Import-XsnInventory $XsnInventoryCsvPath
        foreach ($record in $inventory) {
            $inputs += [pscustomobject]@{ Directory = (Split-Path $record.LabelsPath -Parent); Url = $record.Row.TemplateUrl; Metadata = $record.Row }
        }
        foreach ($row in @(Import-Csv -LiteralPath $XsnInventoryCsvPath | Where-Object Status -ne 'Success')) {
            $warnings.Add("Inventory template '$($row.TemplateUrl)' was not successfully processed and was skipped.")
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
    $dependency = [ordered]@{
        SchemaVersion = 1; Templates = @($templates.ToArray()); Lists = @($lists.Values | Sort-Object SiteUrl, ListId)
        UserSites = @($templates.ToArray() | ForEach-Object { $_.UserSiteUrl } | Where-Object { $_ } | Sort-Object -Unique)
        Warnings = @($warnings.ToArray())
    }
    $dependency | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $outputFile -Encoding UTF8
    @('Local discovery only. Exports represent current data, not historical labels.') + $warnings.ToArray() |
        Set-Content -LiteralPath ($outputFile + '.log') -Encoding UTF8
}
if ($failed) { throw "Dependency discovery had failures. See '$outputFile' and '$outputFile.log'." }
Write-Output $outputFile
