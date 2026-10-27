<#
.SYNOPSIS
    Extracts InfoPath solution metadata and view captions from local XSN versions.

.DESCRIPTION
    Processes XsnTemplateDownloads.csv and the XSN files copied with it.
    Resolves XsnRelativePath relative to the input CSV, not the server's original
    absolute XsnPath. Requires Windows expand.exe, but no SharePoint or InfoPath.

    Writes XsnTemplateVersions.csv, ProcessXsnTemplates.log, and extracted
    manifests, declared views, and namespace-aware field-label maps under
    OutputDirectory. Default-view captions take precedence; other views are
    used only for unambiguous labels. Failed downloads or extraction errors are
    recorded in the inventory; other rows continue, then the script throws if
    any rows failed. Source files and the download inventory are not modified.

.PARAMETER InputCsvPath
    Path to the copied SPS export's XsnTemplateDownloads.csv.

.PARAMETER OutputDirectory
    Folder for processed XSN information. The pipeline uses its XSNTemplates folder.

.PARAMETER ExtractResources
    Also extract declared rendering resources with original relative paths.
    Does not extract or execute managed code.

.EXAMPLE
    .\ProcessXsnTemplates.ps1 `
      -InputCsvPath "D:\Data\XsnExport\XsnTemplateDownloads.csv" `
      -OutputDirectory "D:\Data\PipelineOutput\XSNTemplates"
#>

#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$InputCsvPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputDirectory,

    [switch]$ExtractResources
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path -Path $PSScriptRoot -ChildPath 'Helper\XsnTemplates.ps1')

function Write-XsnMessage {
    param(
        [string]$Message,
        [ValidateSet('Init', 'File', 'Done', 'Warn', 'Error')]
        [string]$Group = 'File'
    )

    $color = switch ($Group) {
        'Init' { 'Cyan' }
        'File' { 'White' }
        'Done' { 'Green' }
        'Warn' { 'DarkYellow' }
        'Error' { 'Red' }
    }
    Write-Host "[$Group] $Message" -ForegroundColor $color
    $logEntries.Add(("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Group, $Message))
}

function Expand-XsnResource {
    param([string]$XsnPath, [string]$ExpandPath, [string]$Resource, [string]$Directory)

    $relative = $Resource.Replace('/', '\')
    if ([string]::IsNullOrWhiteSpace($relative) -or [System.IO.Path]::IsPathRooted($relative) -or '..' -in $relative.Split('\')) {
        throw "Invalid XSN resource path: '$Resource'"
    }
    New-Item -Path $Directory -ItemType Directory -Force | Out-Null
    $expandOutput = & $ExpandPath '-I' $XsnPath ("-F:$Resource") $Directory 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Cannot extract '$Resource' (expand exit $LASTEXITCODE): $($expandOutput -join ' ')"
    }
    $path = Join-Path -Path $Directory -ChildPath ([System.IO.Path]::GetFileName($relative))
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "The XSN does not contain '$Resource': $XsnPath"
    }
    return $path
}

function Read-XsnManifest {
    param(
        [string]$XsnPath,
        [string]$ExpandPath,
        [string]$ManifestDirectory
    )

    $tempDirectory = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("InfoPathManifest-" + [guid]::NewGuid().ToString('N'))
    New-Item -Path $tempDirectory -ItemType Directory | Out-Null
    try {
        $manifestPath = Expand-XsnResource -XsnPath $XsnPath -ExpandPath $ExpandPath -Resource 'manifest.xsf' -Directory $tempDirectory
        $document = Read-XsnXml -Path $manifestPath
        $root = $document.DocumentElement
        if ($null -eq $root -or $root.LocalName -ne 'xDocumentClass') {
            throw "manifest.xsf does not have an xDocumentClass root: $XsnPath"
        }
        $solutionVersion = $root.GetAttribute('solutionVersion')
        if ([string]::IsNullOrWhiteSpace($solutionVersion)) {
            throw "manifest.xsf has no solutionVersion: $XsnPath"
        }
        $viewsNode = $root.SelectSingleNode("*[local-name()='views']")
        $views = if ($null -ne $viewsNode) { @($viewsNode.SelectNodes("*[local-name()='view']")) } else { @() }
        $defaultView = if ($null -ne $viewsNode) { $viewsNode.GetAttribute('default') } else { '' }
        $candidates = [System.Collections.Generic.List[object]]::new()
        $viewFiles = [System.Collections.Generic.List[object]]::new()
        foreach ($view in $views) {
            $pane = $view.SelectSingleNode("*[local-name()='mainpane']")
            if ($null -eq $pane -or -not $pane.GetAttribute('transform')) {
                throw "XSN view '$($view.GetAttribute('name'))' has no mainpane transform."
            }
            $resource = $pane.GetAttribute('transform')
            $viewPath = Expand-XsnResource -XsnPath $XsnPath -ExpandPath $ExpandPath -Resource $resource -Directory (Join-Path -Path $tempDirectory -ChildPath ("View{0}" -f $viewFiles.Count))
            $viewDocument = Read-XsnXml -Path $viewPath
            foreach ($candidate in @(Get-XsnViewLabelCandidates -Document $viewDocument -ViewName $view.GetAttribute('name') -IsDefault ($view.GetAttribute('name') -ceq $defaultView))) {
                $candidates.Add($candidate)
            }
            $viewFiles.Add([pscustomobject]@{ Path = $viewPath; Resource = $resource })
        }
        $resolved = @(Resolve-XsnViewLabels -Candidates $candidates.ToArray())
        $labels = @($resolved | Where-Object Status -eq 'Resolved' | Select-Object Path, Label)
        New-Item -Path $ManifestDirectory -ItemType Directory -Force | Out-Null
        $outputManifest = Join-Path -Path $ManifestDirectory -ChildPath 'manifest.xsf'
        Copy-Item -LiteralPath $manifestPath -Destination $outputManifest -Force
        for ($index = 0; $index -lt $viewFiles.Count; $index++) {
            $outputView = Join-Path -Path $ManifestDirectory -ChildPath ('View{0:D4}.xsl' -f ($index + 1))
            Copy-Item -LiteralPath $viewFiles[$index].Path -Destination $outputView -Force
        }
        $labelsPath = Join-Path -Path $ManifestDirectory -ChildPath 'FieldLabels.json'
        if ($ExtractResources) {
            $resourceDirectory = Join-Path -Path $ManifestDirectory -ChildPath 'Resources'
            New-Item -Path $resourceDirectory -ItemType Directory -Force | Out-Null
            Copy-Item -LiteralPath $manifestPath -Destination (Join-Path $resourceDirectory 'manifest.xsf') -Force
            $resources = @($viewFiles | ForEach-Object { $_.Resource })
            $resources += @($root.SelectNodes("*[local-name()='package']/*[local-name()='files']/*[local-name()='file']") |
                ForEach-Object { $_.GetAttribute('name') } |
                Where-Object { [System.IO.Path]::GetExtension($_).ToLowerInvariant() -in @('.xsl', '.xslt', '.css', '.gif', '.png', '.jpg', '.jpeg', '.bmp', '.ico', '.woff', '.woff2', '.ttf') })
            foreach ($resource in @($resources | Sort-Object -Unique)) {
                $relative = $resource.Replace('/', '\')
                if ([System.IO.Path]::IsPathRooted($relative) -or '..' -in $relative.Split('\')) {
                    throw "Invalid XSN resource path: '$resource'"
                }
                $resourceTarget = Join-Path $resourceDirectory $relative
                $extracted = Expand-XsnResource -XsnPath $XsnPath -ExpandPath $ExpandPath -Resource $resource -Directory (Join-Path $tempDirectory 'Resource')
                New-Item -Path (Split-Path $resourceTarget -Parent) -ItemType Directory -Force | Out-Null
                Copy-Item -LiteralPath $extracted -Destination $resourceTarget -Force
            }
        }
        [pscustomobject]@{ Labels = $labels } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $labelsPath -Encoding UTF8
        if ($candidates.Count -gt 0) {
            $candidates | Export-Csv -LiteralPath (Join-Path -Path $ManifestDirectory -ChildPath 'FieldLabelCandidates.csv') -NoTypeInformation -Encoding UTF8
        }
        if ($resolved.Count -gt 0) {
            $resolved | Export-Csv -LiteralPath (Join-Path -Path $ManifestDirectory -ChildPath 'FieldLabelResolution.csv') -NoTypeInformation -Encoding UTF8
        }
        $ambiguous = @($resolved | Where-Object Status -eq 'Ambiguous').Count
        $unsupported = @($candidates | Where-Object Status -eq 'UnsupportedBinding').Count
        if ($ambiguous -gt 0 -or $unsupported -gt 0 -or $labels.Count -eq 0) {
            Write-XsnMessage -Group Warn -Message "XSN labels: resolved=$($labels.Count), ambiguous=$ambiguous, unsupported bindings=$unsupported. Unresolved fields use XML-name labels."
        }
        return [pscustomobject]@{
            SolutionVersion = $solutionVersion
            SolutionName = $root.GetAttribute('name')
            ManifestPath = $outputManifest
            LabelsPath = $labelsPath
            LabelCount = $labels.Count
            ViewCount = $viewFiles.Count
        }
    }
    finally {
        Remove-Item -LiteralPath $tempDirectory -Recurse -Force
    }
}

New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
$outputRoot = (Resolve-Path -LiteralPath $OutputDirectory).Path
$csvPath = Join-Path -Path $outputRoot -ChildPath 'XsnTemplateVersions.csv'
$logPath = Join-Path -Path $outputRoot -ChildPath 'ProcessXsnTemplates.log'
$logEntries = [System.Collections.Generic.List[string]]::new()
$rows = [System.Collections.Generic.List[object]]::new()
$failedCount = 0

try {
    Write-XsnMessage -Group Init -Message "Processing local XSN export: $InputCsvPath"
    if (-not (Test-Path -LiteralPath $InputCsvPath -PathType Leaf)) {
        throw "XSN download inventory not found: $InputCsvPath"
    }
    $inputPath = (Resolve-Path -LiteralPath $InputCsvPath).Path
    if ($inputPath -eq $csvPath) {
        throw 'The download inventory and processed inventory must be different files.'
    }
    $inputRoot = Split-Path -Path $inputPath -Parent
    $downloads = @(Import-Csv -LiteralPath $inputPath)
    if ($downloads.Count -eq 0) {
        throw "XSN download inventory is empty: $inputPath"
    }
    foreach ($column in @('SiteUrl', 'LibraryName', 'TemplateUrl', 'ConfiguredBy', 'SharePointVersion', 'IsCurrentVersion', 'Modified', 'XsnRelativePath', 'Status', 'Error')) {
        if ($downloads[0].PSObject.Properties.Name -notcontains $column) {
            throw "XSN download inventory is missing '$column'. Use the updated SPS exporter and copy its entire export folder."
        }
    }
    $expandPath = Join-Path -Path $env:SystemRoot -ChildPath 'System32\expand.exe'
    if (-not (Test-Path -LiteralPath $expandPath -PathType Leaf)) {
        throw "Windows expand.exe was not found: $expandPath"
    }

    foreach ($download in $downloads) {
        $row = [pscustomobject][ordered]@{
            SiteUrl = $download.SiteUrl
            LibraryName = $download.LibraryName
            TemplateUrl = $download.TemplateUrl
            ConfiguredBy = $download.ConfiguredBy
            SharePointVersion = $download.SharePointVersion
            IsCurrentVersion = $download.IsCurrentVersion
            Modified = $download.Modified
            XsnRelativePath = $download.XsnRelativePath
            XsnPath = ''
            DownloadStatus = $download.Status
            SolutionVersion = ''
            SolutionName = ''
            ManifestPath = ''
            ResourcesRelativePath = ''
            LabelsRelativePath = ''
            LabelCount = 0
            ViewCount = 0
            Sha256 = ''
            Status = 'Failed'
            Error = ''
        }
        $rows.Add($row)
        try {
            Write-XsnMessage -Message "Processing SharePoint version $($row.SharePointVersion): $($row.TemplateUrl)"
            if ($download.Status -ne 'Success') {
                throw "SPS download was not successful (status='$($download.Status)'): $($download.Error)"
            }
            $relativePath = $download.XsnRelativePath.Replace('/', '\')
            if ([string]::IsNullOrWhiteSpace($relativePath) -or [System.IO.Path]::IsPathRooted($relativePath)) {
                throw "XsnRelativePath must be a nonempty relative path: '$relativePath'"
            }
            $xsnPath = [System.IO.Path]::GetFullPath((Join-Path -Path $inputRoot -ChildPath $relativePath))
            if (-not $xsnPath.StartsWith($inputRoot.TrimEnd('\') + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "XsnRelativePath escapes the copied export folder: '$relativePath'"
            }
            if (-not $xsnPath.EndsWith('.xsn', [System.StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path -LiteralPath $xsnPath -PathType Leaf)) {
                throw "Downloaded XSN file is missing or is not an XSN: $xsnPath"
            }
            $row.XsnPath = $xsnPath
            $row.Sha256 = (Get-FileHash -LiteralPath $xsnPath -Algorithm SHA256).Hash
            $manifestDirectory = Join-Path -Path $outputRoot -ChildPath ('Manifests\{0:D4}' -f $rows.Count)
            $manifest = Read-XsnManifest -XsnPath $xsnPath -ExpandPath $expandPath -ManifestDirectory $manifestDirectory
            $row.SolutionVersion = $manifest.SolutionVersion
            $row.SolutionName = $manifest.SolutionName
            $row.ManifestPath = $manifest.ManifestPath
            if ($ExtractResources) { $row.ResourcesRelativePath = 'Manifests\{0:D4}\Resources' -f $rows.Count }
            $row.LabelsRelativePath = 'Manifests\{0:D4}\FieldLabels.json' -f $rows.Count
            $row.LabelCount = $manifest.LabelCount
            $row.ViewCount = $manifest.ViewCount
            $row.Status = 'Success'
            Write-XsnMessage -Message "solutionVersion=$($row.SolutionVersion); form=$($row.SolutionName); labels=$($row.LabelCount)"
        }
        catch {
            $failedCount++
            $row.Error = $_.Exception.Message
            Write-XsnMessage -Group Error -Message $row.Error
        }
    }
    if ($failedCount -gt 0) {
        throw "XSN processing incomplete: $failedCount failure(s). See '$csvPath' and '$logPath'."
    }
    Write-XsnMessage -Group Done -Message "Processed $($rows.Count) XSN version(s)."
}
catch {
    Write-XsnMessage -Group Error -Message $_.Exception.Message
    throw
}
finally {
    if ($rows.Count -gt 0) {
        $rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
    }
    $logEntries | Set-Content -LiteralPath $logPath -Encoding UTF8
}
