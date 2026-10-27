# InfoPath Processing Pipeline Documentation

## Overview

The pipeline processes local InfoPath XML files for two outputs:

- Normalized CSV extraction for analytics and mapping.
- HTML (and optional PDF) rendering using generated XSLT templates, or optionally
  the original XSN view adapted for a read-only archive.

SharePoint scripts run separately from this local pipeline. The local orchestrator script is `RunInfoPathPipeline.ps1`.

## SharePoint vs Local

SharePoint retrieval/metadata steps are not part of local orchestration:

- `MoveLibraryFilesToLocal_SPS.ps1`
- `GetMetadataforLib_SPS.ps1`
- [GetLibraryXsnVersions_SPS.ps1](./GetLibraryXsnVersions_SPS.ps1)
- [GetXsnReferenceData_SPS.ps1](./GetXsnReferenceData_SPS.ps1) (optional reference snapshots)

Run these separately on SPS when needed, then run the local pipeline against the downloaded XML directory.

### Export the library's XSN template versions

Run [GetLibraryXsnVersions_SPS.ps1](./GetLibraryXsnVersions_SPS.ps1) locally on a
SharePoint Server in **Windows PowerShell 5.1 / SharePoint Management Shell**,
using an account with access to the library, its configured templates, and
their version histories. Requires the `Microsoft.SharePoint.PowerShell`
snap-in. The server performs downloads only; it does not require `expand.exe`
or InfoPath and does not inspect XSN contents.

```powershell
.\GetLibraryXsnVersions_SPS.ps1 `
  -SiteUrl "http://sharepoint/sites/Finance" `
  -LibraryName "Applications" `
  -LocalPath "D:\SPExports\Templates"
```

The script:

- Discovers the library's configured document template and its content types'
  configured templates. Ignores non-XSN templates and exports shared URLs only once.
- Downloads the current XSN and every available historical SharePoint file
  version, including available minor versions.
- Creates a fresh `XsnTemplates_<timestamp>_<run-id>` folder on each run.
  Per-template URL-hash folders contain `template_sp-<file-version>.xsn`.
- Writes `XsnTemplateDownloads.csv` and `GetLibraryXsnVersions.log`. The CSV
  includes template URL, configuration sources, SharePoint file version,
  current/history flag, modification date, relative and original server XSN
  paths, download status, and error details. `Success` here means downloaded,
  not extracted or validated as an InfoPath package.
- Continues with other versions/templates after an individual failure, preserves
  the partial inventory, and throws if any downloads failed.

### Optional offline lookup labels and user names

The local HTML pipeline can consume a read-only SharePoint reference-data
snapshot. Neither HTML renderer authenticates to SharePoint or runs XSN code.
This option affects HTML/PDF display only; source XML and normalized CSVs retain
their stored values. It remains optional, including when XSN rendering is off.

**Collect everything on SPS first; no pipeline run or server/offline round trip
is needed to obtain reference data.**

1. On SPS, export XSN versions using the server downloader above.
2. Still on SPS, discover dependencies directly from that download export:

   ```powershell
   .\GetXsnReferenceDependencies.ps1 `
     -XsnExportCsvPath "D:\SPExports\Templates\XsnTemplates_20261027_101341_179ba7ca" `
     -OutputPath "D:\SPExports\ReferenceDependencies.json"
   ```

   Replace the example timestamped folder with your actual export folder.
   `-XsnExportCsvPath` also accepts `XsnTemplateDownloads.csv` itself, or a parent
   containing exactly one download inventory recursively. Multiple exports
   require selecting a specific folder/CSV.

   Discovery automatically calls the existing XSN processor to extract only
   manifests/views and metadata into a unique temporary folder. Windows
   `expand.exe` is required; InfoPath is not. It does not execute package code,
   load DLLs, contact SharePoint, or alter the original download export. Temporary
   extraction is removed on success and failure; dependency JSON and its adjacent
   `.log` retain results and preparation diagnostics. Keep
   `ProcessXsnTemplates.ps1` and the `Helper` folder alongside the discovery
   script on SPS (copying the script folder together is recommended).

   **`-OutputPath` is a JSON filename, not a folder.** Its parent folders are
   created automatically.

3. Still on SPS, run the companion exporter in Windows PowerShell 5.1 /
   SharePoint Management Shell under an authorized account:

   ```powershell
   .\GetXsnReferenceData_SPS.ps1 `
     -DependencyManifestPath "D:\SPExports\ReferenceDependencies.json" `
     -LocalPath "D:\SPExports\ReferenceData"
   ```

4. Copy the **original complete XSN download export**, the **complete
   reference-data export folder**, and the downloaded form XML to the offline
   processing machine. No return to SPS is required. For a first pipeline run:

   ```powershell
   .\RunInfoPathPipeline.ps1 `
     -SourceDirectory "D:\Data\InfoPathXml" `
     -OutputRootDirectory "D:\Data\PipelineOutput" `
     -XsnExportCsvPath "D:\Data\XsnExport" `
     -ReferenceDataPath "D:\Data\ReferenceDataExport" `
     -UseXSN4HTML $true `
     -StartStep 1
   ```

   Local pipeline step 1 prepares its own persistent XSN metadata and rendering
   resources; it does not depend on the server's deleted temporary extraction.
   If earlier pipeline outputs already exist, resume with `-StartStep 7`.
   Omit `-UseXSN4HTML` to apply the same reference data to generated HTML. The
   copied dependency definitions let generated HTML resolve values even without
   the original XSN packages on that machine.

Existing input modes remain available without new extraction:

```powershell
# Previously processed inventory.
.\GetXsnReferenceDependencies.ps1 `
  -XsnInventoryCsvPath "D:\Data\PipelineOutput\XSNTemplates\XsnTemplateVersions.csv" `
  -OutputPath "D:\Data\ReferenceDependencies.json"

# Already-extracted individual XSN.
.\GetXsnReferenceDependencies.ps1 `
  -ExtractedXsnDirectory "D:\Data\InfopathData\XSNSample" `
  -TemplateUrl "http://sharepoint/sites/Finance/Applications/Forms/template.xsn" `
  -OutputPath "D:\Data\ReferenceDependencies.json"
```

[GetXsnReferenceDependencies.ps1](./GetXsnReferenceDependencies.ps1) records
template identities/versions, namespaces, list connections, requested internal
columns, and supported bound value/display pairs across views. Country-style
controls can store `Value` and display `Title`; the exporter does not assume
every lookup uses `ID`. Review discovery warnings before exporting: unsupported
expressions are not guessed.

[GetXsnReferenceData_SPS.ps1](./GetXsnReferenceData_SPS.ps1) exports paged list
snapshots and site-collection user information, retaining source/provenance,
row counts, status, errors, and UTC export time in `Manifest.json`. Reused
connections are deduplicated. It does not create users, change SharePoint data,
expand groups, or query Active Directory. Partial exports remain available for
inspection, but export failures are reported explicitly.
List and user-source URLs may point to subwebs rather than site-collection roots.
The exporter opens each exact requested web with `Get-SPWeb`, obtains the owning
collection from that web, and exports user information once per collection.
It does not pass subweb URLs to `Get-SPSite` or silently substitute a parent web.

Resolution prefers saved display text, including matching cached rows in the
XML, before using an exported label. Existing explicit HTML field-rule value
mappings remain intentional overrides. Unresolved, missing, or conflicting
values stay readable as their original codes/accounts, with a warning and
details in `Html\InfoPathReferenceResolutions.csv`. That audit records reference
resolution before any explicit field-rule override. List keys are scoped to
site/list and value column; user IDs are scoped to a site collection. Login
matching preserves claims-provider identities; only the known Windows claims
prefix is treated as equivalent to its `DOMAIN\user` form.

Reference definitions follow the same exact/nearest-higher template selection
policy as XSN matching. When a processed XSN is selected, the reference definition
must agree with that selected version and package hash when available; missing
or ambiguous definitions are not substituted silently.

**Archival and privacy limits:** this is a current snapshot, not proof of the
label a list or user displayed years ago. Stored display text is stronger
historical evidence. Deleted or inaccessible users/items may remain unresolved.
Protect these exports like the form data: they contain list contents and user
identities, but no credentials. Keep a reference bundle with the archival output
if later reproducibility is required.

### Process the copied XSN export on the pipeline machine

Copy the **entire timestamped export folder**, including
`XsnTemplateDownloads.csv` and the per-template XSN folders, to the processing
machine. Keep their relative layout. The pipeline does not use the original
server's absolute paths, so the copied export can be on a different drive or
machine without editing the CSV.

`-XsnExportCsvPath` accepts the CSV itself, its containing folder, or a parent
folder such as `C:\ScriptTesting\XSNTemplates`. Folder inputs are searched
recursively for `XsnTemplateDownloads.csv`. Exactly one inventory must be found;
if multiple timestamped exports exist, supply the specific export folder or CSV.
The discovered CSV's directory remains the base for its relative XSN paths.

```powershell
.\RunInfoPathPipeline.ps1 `
  -SourceDirectory "D:\Data\InfoPathXml" `
  -OutputRootDirectory "D:\Data\PipelineOutput" `
  -XsnExportCsvPath "D:\Data\XsnExport\XsnTemplateDownloads.csv" `
  -StartStep 1
```

Step 1 runs [ProcessXsnTemplates.ps1](./ProcessXsnTemplates.ps1) locally.
It requires Windows `expand.exe`, but no SharePoint snap-in or InfoPath.
It extracts `manifest.xsf` and its declared view XSL files, reads
`solutionVersion` and the manifest's `name` form identity, recovers field
captions from view bindings, and computes the XSN file's SHA-256 on this machine.
Outputs are grouped under `OutputRootDirectory\XSNTemplates`:

- `XsnTemplateVersions.csv`: SharePoint metadata plus resolved local XSN path,
  `SolutionVersion`, `SolutionName`, `ManifestPath`, SHA-256, download status,
  processing status, view count, resolved label count, relative label-map path,
  and error details.
- `Manifests\<row-number>\manifest.xsf`: extracted manifests for successful rows.
- `Manifests\<row-number>\View*.xsl`: declared views retained for inspection.
- `Manifests\<row-number>\FieldLabels.json`: resolved, namespace-aware field labels.
- `FieldLabelCandidates.csv` and `FieldLabelResolution.csv` beside each manifest:
  binding/view/association details and ambiguity decisions, when labels are present.
- `ProcessXsnTemplates.log`: local processing messages and errors.

Failed downloads and invalid/missing packages remain failed in the processed
inventory. Other packages are processed, then the standalone processor throws
if any rows failed. The orchestrator warns and continues without the XSN inventory
for that run, preserving XML-driven CSV extraction and generated HTML. The source files and download CSV
are not changed. If `-XsnExportCsvPath` is omitted, step 1 explicitly reports
that it is skipped; XML-only runs remain supported. Steps 6 and 7 reuse an
existing `XSNTemplates\XsnTemplateVersions.csv` when resuming without the
original download CSV.

**Matching data XML to templates:** compare the XML processing instruction's
`mso-infoPathSolution solutionVersion` (the local classification report's
`TemplateVersion`) to the processed inventory's `SolutionVersion`, not its
`SharePointVersion`. For example, SharePoint file version `3.0` may contain
InfoPath solution version `1.0.0.27`. Also check the form identity (`name` in
the XML processing instruction versus `SolutionName`) and template URL (`href`
in the XML), because different forms can share the same numeric version.
Only use rows with `Status = Success`. Multiple SharePoint versions may contain
the same internal solution version; the inventory retains them separately.
During HTML generation, the same form's exact solution version is preferred.
If absent, the nearest available **higher numeric version** is used; for
example, `1.0.0.4` can use `1.0.0.10`, but never an older version or another
form's template. Within one solution version, a matching template URL is
preferred, then the highest SharePoint file version and latest modification
date. Without a form identity, a template URL or unambiguous filename is
required; ambiguous identities fall back to readable XML-name labels.

If the XML's form name has no eligible version but its `href` matches an exact
template URL, matching may use that URL despite a changed internal form name
(for example, `Applications` versus `template`). The eligible URL-matched records
must have one unambiguous form identity. Name-based matches remain preferred;
the fallback never uses filename-only similarity when a conflicting name exists,
and never selects an older version. Multiple eligible identities at the linked
URL are rejected rather than guessed.

Stage 7 logs a compact exact/next-match summary with the XML and selected XSN
versions, noting a template-URL identity fallback without printing full URLs or
form URNs. When no match is made, it reports available linked versions and the
failure reason. Full URLs, identities, and matching details remain in the audit
CSV. These are local inventory checks; the renderer does not contact the URL.

`Html\InfoPathXsnMatches.csv` records each data file's requested version,
identity, template href, exact/next-version/no-match result, selected version,
XSN and label-map paths, applied-label count, and errors. `MatchBasis`,
`LinkedTemplateVersions`, and `MatchedSolutionName` make URL-based identity
fallback and the available linked versions explicit. The generated
`Templates\InfoPathXsltGroupMap.csv` also records the XSN selected for each
template's sample. Runtime matching is per data file, not merely per group.

**Limits:** this exports only the configured template URLs and history still
retained and accessible in SharePoint. It cannot recover deleted history,
previously configured template URLs, or templates referenced only by old XML
forms. If template history was not retained, only the current XSN may be
available. The server downloader remains separate from the local pipeline.
By default, XSN captions are used in the generated archive HTML, but the original
InfoPath layout is not reproduced. Use the optional XSN rendering path below to
reuse its view layout without operational behavior. Older exports without `XsnRelativePath` need to be
regenerated with the updated SPS downloader before using this workflow.
Processed inventories created before caption extraction must be regenerated
with step 1 before they can supply XSN labels.

## Current Local Pipeline Order

The current order is dependency-correct and implemented in `RunInfoPathPipeline.ps1`:

1. `ProcessXsnTemplates.ps1` (when `-XsnExportCsvPath` is supplied)
2. `ClassifyInfoPathVersions.ps1`
3. `ProcessInfoPathAnalyze.ps1`
4. `ProcessInfoPathProperties.ps1`
5. `AnalyzeInfoPathGroupingFromCsv.ps1`
6. `BuildXsltFromGroupingCsv.ps1`
7. `ProcessInfopathToHTML.ps1`

Note: `AnalyzeInfoPathGroupingFromCsv.ps1` is not required before `ProcessInfoPathProperties.ps1`.

## Stage Details

### 1) Extract XSN template information

Script: [ProcessXsnTemplates.ps1](./ProcessXsnTemplates.ps1)

Input:

- `-InputCsvPath`: the copied `XsnTemplateDownloads.csv`, its containing folder,
  or a parent folder containing exactly one inventory recursively.
- XSN files beside the CSV, located using its `XsnRelativePath` column.
- `-OutputDirectory`: the pipeline's `XSNTemplates` output folder.

Output:

- `XsnTemplateVersions.csv`
- Extracted `manifest.xsf` files and `ProcessXsnTemplates.log`

Purpose:

- Process downloaded templates off-server and identify their internal
  InfoPath solution versions and form identities for matching to XML data.

### 2) Classify versions

Script: `ClassifyInfoPathVersions.ps1`

Input:

- `-SourceDirectory`
- `-Recurse` (optional)

Output:

- `InfoPathVersionMap.csv`

Purpose:

- Extract template version metadata and parse status for each XML file.

### 3) Analyze structure by version

Script: `ProcessInfoPathAnalyze.ps1`

Input:

- `-SourceDirectory`
- `-OutputDirectory`
- `-VersioningCsvPath` (from Stage 2, optional but recommended)
- `-MaxFilesToAnalyze`
- `-Recurse` (optional)

Output:

- `AnalysisSummary.csv`
- `StructureReport_[Version].json`
- `StructureRecommendations_[Version].csv`

Purpose:

- Discover structure and repeating groups per template version.

### 4) Extract normalized properties

Script: `ProcessInfoPathProperties.ps1`

Input:

- `-SourceDirectory`
- `-OutputDirectory`
- `-VersioningCsvPath` (from Stage 2)
- `-AnalysisDirectory` (from Stage 3)
- `-Recurse` (optional)

Output under `OutputDirectory/Normalized`:

- `Main.csv`
- One CSV per repeating group
- `PropertyMapping.csv`

Attachment behavior (updated):

- Embedded InfoPath attachment envelopes are decoded as raw bytes; documents, images, email messages, archives, and other file types are supported.
- The embedded filename and extension are preserved. An attachment without an extension is written with `.bin`.
- Field values are converted to markers: `ATTACHMENT::relative/path/to/file`.
- Attachment files are written to per-source attachment folders under `Normalized`.

Attachment metadata columns (updated):

- For each extracted field column `<ShortName>`, these columns are emitted:
- `<ShortName>_AttachSizeBytes`
- `<ShortName>_AttachMimeType`

PropertyMapping metadata (updated):

- `PropertyMapping.csv` now includes:
- `HasAttachmentMetadata`
- `AttachSizeBytesColumnName`
- `AttachMimeTypeColumnName`

### 5) Group files for template selection

Script: `AnalyzeInfoPathGroupingFromCsv.ps1`

Input:

- `-InputCsvPath` (typically `InfoPathVersionMap.csv`)

Output:

- `InfoPathGroupingDetailed.csv`
- `InfoPathGroupingSummary.csv`

Purpose:

- Build `SuggestedGroup` values for template generation and HTML transform lookup.

### 6) Build XSLT templates by group

Script: `BuildXsltFromGroupingCsv.ps1`

Input:

- `-GroupingCsvPath` (from Stage 5)
- `-TemplateOutputDirectory`
- `-UseDisplayName` (optional, defaults to `$true`)
- `-FieldRulesPath` (optional JSON HTML field definitions)
- `-XsnInventoryCsvPath` (optional processed XSN inventory from Stage 1)

Output:

- `InfoPathXsltGroupMap.csv`
- Template files: `tmpl__[full_version]__[major.minor.patch].xslt`
- Corresponding `.css` and `.log`

Notes:

- Calls `CreateXSLTfromInfopathFile.ps1` internally.
- Chooses representative sample with highest field count per group.
- Dotted version segments in the output base name are preserved. For example, `tmpl__1.0.0.3667__1.0.0` produces `tmpl__1.0.0.3667__1.0.0.xslt`.
- A mapping is marked successful only after its generated XSLT, CSS, and log files are verified on disk.
- New templates support readable/internal label modes and read-only boolean checkboxes.
- Date formatting is applied in Stage 7, not baked into the templates.

### 7) Transform XML to HTML

Script: `ProcessInfopathToHTML.ps1`

Input:

- `-InfoPathDirectory`
- `-MetadataCsvPath` (Stage 2 CSV)
- `-GroupingCsvPath` (Stage 5 CSV)
- `-TemplateMappingCsvPath` (Stage 6 CSV)
- `-OutputDirectory`
- `-PrintToPdf` (optional)
- `-UseDisplayName` (optional, defaults to `$true`)
- `-DateFormat` (optional, defaults to `MM/dd/yyyy`; HTML display only)
- `-DateTimeFormat` (optional, defaults to `MM/dd/yyyy HH:mm`; HTML display only)
- `-FieldRulesPath` (optional; use the same JSON definitions as Stage 6)
- `-ReferenceDataPath` (optional offline reference-data folder or `Manifest.json`)
- `-XsnInventoryCsvPath` (optional; the pipeline supplies Stage 1's inventory)
- `-UseXSN4HTML` (optional Boolean, defaults to `$false`; prefer the XSN view)
- `-Recurse` (optional)

Output:

- HTML files mirroring source folder layout
- Optional PDFs generated through headless Microsoft Edge or Google Chrome
- Copied CSS assets
- `ProcessInfopathToHTML-errors.log`
- `InfoPathXsnMatches.csv`
- `InfoPathReferenceResolutions.csv` (header-only when reference resolution is unused)
- `FailedInfoPathFiles` folder for failures

HTML display behavior:

- Readable field labels are enabled by default; use `-UseDisplayName $false` for internal XML names.
- ISO date-only values display as `MM/dd/yyyy`; timestamps display as
  `MM/dd/yyyy HH:mm`. Configure JSON `DateFormat` and `DateTimeFormat` to override.
- With newly generated templates, boolean values display as disabled checkboxes rather than `True`/`False` text.
- These display changes do not modify the original XML files or normalized CSV output.

Attachment behavior (updated):

- Embedded attachments are decoded and saved beside generated HTML in `<baseName>_attachments` folders.
- XML values are rewritten to `ATTACHMENT::...` markers before transform.
- Generated templates render these markers as clickable relative links.
- Attachment `href` values and stylesheet paths are relative to the generated HTML file, so the HTML, CSS, and attachment folders can be moved together.
- The original XML source path displayed in the HTML is absolute informational text, not a resource link.

PDF behavior:

- Set `-PrintToPdf true` to generate a PDF beside each HTML file.
- The script searches system and per-user installations of Microsoft Edge first, followed by Google Chrome.
- Common 64-bit and 32-bit `Program Files` locations are supported.
- Browser argument paths are explicitly quoted, and the HTML input uses an escaped
  `file:` URI, so source/output paths containing spaces are supported.
- Each conversion uses a unique temporary browser profile instead of the user's
  normal Edge/Chrome profile. First-run prompts and extensions are disabled.
- `--no-pdf-header-footer` suppresses the browser-added date, title, URL, and page
  headers/footers; content and styling from the HTML remain intact.
- The browser must exit successfully and produce a new PDF with a `%PDF-` signature
  and `%%EOF` end marker. File existence alone is not treated as success.
- The PDF is generated in a temporary location and moved beside the HTML only
  after validation. A stale PDF from an earlier run cannot mask a failed conversion;
  an existing destination PDF is replaced only after a new PDF passes validation.
- Temporary browser profiles and intermediate output are removed after conversion.
- Conversion failures are reported in `ProcessInfopathToHTML-errors.log`, and the
  source XML is copied to `FailedInfoPathFiles`. The generated HTML remains available.

## XSLT Generator Update

Script: `CreateXSLTfromInfopathFile.ps1`

Update:

- Generated XSLT now detects values prefixed with `ATTACHMENT::`.
- Such values are rendered as `<a href="relative/path">...</a>` instead of plain text.
- Boolean values render as disabled checkboxes, including in repeating tables.
- XHTML/rich-text formatting is rendered inside the owning form field rather
  than generating extra fields or sections for tags such as `div`, `span`, or `p`.
- Nested form groups inside repeating records retain their section and repeating
  table structure instead of being flattened into one text cell.

## Recommended Execution

### Pipeline parameters

[RunInfoPathPipeline.ps1](./RunInfoPathPipeline.ps1) accepts:

| Parameter | Required | Default | Purpose |
|-----------|----------|---------|---------|
| `-SourceDirectory` | Yes | None | Root folder containing InfoPath XML files. |
| `-OutputRootDirectory` | Yes | None | Root folder for reports, analysis, XSN information, XSLT templates, and HTML. |
| `-StartStep` | No | Prompt (`1`) | Start at step `1` through `7`, then run all remaining steps. Supplying this parameter bypasses the prompt. |
| `-XsnExportCsvPath` | No | None | Copied `XsnTemplateDownloads.csv` or a folder containing exactly one such inventory recursively; step 1 processes its XSNs under `XSNTemplates`. Omit to explicitly skip XSN processing. |
| `-Recurse` | No | Off | Include XML files in source subfolders. |
| `-MaxFilesToAnalyze` | No | `5` | Sample files analyzed per version; accepts `1` through `100`. |
| `-PrintToPdf` | No | `"false"` | Print HTML to PDF; accepts `yes`, `no`, `true`, `false`, `1`, or `0`. |
| `-UseXSN4HTML` | No | `$false` | Prefer an exact-version XSN view, otherwise the nearest higher version of the same form; fall back to generated HTML if unavailable or unsupported. |
| `-FieldRulesPath` | No | None | JSON field types, formats, and stored/display mappings for HTML/PDF. |
| `-ReferenceDataPath` | No | None | Offline list/user export folder or `Manifest.json`; affects HTML/PDF only. |

`-UseXSN4HTML` is a Boolean parameter, so explicitly pass `$true` or `$false`.
`-Recurse` is a switch. Pipeline labels are always readable; date formats now
belong in the JSON file supplied by `-FieldRulesPath`, not pipeline parameters.
Checkbox rendering is the default behavior of generated templates and has no
separate pipeline parameter.

### Optional XSN archival rendering

```powershell
.\RunInfoPathPipeline.ps1 `
  -SourceDirectory "D:\Data\InfoPathXml" `
  -OutputRootDirectory "D:\Data\InfoPathOutput" `
  -XsnExportCsvPath "D:\Data\CopiedExport\XsnTemplateDownloads.csv" `
  -UseXSN4HTML $true `
  -FieldRulesPath "D:\Data\HtmlFieldRules.json" `
  -PrintToPdf "true" `
  -StartStep 1
```

Copy the entire export from [GetLibraryXsnVersions_SPS.ps1](./GetLibraryXsnVersions_SPS.ps1).
The CSV's relative paths identify the XSN CAB files. Step 1 extracts declared
views and rendering resources into `XSNTemplates\Manifests\<row>\Resources`,
preserving their original relative names. Managed code is neither extracted
for rendering nor executed. Previously processed inventories without
`ResourcesRelativePath` require rerunning step 1 with `-UseXSN4HTML $true`.
Standalone extraction uses [ProcessXsnTemplates.ps1](./ProcessXsnTemplates.ps1)
with `-ExtractResources`.

The dedicated [XsnHtmlRendering.ps1](./Helper/XsnHtmlRendering.ps1) renderer:

- Prefers an exact solution-version match for the same form using the existing
  identity matching rules. If absent, uses the nearest available higher numeric
  solution version for both the view and captions, never an older version or
  another form. `NextVersion` messages identify the requested and selected versions;
  they do not mean the renderer has fallen back to generated HTML.
- Selects the manifest's default view and reuses its layout and captions.
- Replaces bound controls with stored values formatted by HTML field rules.
  Values render as borderless text without input backgrounds, shadows, or fixed
  field heights. Original typography and field widths are retained; modest
  `2px 4px` value padding and a `1.35` line height improve spacing and wrapping.
  Layout-table borders and section styling are not removed.
  Fixed InfoPath section/layout-region heights become minimum heights; sections
  expand with their visible content instead of overlapping subsequent sections.
  Checkboxes/radio buttons use static `[X]`/`[ ]` and `(X)`/`( )` markers,
  retaining saved selection tests rather than making every numeric field Boolean.
- Uses saved lookup text, configured display mappings, or the stored value/ID.
  It does not retrieve lookup lists or contact SharePoint.
- Uses stored calculated results, not bound-control calculation expressions.
- Removes conditional visibility gates around form content, hidden styles,
  operational buttons, scripts, and legacy InfoPath browser behaviors.
  Empty sections may remain visible. Missing repeating records are not invented.
- Copies referenced images, CSS, and supported fonts under
  `Html\Assets\<XSN-SHA256>`, shared by documents using the same package.
  References remain relative, including for HTML in nested source folders.
  Keep the entire HTML output folder together when moving the archive.
- Uses the same PDF printing process as generated HTML.

Missing same-form XSNs at the requested version or higher, missing assets, or unsupported view features
produce a warning and generated-layout fallback. `Html\InfoPathXsnMatches.csv`
records `RequestedRenderer`, `ActualRenderer`, `ViewName`, `FallbackReason`,
HTML/PDF paths, and errors alongside template-match details. XSN preparation
failures do not prevent CSV extraction. Invalid field rules still fail explicitly.
If rendering or PDF creation fails for any document, the script retains the
per-file error log and throws after processing the remaining files.

This is not an InfoPath runtime or a guarantee of pixel-identical rendering.
XSLT imports/includes, active CSS/imports, nonlocal rendering resources, and
remaining unsupported extension functions trigger fallback. View-specific print
settings are not currently translated; PDF pagination uses the browser.
Validate representative saved records before accepting an archival batch.

### One-command orchestration (recommended)

```powershell
.\RunInfoPathPipeline.ps1 `
  -SourceDirectory "C:\Data\InfoPathXml" `
  -OutputRootDirectory "C:\Data\InfoPathOutput" `
  -Recurse
```

Optional PDF output:

```powershell
.\RunInfoPathPipeline.ps1 `
  -SourceDirectory "C:\Data\InfoPathXml" `
  -OutputRootDirectory "C:\Data\InfoPathOutput" `
  -Recurse `
  -PrintToPdf "true"
```

### Starting or resuming at a specific step

At startup, [RunInfoPathPipeline.ps1](./RunInfoPathPipeline.ps1) lists all seven
numbered steps with their descriptions and script names. Enter a step number
to start there and run through step 7, or press Enter (or enter `1`) to run
the entire pipeline. Invalid entries prompt again.

Each step announces its number, description, and script before running and
reports completion afterward. Required-step errors stop the pipeline and identify
the failed step. Optional XSN preparation errors warn and continue without XSNs.

When resuming, use the same source and output folders and retain the required
outputs from earlier steps; skipped steps are not rerun or regenerated.
For automated runs, supply `-StartStep 1` to avoid the interactive prompt.
The new XSN extraction stage is step 1; former steps 1-6 are now steps 2-7.
For example, to rebuild templates and then render HTML:

```powershell
.\RunInfoPathPipeline.ps1 `
  -SourceDirectory "C:\Data\InfoPathXml" `
  -OutputRootDirectory "C:\Data\InfoPathOutput" `
  -StartStep 6
```

### HTML field labels

`BuildXsltFromGroupingCsv.ps1`,
`CreateXSLTfromInfopathFile.ps1`, and `ProcessInfopathToHTML.ps1` accept
`[bool] -UseDisplayName`, defaulting to `$true`.
The orchestrator no longer exposes this parameter and always uses readable
labels. The XSN rendering path preserves captions in the view regardless of
standalone generated-renderer label settings.

- `$true`: original captions recovered from the matched XSN when available.
  The XSN's default view takes precedence. Captions found only in other views
  are used when unambiguous; conflicting captions fall back to XML-name labels.
  Recovery supports explicit HTML label associations, ARIA labels, simple
  label/control table cells, and repeating-table column headers. Binding
  paths are namespace-aware and include repeating `xsl:for-each` contexts.
  Complex/dynamic bindings and layouts without a clear label association
  are not guessed; inspect the caption audit CSVs for unresolved fields.
  Without an XSN caption, readable labels are derived from XML names, such as `my:RequestorType`
  becoming `Requestor Type` and `HTTPResponse` becoming `HTTP Response`.
  Underscores are also replaced with spaces: `my:Requestor_Type` becomes
  `Requestor Type`. Existing capitalization is preserved.
  These fallback labels are not the actual captions configured in InfoPath.
- `$false`: exact internal XML names, including namespace prefixes.
- Applies to field labels, section/collection titles, and repeating-table field headers.
- XPath bindings, CSS classes, grouping, and normalized CSV names remain unchanged.
- Newly generated templates support either mode at transform time. Regenerate older
  templates (Stage 6) once before using this option with Stage 7, including to
  apply updated underscore-to-space labels or XSN caption support.

The pipeline passes its processed XSN inventory automatically. When running
the generator, builder, or renderer separately, supply
`-XsnInventoryCsvPath "D:\Data\PipelineOutput\XSNTemplates\XsnTemplateVersions.csv"`.
Direct XSLT transforms use the selected sample's embedded captions; use the
HTML renderer with the inventory to rematch captions for each data file.
The renderer disables embedded sample captions, so an unmatched file cannot
inherit a different form's labels.

To render internal names with generated HTML, run Stage 7 separately with
`-UseDisplayName $false`;
its default is `$true`, regardless of the default baked into the template.

Run `.\Helper\TestUseDisplayName.ps1` to verify both label modes end-to-end,
template defaults, transform overrides, and unchanged values and structural bindings.
The test uses temporary synthetic XML and removes its generated files afterward.

### HTML dates and boolean fields

Both renderers read top-level `DateFormat` and `DateTimeFormat` from the JSON
file supplied by `-FieldRulesPath`. The defaults are `MM/dd/yyyy` for dates and
`MM/dd/yyyy HH:mm` for timestamps. A formats-only file is supported:

```json
{
  "DateFormat": "dd/MM/yyyy",
  "DateTimeFormat": "dd/MM/yyyy HH:mm",
  "Fields": []
}
```

```powershell
.\RunInfoPathPipeline.ps1 `
  -SourceDirectory "C:\Data\InfoPathXml" `
  -OutputRootDirectory "C:\Data\InfoPathOutput" `
  -FieldRulesPath "C:\Data\HtmlFieldRules.json"
```

The standalone [ProcessInfopathToHTML.ps1](./ProcessInfopathToHTML.ps1) retains
its legacy date-format parameters for compatibility. Explicit standalone
parameters override JSON defaults; field-specific `Format` rules override both.

- Only ISO date (`yyyy-MM-dd`) and date/time leaf values are formatted for HTML,
  including nested fields and repeating-table values. Optional fractional seconds
  (up to seven digits) and `Z` or numeric time-zone offsets are supported.
- Formatting uses invariant culture without converting to the machine's time zone.
  Date/time values retain their time by default. Use JSON `DateTimeFormat` to include
  seconds or the stored offset, or to intentionally hide time.
- Source XML, normalized CSVs, metadata, and generated templates are unchanged.
  Empty values, invalid dates, and other text remain unchanged.
- When running Stage 7 separately, supply the same field-rules file; no template
  regeneration is needed to change date formatting.
- Empty format strings and invalid .NET format strings are rejected during
  configuration validation, before processing begins. Format tokens are case-sensitive:
  `MM` means month, while `mm` means minute.
- Newly generated templates display case-insensitive `true`/`false` field values
  as checked/unchecked, disabled checkboxes. Numeric `1`/`0` values render as
  checkboxes only when the element has an explicit `xsi:type` of `boolean`.
  Ordinary numeric fields remain text.
- Regenerate templates (Stage 6) to enable checkboxes in existing output, then
  rerun Stage 7. Checkboxes are read-only because these are archived forms.
- PDFs printed from the HTML inherit its displayed dates and checkboxes.

Timestamp display examples for the source value `2026-10-23T14:05:06-04:00`:

| JSON `DateTimeFormat` | Displayed value |
|---------------|-----------------|
| `"MM/dd/yyyy HH:mm"` (default) | `10/23/2026 14:05` |
| `"MM/dd/yyyy"` | `10/23/2026` |
| `"dd/MM/yyyy"` | `23/10/2026` |
| `"yyyy-MM-dd"` | `2026-10-23` |
| `"MM/dd/yyyy HH:mm:ss"` | `10/23/2026 14:05:06` |
| `"yyyy-MM-dd HH:mm:ss zzz"` | `2026-10-23 14:05:06 -04:00` |

For a date-only value such as `2026-10-23`, use JSON `DateFormat`; the default
output is `10/23/2026`. Move former pipeline date parameters into the field-rules
file when migrating existing commands.

Checkbox rules:

| XML field value | HTML display |
|-----------------|--------------|
| `true`, `True`, or `TRUE` | Checked, disabled checkbox |
| `false`, `False`, or `FALSE` | Unchecked, disabled checkbox |
| `1` with `xsi:type="xs:boolean"` | Checked, disabled checkbox |
| `0` with `xsi:type="xs:boolean"` | Unchecked, disabled checkbox |
| `1` or `0` without an explicit boolean type | Unchanged numeric text |
| Empty or `xsi:nil="true"` boolean field | No checkbox |

Surrounding whitespace is ignored when recognizing boolean values. The `xsi`
namespace is `http://www.w3.org/2001/XMLSchema-instance`; in the typed examples,
`xs` is bound to `http://www.w3.org/2001/XMLSchema`. Checkboxes include accessible
`True`/`False` labels. Scalar fields, nested leaf fields, repeated leaf values,
and leaf cells in repeating tables use the same rendering rules.

### Customer field-rendering rules

The HTML archive uses the following rules. Automatic handling is deliberately
limited to recognizable stored-value patterns; it does not guess whether an
untyped decimal represents money, a percentage, or an identifier.

| Customer field type | Archive behavior |
|---------------------|------------------|
| Attachment / File Upload | Decode existing InfoPath attachment envelopes. With an `Attachment` rule, also copy local file references or decode explicitly configured raw base64. Display a clickable filename, not the folder path. |
| Calculated Field | Preserve the stored result without evaluating formulas. Use `Calculated` to prevent automatic date/boolean/delimiter interpretation. |
| Checkbox / Yes-No | Display checked/unchecked, disabled checkboxes. A `Boolean` rule also recognizes numeric `1`/`0`. |
| Choice - Multiple | Split `Red;#Blue;#Green` into `Red, Blue, Green`. `MultiChoice` rules can map stored choices to display labels. |
| Currency | A `Currency` rule formats `1500.5` as `$1,500.50` with the default `en-US` culture. |
| Date & Time | Format `2026-10-21T14:30:00` as `10/21/2026 14:30`; keep the stored time, without local-time conversion. |
| Date Only | Format `2026-10-21` as `10/21/2026`. |
| Drop-down | Use `Dropdown` plus `Values` to resolve IDs; already displayed selections remain visible. |
| Empty Field | Keep the field consistently blank, including null/nil values; do not substitute placeholder text. |
| Hyperlink | Automatically link HTTP(S), mail, and telephone values, including `URL, Description`. `Hyperlink` rules can select URL/description children. |
| Lookup | Strip numeric SharePoint IDs: `15;#New York` becomes `New York`. Multiple ID/value pairs become comma-separated display values. |
| Multi-Line Text | Preserve source line breaks with `white-space: pre-wrap`; content remains escaped plain text. |
| Multiple Checkboxes | Render `Option1=true; Option2=false` as labeled, read-only selections, including unchecked options. Existing groups of boolean fields retain their own labels. |
| Number | A `Number` rule formats `1250` as `1,250`, preserving fractional digits by default. |
| Percentage | A `Percentage` rule formats the stored fraction `0.75` as `75%`. |
| Person / People Picker | A `Person` rule uses stored display names or configured account/name mappings. Unresolved accounts remain visible; no directory requests or invented names. |
| Radio Button | A `Radio` rule maps the stored code to the selected label, for example `Y` to `Yes`. |
| Repeating Section | Preserve each record and its nested sections/collections; mapped values are evaluated per record. |
| Repeating Table | Preserve field headers and rows. |
| Rich Text | Preserve readable formatting while filtering executable elements, event attributes, and unsafe link schemes as described below. |
| Signature Field | A `Signature` rule displays the stored signer and signed date. It does not verify or recreate a digital signature. |
| Single Line Text | Preserve the text as-is. Use `Text` for values that resemble dates, booleans, lookup serialization, or other automatic patterns. |

### Optional JSON field definitions

Copy and customize [HtmlFieldRules.example.json](./Examples/HtmlFieldRules.example.json).
Replace its example namespace URI and XPaths with your form's actual values:

```powershell
.\RunInfoPathPipeline.ps1 `
  -SourceDirectory "C:\Data\InfoPathXml" `
  -OutputRootDirectory "C:\Data\InfoPathOutput" `
  -FieldRulesPath "C:\Data\FieldRules.json" `
  -PrintToPdf "true"
```

The file is a JSON object with:

- `Culture`: .NET culture for numeric/currency formatting; defaults to `en-US`.
- `DateFormat` and `DateTimeFormat`: global .NET display formats, defaulting to
  `MM/dd/yyyy` and `MM/dd/yyyy HH:mm`; field-specific `Format` takes precedence.
- `Namespaces`: XPath prefix/URI pairs. Use an explicit prefix for namespaced XML,
  even when the source uses a default namespace. These prefixes need not match
  those used in the source document.
- `Fields`: an optional array (default empty) of objects with required `XPath` and `Type`. Select the form
  element containing the value, not an attribute or an XHTML formatting element.
  Rules can match multiple occurrences and can be shared across form versions;
  an XPath with no matches is allowed. Overlapping rules or rules selecting both
  a container and its descendants are rejected.
- Optional `Values`: stored/display-value pairs for choices, radio buttons, people,
  signatures, lookup values, and multiple-checkbox labels. Unmapped values remain
  visible rather than being guessed.
- Optional `Format`: a .NET format for `Currency`, `Number`, `Percentage`, `Date`,
  or `DateTime`. Defaults are `C2`, `#,##0.############################`, `0.##%`,
  JSON `DateFormat`, and `DateTimeFormat`, respectively. Percentage values are fractions,
  not already multiplied by 100.
- Optional relative child selectors: `UrlXPath`/`DescriptionXPath` for `Hyperlink`,
  `DisplayNameXPath` for `Person`, and required `SignerXPath`/`DateXPath` for
  `Signature`. They use the same namespace mappings as the field selector.
- `Attachment` rules default to local file-reference handling. For non-InfoPath
  raw base64, set `"Encoding": "base64"` and a filename such as
  `"FileName": "Invoice.pdf"`. InfoPath envelopes continue to use their embedded
  filenames without extra configuration.

Supported `Type` names are `Text`, `Calculated`, `Boolean`, `Currency`, `Number`,
`Percentage`, `Date`, `DateTime`, `Choice`, `MultiChoice`, `Dropdown`, `Radio`,
`Lookup`, `Person`, `Signature`, `Hyperlink`, `Attachment`, `MultipleCheckboxes`,
and `RichText`. Without definitions, ambiguous numeric values and account IDs
remain unchanged; automatic choice/lookup/date/link handling still works.

Local attachment references are resolved relative to each source XML file and
must remain within `-InfoPathDirectory` (the pipeline's source directory). Missing
files, references outside that root, and references traversing symbolic links or
junctions below that root fail explicitly. Files are copied into the
HTML attachment folder; HTTP(S) references remain external links. Filenames and
bytes are preserved, with unique destination names for collisions. Relative
attachment links escape spaces and URL-special characters.

Rules operate only on the in-memory rendering document. Source XML, normalized
CSVs, and pipeline metadata remain unchanged. Invalid JSON, unsupported types,
unknown rule properties, invalid format strings, and invalid typed values are
reported explicitly; per-file rendering errors use the normal error log and
failed-file copy mechanism.

Pass the same `-FieldRulesPath` to Stage 6 and Stage 7 when running them separately.
Adding/changing structured field types (for example a signature or structured
hyperlink) requires regenerating Stage 6 templates before Stage 7. Display-map,
culture, and format-only changes can be applied by rerunning Stage 7 with the
existing current templates. Direct XSLT execution does not load the JSON file;
use Stage 7 to apply these rules.

### Run HTML rendering separately

After the reports and templates exist, rerun only
[ProcessInfopathToHTML.ps1](./ProcessInfopathToHTML.ps1) to change date formatting,
label mode, or PDF output without repeating extraction:

```powershell
.\ProcessInfopathToHTML.ps1 `
  -InfoPathDirectory "C:\Data\InfoPathXml" `
  -MetadataCsvPath "C:\Data\InfoPathOutput\Reports\InfoPathVersionMap.csv" `
  -GroupingCsvPath "C:\Data\InfoPathOutput\Reports\InfoPathGroupingDetailed.csv" `
  -TemplateMappingCsvPath "C:\Data\InfoPathOutput\Templates\InfoPathXsltGroupMap.csv" `
  -OutputDirectory "C:\Data\InfoPathOutput\Html" `
  -Recurse `
  -UseDisplayName $false `
  -FieldRulesPath "C:\Data\FieldRules.json" `
  -PrintToPdf "true"
```

Stage 7 defaults to readable labels, `MM/dd/yyyy` dates, and `MM/dd/yyyy HH:mm`
timestamps each time it runs; pass preferred options and `-FieldRulesPath`
explicitly. Directly executing a generated XSLT outside Stage 7 supports its
built-in labels and checkboxes but does not apply date formats or field rules.

### Rich-text fields and form structure

- Elements in the XHTML namespace (`http://www.w3.org/1999/xhtml`) are recognized
  as formatting, whether they use a prefix or a default namespace. Common
  unqualified HTML tags are also recognized. Namespaced form elements such as
  `my:div` remain actual form data rather than being classified by tag name alone.
- A rich-text field has one form-field label. Paragraphs, inline text, emphasis,
  line breaks, lists, tables, links, images, and common formatting attributes are
  preserved as HTML within that field. XHTML namespace prefixes are removed from
  the displayed markup.
- Whitespace between inline elements is retained, so adjacent words do not run
  together. Mixed text surrounding inline markup is retained as well.
- Rich-text rendering is evaluated against each source document at transform
  time, including fields that were empty in the template sample.
- Repeated rich-text fields remain one value per row. Rich-text fields inside a
  repeating record remain in that record's column, and nested form groups and
  repeating collections are rendered within their owning record.
- Dates, boolean-like text, and attachment-like text inside XHTML formatting
  content are not reinterpreted as standalone form-field values. Normal
  date/boolean/attachment handling remains enabled for actual form fields.
- Executable elements such as scripts and iframes, event-handler attributes, and
  executable link schemes are not copied. Unknown formatting wrappers are
  unwrapped while retaining their text; this is not intended as a general-purpose
  HTML sanitizer or an exact recreation of the original InfoPath layout.
- Keep [HtmlRendering.ps1](./Helper/HtmlRendering.ps1) with the scripts; the
  generator and HTML processor share its namespace-aware classification rules.

Regenerate templates (Stage 6), then rerun Stage 7 to apply this fix to existing
output, or rerun the full pipeline. PDF output inherits the corrected HTML.

### Updating existing output

- To enable checkboxes or label-mode support in older templates, rerun Stage 6
  and then Stage 7, or rerun the full pipeline. Existing templates are not
  upgraded merely by rerunning Stage 7.
- To change only date formatting with current templates, update JSON `DateFormat`
  and/or `DateTimeFormat` and rerun Stage 7 with `-FieldRulesPath`;
  regeneration of XSLT/CSS is unnecessary.
- Generated templates continue to render `ATTACHMENT::` markers as relative links,
  including markers in repeated leaf values. Keep the HTML, CSS, and attachment
  folders together when moving the archive.

### Verification

Run the integration checks from the repository root:

```powershell
.\Helper\TestUseDisplayName.ps1
.\Helper\TestHtmlValueFormatting.ps1
.\Helper\TestPdfGeneration.ps1
.\Helper\TestRichTextRendering.ps1
.\Helper\TestCustomerFieldRendering.ps1
.\Helper\TestXsnHtmlRendering.ps1 -PrintToPdf
.\Helper\TestHtmlReferenceData.ps1
.\Helper\TestXsnReferenceExport.ps1
.\Helper\TestXsnReferenceDownloadDiscovery.ps1
```

[TestUseDisplayName.ps1](./Helper/TestUseDisplayName.ps1) verifies pipeline label
defaults, standalone internal-name mode, generated-template defaults, overrides,
and unchanged field values and structural bindings.

Run `.\Helper\TestHtmlValueFormatting.ps1` to verify date formatting and checkboxes
end-to-end, standalone Stage 7 overrides, and unchanged XML/CSV values.
[TestHtmlValueFormatting.ps1](./Helper/TestHtmlValueFormatting.ps1) also checks
exact repeated-date counts, time-zone offsets, checked/unchecked and read-only
states, ordinary numeric values, empty fields, invalid dates and format strings,
and attachment links. Both tests use synthetic XML in temporary directories and
remove their generated files afterward.

[TestPdfGeneration.ps1](./Helper/TestPdfGeneration.ps1) checks quoted browser
arguments, isolated temporary profiles, header/footer suppression, validation of
missing/empty/invalid/truncated PDF output, browser failures, and protection against
stale destination PDFs. It also runs the full pipeline with `-PrintToPdf "true"`
against the installed Edge/Chrome browser using source and output paths containing
spaces. A supported browser must be installed for this test.

[TestRichTextRendering.ps1](./Helper/TestRichTextRendering.ps1) verifies rich-text
formatting and whitespace, prefixed/default XHTML namespaces, unqualified markup,
genuine form groups, both label modes, empty-to-rich/plain value changes, and
nested/repeating record boundaries. It also checks that HTML date and boolean
processing does not alter rich-text prose, and that source XML remains unchanged.

[TestCustomerFieldRendering.ps1](./Helper/TestCustomerFieldRendering.ps1) verifies
the customer field rules, mappings and unresolved values, blank fields,
file-reference/base64 attachments, signature summaries, multiline text,
standalone overrides, configuration validation, unchanged source XML/CSVs, and
real browser PDF generation. A supported browser must be installed.

[TestHtmlReferenceData.ps1](./Helper/TestHtmlReferenceData.ps1) checks offline
stored-value/display-column mappings, list isolation, saved cached labels,
site-scoped user IDs, Windows claims aliases, saved person names, explicit
field-rule overrides, ambiguous/missing references, package-hash agreement,
and both generated/XSN HTML paths. It also checks raw radio selection tests,
unchanged XML/CSV, partial-export handling, and reference CSV path containment.
Supply `-ExtractedSampleDirectory` to additionally verify dependency discovery
and all four Country labels through the extracted sample's actual XSN view.
With an installed supported browser, `-PrintToPdf` also checks PDF generation
from both reference-enabled HTML paths.

[TestXsnReferenceExport.ps1](./Helper/TestXsnReferenceExport.ps1) exercises
dependency discovery and the SPS exporter with an in-process SharePoint facade,
including paging, deduplication, partial failures, and object disposal. This
does not replace an authorized integration run on a live SharePoint Server.
Its subweb regressions restrict `Get-SPSite` to collection roots, exercise
separate lookup and user-source webs under one collection, and reject missing
webs without exporting a parent's data.

[TestXsnReferenceDownloadDiscovery.ps1](./Helper/TestXsnReferenceDownloadDiscovery.ps1)
checks direct download-inventory/CAB inputs, recursive folder discovery, all
versions/views, original package hashes, partial/corrupt downloads, ambiguous
export folders, actionable output-path errors, retained diagnostics, and
temporary-directory cleanup on success and failure.

### Orchestrator output layout

```text
<OutputRootDirectory>
|- Reports
|  |- InfoPathVersionMap.csv
|  |- InfoPathGroupingDetailed.csv
|  |- InfoPathGroupingSummary.csv
|  \- Normalized
|     |- Main.csv
|     |- PropertyMapping.csv
|     |- <RepeatingGroup>.csv
|     \- <Per-file attachment folders>
|- Analysis
|  |- AnalysisSummary.csv
|  |- StructureReport_<Version>.json
|  \- StructureRecommendations_<Version>.csv
|- Templates
|  |- InfoPathXsltGroupMap.csv
|  |- tmpl__*.xslt
|  |- tmpl__*.css
|  \- tmpl__*.log
\- Html
  |- <mirrored source folders, when -Recurse is used>
  |- <baseName>.html
  |- <baseName>.pdf (when -PrintToPdf is enabled)
  |- <baseName>_attachments
  |- tmpl__*.css
   |- FailedInfoPathFiles
   \- ProcessInfopathToHTML-errors.log
```

## Dependency Notes

- `ProcessInfoPathProperties.ps1` depends on Stage 2 and Stage 3 outputs, not on grouping CSV.
- HTML rendering path depends on grouping and template mapping outputs.
- If templates are stale, rerun Stage 6 and Stage 7.

## Troubleshooting

### Missing template mapping or XSLT

- Confirm `InfoPathGroupingDetailed.csv` exists and has `SuggestedGroup` values.
- Confirm `InfoPathXsltGroupMap.csv` contains those groups.
- Confirm `XsltPath` files exist on disk.
- Rerun Stage 6 if a mapping refers to an older or missing template filename.

### Attachment links show but file missing

- Check attachment extraction folders under:
- `Reports/Normalized` for property path output
- `Html/<baseName>_attachments` for HTML output
- Confirm source XML payload is valid InfoPath attachment format.
- Keep each HTML file, its matching `<baseName>_attachments` folder, and the copied `tmpl__*.css` file together when moving output.

### Edge or Chrome not detected for PDF output

- Confirm Edge or Chrome is installed as a desktop application, not only available through a URL protocol or shortcut.
- Common locations include `C:\Program Files\Microsoft\Edge\Application\msedge.exe`, `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`, and per-user installations under `%LOCALAPPDATA%`.
- Run without `-PrintToPdf true` when only HTML output is needed.

### PrintToPdf is enabled but no new PDF appears

- Review `ProcessInfopathToHTML-errors.log` for the browser exit code, missing
  output, or PDF validation failure. HTML generation can succeed even if printing
  fails; check the final PDF count as well as the HTML count.
- Confirm the output folder and Windows temporary folder are writable, and that
  endpoint/browser policies allow headless printing with a temporary profile.
- Paths containing spaces are supported; keep path arguments quoted in your
  PowerShell command as shown in the examples.
- Existing PDFs are preserved if a new conversion fails, so an older PDF is not
  evidence that the latest run succeeded.
- After updating the script, rerun Stage 7 with `-PrintToPdf "true"`. Template
  regeneration is not required for the PDF fix.

### Empty extraction output

- Check `ParseStatus` in `InfoPathVersionMap.csv`.
- Ensure `Analysis` folder contains `StructureReport_[Version].json` for detected versions.
- Verify `-SourceDirectory` and `-Recurse` usage.

### Dates still show the original values

- Confirm the value is an ISO date or date/time leaf value; arbitrary date-like
  text is not reformatted.
- Run Stage 7 with the desired JSON `DateFormat` and `DateTimeFormat` using
  `-FieldRulesPath`; opening or directly transforming
  the XSLT does not perform Stage 7's in-memory date formatting.
- Date-only values use JSON `DateFormat`; timestamps use JSON `DateTimeFormat` and show
  hours/minutes by default. Include seconds/offset tokens when those should be visible.

### Boolean fields still show True/False text

- Regenerate the templates with Stage 6, then rerun Stage 7.
- Confirm the mapping points to the regenerated XSLT.
- Untyped numeric `1`/`0` values intentionally remain text to avoid treating
  ordinary numeric fields as booleans.

### HTML tags appear as field labels or rich-text content is flattened

- Regenerate Stage 6 templates with the updated generator, then rerun Stage 7.
  Updating only the HTML processor does not repair an older template's structure.
- Confirm the group mapping points to the regenerated XSLT.
- XHTML markup should use `http://www.w3.org/1999/xhtml`; common unqualified HTML
  tags are supported too. Elements in the form's own namespace are treated as
  form fields/groups, even if their local names match HTML tags.
