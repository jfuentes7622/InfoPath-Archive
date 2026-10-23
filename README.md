# InfoPath Processing Pipeline Documentation

## Overview

The pipeline processes local InfoPath XML files for two outputs:

- Normalized CSV extraction for analytics and mapping.
- HTML (and optional PDF) rendering using generated XSLT templates.

SharePoint scripts run separately from this local pipeline. The local orchestrator script is `RunInfoPathPipeline.ps1`.

## SharePoint vs Local

SharePoint retrieval/metadata steps are not part of local orchestration:

- `MoveLibraryFilesToLocal_SPS.ps1`
- `GetMetadataforLib_SPS.ps1`

Run these separately on SPS when needed, then run the local pipeline against the downloaded XML directory.

## Current Local Pipeline Order

The current order is dependency-correct and implemented in `RunInfoPathPipeline.ps1`:

1. `ClassifyInfoPathVersions.ps1`
2. `ProcessInfoPathAnalyze.ps1`
3. `ProcessInfoPathProperties.ps1`
4. `AnalyzeInfoPathGroupingFromCsv.ps1`
5. `BuildXsltFromGroupingCsv.ps1`
6. `ProcessInfopathToHTML.ps1`

Note: `AnalyzeInfoPathGroupingFromCsv.ps1` is not required before `ProcessInfoPathProperties.ps1`.

## Stage Details

### 1) Classify versions

Script: `ClassifyInfoPathVersions.ps1`

Input:

- `-SourceDirectory`
- `-Recurse` (optional)

Output:

- `InfoPathVersionMap.csv`

Purpose:

- Extract template version metadata and parse status for each XML file.

### 2) Analyze structure by version

Script: `ProcessInfoPathAnalyze.ps1`

Input:

- `-SourceDirectory`
- `-OutputDirectory`
- `-VersioningCsvPath` (from Stage 1, optional but recommended)
- `-MaxFilesToAnalyze`
- `-Recurse` (optional)

Output:

- `AnalysisSummary.csv`
- `StructureReport_[Version].json`
- `StructureRecommendations_[Version].csv`

Purpose:

- Discover structure and repeating groups per template version.

### 3) Extract normalized properties

Script: `ProcessInfoPathProperties.ps1`

Input:

- `-SourceDirectory`
- `-OutputDirectory`
- `-VersioningCsvPath` (from Stage 1)
- `-AnalysisDirectory` (from Stage 2)
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

### 4) Group files for template selection

Script: `AnalyzeInfoPathGroupingFromCsv.ps1`

Input:

- `-InputCsvPath` (typically `InfoPathVersionMap.csv`)

Output:

- `InfoPathGroupingDetailed.csv`
- `InfoPathGroupingSummary.csv`

Purpose:

- Build `SuggestedGroup` values for template generation and HTML transform lookup.

### 5) Build XSLT templates by group

Script: `BuildXsltFromGroupingCsv.ps1`

Input:

- `-GroupingCsvPath` (from Stage 4)
- `-TemplateOutputDirectory`
- `-UseDisplayName` (optional, defaults to `$true`)

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
- Date formatting is applied in Stage 6, not baked into the templates.

### 6) Transform XML to HTML

Script: `ProcessInfopathToHTML.ps1`

Input:

- `-InfoPathDirectory`
- `-MetadataCsvPath` (Stage 1 CSV)
- `-GroupingCsvPath` (Stage 4 CSV)
- `-TemplateMappingCsvPath` (Stage 5 CSV)
- `-OutputDirectory`
- `-PrintToPdf` (optional)
- `-UseDisplayName` (optional, defaults to `$true`)
- `-DateFormat` (optional, defaults to `MM/dd/yyyy`; HTML display only)
- `-Recurse` (optional)

Output:

- HTML files mirroring source folder layout
- Optional PDFs generated through headless Microsoft Edge or Google Chrome
- Copied CSS assets
- `ProcessInfopathToHTML-errors.log`
- `FailedInfoPathFiles` folder for failures

HTML display behavior:

- Readable field labels are enabled by default; use `-UseDisplayName $false` for internal XML names.
- ISO date and date/time field values display as `MM/dd/yyyy` by default; use `-DateFormat` to override.
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

## XSLT Generator Update

Script: `CreateXSLTfromInfopathFile.ps1`

Update:

- Generated XSLT now detects values prefixed with `ATTACHMENT::`.
- Such values are rendered as `<a href="relative/path">...</a>` instead of plain text.
- Boolean values render as disabled checkboxes, including in repeating tables.

## Recommended Execution

### Pipeline parameters

[RunInfoPathPipeline.ps1](./RunInfoPathPipeline.ps1) accepts:

| Parameter | Required | Default | Purpose |
|-----------|----------|---------|---------|
| `-SourceDirectory` | Yes | None | Root folder containing InfoPath XML files. |
| `-OutputRootDirectory` | Yes | None | Root folder for reports, analysis, templates, and HTML. |
| `-Recurse` | No | Off | Include XML files in source subfolders. |
| `-MaxFilesToAnalyze` | No | `5` | Sample files analyzed per version; accepts `1` through `100`. |
| `-PrintToPdf` | No | `"false"` | Print HTML to PDF; accepts `yes`, `no`, `true`, `false`, `1`, or `0`. |
| `-UseDisplayName` | No | `$true` | Use readable labels; `$false` displays internal XML names. |
| `-DateFormat` | No | `"MM/dd/yyyy"` | .NET date/time format used only when rendering HTML and PDFs printed from it. |

`-UseDisplayName` is a Boolean parameter, so explicitly pass `$true` or `$false`
when overriding it in PowerShell. `-Recurse` is a switch.
Checkbox rendering is the default behavior of generated templates and has no
separate pipeline parameter.

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

### HTML field labels

`RunInfoPathPipeline.ps1`, `BuildXsltFromGroupingCsv.ps1`,
`CreateXSLTfromInfopathFile.ps1`, and `ProcessInfopathToHTML.ps1` accept
`[bool] -UseDisplayName`, defaulting to `$true`.

- `$true`: readable labels derived from XML names, such as `my:RequestorType`
  becoming `Requestor Type` and `HTTPResponse` becoming `HTTP Response`.
  These are not the actual labels configured in the original InfoPath form.
- `$false`: exact internal XML names, including namespace prefixes.
- Applies to field labels, section/collection titles, and repeating-table field headers.
- XPath bindings, CSS classes, grouping, and normalized CSV names remain unchanged.
- Newly generated templates support either mode at transform time. Regenerate older
  templates (Stage 5) once before using this option with Stage 6.

To render internal names:

```powershell
.\RunInfoPathPipeline.ps1 `
  -SourceDirectory "C:\Data\InfoPathXml" `
  -OutputRootDirectory "C:\Data\InfoPathOutput" `
  -UseDisplayName $false
```

When running Stage 6 separately, pass `-UseDisplayName $false` there as well;
its default is `$true`, regardless of the default baked into the template.

Run `.\Helper\TestUseDisplayName.ps1` to verify both label modes end-to-end,
template defaults, transform overrides, and unchanged values and structural bindings.
The test uses temporary synthetic XML and removes its generated files afterward.

### HTML dates and boolean fields

`RunInfoPathPipeline.ps1` accepts `-DateFormat` and passes it only to
`ProcessInfopathToHTML.ps1` (Stage 6). The default is `MM/dd/yyyy`.
Use a .NET date/time format string, for example:

```powershell
.\RunInfoPathPipeline.ps1 `
  -SourceDirectory "C:\Data\InfoPathXml" `
  -OutputRootDirectory "C:\Data\InfoPathOutput" `
  -DateFormat "dd/MM/yyyy"
```

- Only ISO date (`yyyy-MM-dd`) and date/time leaf values are formatted for HTML,
  including nested fields and repeating-table values. Optional fractional seconds
  (up to seven digits) and `Z` or numeric time-zone offsets are supported.
- Formatting uses invariant culture without converting to the machine's time zone.
  Include time components in the format (for example, `MM/dd/yyyy HH:mm:ss`)
  to display them; the default displays only the date.
- Source XML, normalized CSVs, metadata, and generated templates are unchanged.
  Empty values, invalid dates, and other text remain unchanged.
- When running Stage 6 separately, supply `-DateFormat` there; no template
  regeneration is needed to change date formatting.
- Empty format strings and invalid .NET format strings are rejected during
  parameter validation, before processing begins. Format tokens are case-sensitive:
  `MM` means month, while `mm` means minute.
- Newly generated templates display case-insensitive `true`/`false` field values
  as checked/unchecked, disabled checkboxes. Numeric `1`/`0` values render as
  checkboxes only when the element has an explicit `xsi:type` of `boolean`.
  Ordinary numeric fields remain text.
- Regenerate templates (Stage 5) to enable checkboxes in existing output, then
  rerun Stage 6. Checkboxes are read-only because these are archived forms.
- PDFs printed from the HTML inherit its displayed dates and checkboxes.

Date display examples for the source value `2026-10-23T14:05:06-04:00`:

| `-DateFormat` | Displayed value |
|---------------|-----------------|
| `"MM/dd/yyyy"` (default) | `10/23/2026` |
| `"dd/MM/yyyy"` | `23/10/2026` |
| `"yyyy-MM-dd"` | `2026-10-23` |
| `"MM/dd/yyyy HH:mm:ss"` | `10/23/2026 14:05:06` |
| `"yyyy-MM-dd HH:mm:ss zzz"` | `2026-10-23 14:05:06 -04:00` |

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
  -DateFormat "MM/dd/yyyy HH:mm:ss" `
  -PrintToPdf "true"
```

Stage 6 defaults to readable labels and `MM/dd/yyyy` each time it runs; pass your
preferred options explicitly. Directly executing a generated XSLT outside Stage 6
supports its built-in labels and checkboxes but does not apply `-DateFormat`.

### Updating existing output

- To enable checkboxes or label-mode support in older templates, rerun Stage 5
  and then Stage 6, or rerun the full pipeline. Existing templates are not
  upgraded merely by rerunning Stage 6.
- To change only date formatting with current templates, rerun Stage 6 with
  `-DateFormat`; regeneration of XSLT/CSS is unnecessary.
- Generated templates continue to render `ATTACHMENT::` markers as relative links,
  including markers in repeated leaf values. Keep the HTML, CSS, and attachment
  folders together when moving the archive.

### Verification

Run the integration checks from the repository root:

```powershell
.\Helper\TestUseDisplayName.ps1
.\Helper\TestHtmlValueFormatting.ps1
```

[TestUseDisplayName.ps1](./Helper/TestUseDisplayName.ps1) verifies pipeline label
defaults, internal-name mode, generated-template defaults, standalone overrides,
and unchanged field values and structural bindings.

Run `.\Helper\TestHtmlValueFormatting.ps1` to verify date formatting and checkboxes
end-to-end, standalone Stage 6 overrides, and unchanged XML/CSV values.
[TestHtmlValueFormatting.ps1](./Helper/TestHtmlValueFormatting.ps1) also checks
exact repeated-date counts, time-zone offsets, checked/unchecked and read-only
states, ordinary numeric values, empty fields, invalid dates and format strings,
and attachment links. Both tests use synthetic XML in temporary directories and
remove their generated files afterward.

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

- `ProcessInfoPathProperties.ps1` depends on Stage 1 and Stage 2 outputs, not on grouping CSV.
- HTML rendering path depends on grouping and template mapping outputs.
- If templates are stale, rerun Stage 5 and Stage 6.

## Troubleshooting

### Missing template mapping or XSLT

- Confirm `InfoPathGroupingDetailed.csv` exists and has `SuggestedGroup` values.
- Confirm `InfoPathXsltGroupMap.csv` contains those groups.
- Confirm `XsltPath` files exist on disk.
- Rerun Stage 5 if a mapping refers to an older or missing template filename.

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

### Empty extraction output

- Check `ParseStatus` in `InfoPathVersionMap.csv`.
- Ensure `Analysis` folder contains `StructureReport_[Version].json` for detected versions.
- Verify `-SourceDirectory` and `-Recurse` usage.

### Dates still show the original values

- Confirm the value is an ISO date or date/time leaf value; arbitrary date-like
  text is not reformatted.
- Run Stage 6 with the desired `-DateFormat`; opening or directly transforming
  the XSLT does not perform Stage 6's in-memory date formatting.
- Include time tokens if time should be visible. The default intentionally
  displays only the date.

### Boolean fields still show True/False text

- Regenerate the templates with Stage 5, then rerun Stage 6.
- Confirm the mapping points to the regenerated XSLT.
- Untyped numeric `1`/`0` values intentionally remain text to avoid treating
  ordinary numeric fields as booleans.
