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

- Embedded InfoPath attachments (base64 payloads) are extracted to files.
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

Output:

- `InfoPathXsltGroupMap.csv`
- Template files: `tmpl__[full_version]__[major.minor.patch].xslt`
- Corresponding `.css` and `.log`

Notes:

- Calls `CreateXSLTfromInfopathFile.ps1` internally.
- Chooses representative sample with highest field count per group.

### 6) Transform XML to HTML

Script: `ProcessInfopathToHTML.ps1`

Input:

- `-InfoPathDirectory`
- `-MetadataCsvPath` (Stage 1 CSV)
- `-GroupingCsvPath` (Stage 4 CSV)
- `-TemplateMappingCsvPath` (Stage 5 CSV)
- `-OutputDirectory`
- `-PrintToPdf` (optional)
- `-Recurse` (optional)

Output:

- HTML files mirroring source folder layout
- Optional PDFs
- Copied CSS assets
- `ProcessInfopathToHTML-errors.log`
- `FailedInfoPathFiles` folder for failures

Attachment behavior (updated):

- Embedded attachments are decoded and saved beside generated HTML in `<baseName>_attachments` folders.
- XML values are rewritten to `ATTACHMENT::...` markers before transform.
- Generated templates render these markers as clickable relative links.

## XSLT Generator Update

Script: `CreateXSLTfromInfopathFile.ps1`

Update:

- Generated XSLT now detects values prefixed with `ATTACHMENT::`.
- Such values are rendered as `<a href="relative/path">...</a>` instead of plain text.

## Recommended Execution

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
   |- <mirrored folders and .html/.pdf>
   |- <baseName>_attachments
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

### Attachment links show but file missing

- Check attachment extraction folders under:
- `Reports/Normalized` for property path output
- `Html/<baseName>_attachments` for HTML output
- Confirm source XML payload is valid InfoPath attachment format.

### Empty extraction output

- Check `ParseStatus` in `InfoPathVersionMap.csv`.
- Ensure `Analysis` folder contains `StructureReport_[Version].json` for detected versions.
- Verify `-SourceDirectory` and `-Recurse` usage.
