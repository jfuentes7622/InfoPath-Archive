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
- `-FieldRulesPath` (optional JSON HTML field definitions)

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
- `-DateTimeFormat` (optional, defaults to `MM/dd/yyyy HH:mm`; HTML display only)
- `-FieldRulesPath` (optional; use the same JSON definitions as Stage 5)
- `-Recurse` (optional)

Output:

- HTML files mirroring source folder layout
- Optional PDFs generated through headless Microsoft Edge or Google Chrome
- Copied CSS assets
- `ProcessInfopathToHTML-errors.log`
- `FailedInfoPathFiles` folder for failures

HTML display behavior:

- Readable field labels are enabled by default; use `-UseDisplayName $false` for internal XML names.
- ISO date-only values display as `MM/dd/yyyy`; timestamps display as
  `MM/dd/yyyy HH:mm`. Use `-DateFormat` and `-DateTimeFormat` respectively to override.
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
| `-OutputRootDirectory` | Yes | None | Root folder for reports, analysis, templates, and HTML. |
| `-StartStep` | No | Prompt (`1`) | Start at step `1` through `6`, then run all remaining steps. Supplying this parameter bypasses the prompt. |
| `-Recurse` | No | Off | Include XML files in source subfolders. |
| `-MaxFilesToAnalyze` | No | `5` | Sample files analyzed per version; accepts `1` through `100`. |
| `-PrintToPdf` | No | `"false"` | Print HTML to PDF; accepts `yes`, `no`, `true`, `false`, `1`, or `0`. |
| `-UseDisplayName` | No | `$true` | Use readable labels; `$false` displays internal XML names. |
| `-DateFormat` | No | `"MM/dd/yyyy"` | .NET date-only display format for HTML/PDF. |
| `-DateTimeFormat` | No | `"MM/dd/yyyy HH:mm"` | .NET timestamp display format; retains the stored clock time and offset. |
| `-FieldRulesPath` | No | None | JSON field types, formats, and stored/display mappings for HTML/PDF. |

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

### Starting or resuming at a specific step

At startup, [RunInfoPathPipeline.ps1](./RunInfoPathPipeline.ps1) lists all six
numbered steps with their descriptions and script names. Enter a step number
to start there and run through step 6, or press Enter (or enter `1`) to run
the entire pipeline. Invalid entries prompt again.

Each step announces its number, description, and script before running and
reports completion afterward. If a step throws an error, the pipeline stops
and reports the failed step so you can fix the problem and restart there.

When resuming, use the same source and output folders and retain the required
outputs from earlier steps; skipped steps are not rerun or regenerated.
For automated runs, supply `-StartStep 1` to avoid the interactive prompt.
For example, to rebuild templates and then render HTML:

```powershell
.\RunInfoPathPipeline.ps1 `
  -SourceDirectory "C:\Data\InfoPathXml" `
  -OutputRootDirectory "C:\Data\InfoPathOutput" `
  -StartStep 5
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

`RunInfoPathPipeline.ps1` accepts `-DateFormat` and `-DateTimeFormat` and passes
them only to `ProcessInfopathToHTML.ps1` (Stage 6). The defaults are `MM/dd/yyyy`
for dates and `MM/dd/yyyy HH:mm` for timestamps.
Use a .NET date/time format string, for example:

```powershell
.\RunInfoPathPipeline.ps1 `
  -SourceDirectory "C:\Data\InfoPathXml" `
  -OutputRootDirectory "C:\Data\InfoPathOutput" `
  -DateFormat "dd/MM/yyyy" `
  -DateTimeFormat "dd/MM/yyyy HH:mm"
```

- Only ISO date (`yyyy-MM-dd`) and date/time leaf values are formatted for HTML,
  including nested fields and repeating-table values. Optional fractional seconds
  (up to seven digits) and `Z` or numeric time-zone offsets are supported.
- Formatting uses invariant culture without converting to the machine's time zone.
  Date/time values retain their time by default. Use `-DateTimeFormat` to include
  seconds or the stored offset, or to intentionally hide time.
- Source XML, normalized CSVs, metadata, and generated templates are unchanged.
  Empty values, invalid dates, and other text remain unchanged.
- When running Stage 6 separately, supply both preferred format options there; no template
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

Timestamp display examples for the source value `2026-10-23T14:05:06-04:00`:

| `-DateTimeFormat` | Displayed value |
|---------------|-----------------|
| `"MM/dd/yyyy HH:mm"` (default) | `10/23/2026 14:05` |
| `"MM/dd/yyyy"` | `10/23/2026` |
| `"dd/MM/yyyy"` | `23/10/2026` |
| `"yyyy-MM-dd"` | `2026-10-23` |
| `"MM/dd/yyyy HH:mm:ss"` | `10/23/2026 14:05:06` |
| `"yyyy-MM-dd HH:mm:ss zzz"` | `2026-10-23 14:05:06 -04:00` |

For a date-only value such as `2026-10-23`, use `-DateFormat`; the default output
is `10/23/2026`. Earlier versions used `-DateFormat` for timestamps too; supply
`-DateTimeFormat` explicitly when migrating an existing custom timestamp format.

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
  -DateFormat "MM/dd/yyyy" `
  -DateTimeFormat "MM/dd/yyyy HH:mm" `
  -PrintToPdf "true"
```

The file is a JSON object with:

- `Culture`: .NET culture for numeric/currency formatting; defaults to `en-US`.
- `Namespaces`: XPath prefix/URI pairs. Use an explicit prefix for namespaced XML,
  even when the source uses a default namespace. These prefixes need not match
  those used in the source document.
- `Fields`: an array of objects with required `XPath` and `Type`. Select the form
  element containing the value, not an attribute or an XHTML formatting element.
  Rules can match multiple occurrences and can be shared across form versions;
  an XPath with no matches is allowed. Overlapping rules or rules selecting both
  a container and its descendants are rejected.
- Optional `Values`: stored/display-value pairs for choices, radio buttons, people,
  signatures, lookup values, and multiple-checkbox labels. Unmapped values remain
  visible rather than being guessed.
- Optional `Format`: a .NET format for `Currency`, `Number`, `Percentage`, `Date`,
  or `DateTime`. Defaults are `C2`, `#,##0.############################`, `0.##%`,
  `-DateFormat`, and `-DateTimeFormat`, respectively. Percentage values are fractions,
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

Pass the same `-FieldRulesPath` to Stage 5 and Stage 6 when running them separately.
Adding/changing structured field types (for example a signature or structured
hyperlink) requires regenerating Stage 5 templates before Stage 6. Display-map,
culture, and format-only changes can be applied by rerunning Stage 6 with the
existing current templates. Direct XSLT execution does not load the JSON file;
use Stage 6 to apply these rules.

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
  -DateFormat "MM/dd/yyyy" `
  -DateTimeFormat "MM/dd/yyyy HH:mm:ss" `
  -PrintToPdf "true"
```

Stage 6 defaults to readable labels, `MM/dd/yyyy` dates, and `MM/dd/yyyy HH:mm`
timestamps each time it runs; pass preferred options and `-FieldRulesPath`
explicitly. Directly executing a generated XSLT outside Stage 6 supports its
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

Regenerate templates (Stage 5), then rerun Stage 6 to apply this fix to existing
output, or rerun the full pipeline. PDF output inherits the corrected HTML.

### Updating existing output

- To enable checkboxes or label-mode support in older templates, rerun Stage 5
  and then Stage 6, or rerun the full pipeline. Existing templates are not
  upgraded merely by rerunning Stage 6.
- To change only date formatting with current templates, rerun Stage 6 with
  `-DateFormat` and/or `-DateTimeFormat`; regeneration of XSLT/CSS is unnecessary.
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
- After updating the script, rerun Stage 6 with `-PrintToPdf "true"`. Template
  regeneration is not required for the PDF fix.

### Empty extraction output

- Check `ParseStatus` in `InfoPathVersionMap.csv`.
- Ensure `Analysis` folder contains `StructureReport_[Version].json` for detected versions.
- Verify `-SourceDirectory` and `-Recurse` usage.

### Dates still show the original values

- Confirm the value is an ISO date or date/time leaf value; arbitrary date-like
  text is not reformatted.
- Run Stage 6 with the desired `-DateFormat` and `-DateTimeFormat`; opening or directly transforming
  the XSLT does not perform Stage 6's in-memory date formatting.
- Date-only values use `-DateFormat`; timestamps use `-DateTimeFormat` and show
  hours/minutes by default. Include seconds/offset tokens when those should be visible.

### Boolean fields still show True/False text

- Regenerate the templates with Stage 5, then rerun Stage 6.
- Confirm the mapping points to the regenerated XSLT.
- Untyped numeric `1`/`0` values intentionally remain text to avoid treating
  ordinary numeric fields as booleans.

### HTML tags appear as field labels or rich-text content is flattened

- Regenerate Stage 5 templates with the updated generator, then rerun Stage 6.
  Updating only the HTML processor does not repair an older template's structure.
- Confirm the group mapping points to the regenerated XSLT.
- XHTML markup should use `http://www.w3.org/1999/xhtml`; common unqualified HTML
  tags are supported too. Elements in the form's own namespace are treated as
  form fields/groups, even if their local names match HTML tags.
