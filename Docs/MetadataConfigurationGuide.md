# Metadata Configuration Guide

[Back to project README](../README.md) | [Getting Started](GettingStarted.md) | [Report Header Reference](ReportHeaderReference.md) | [`tblReportSequence` Reference](TblReportSequenceReference.md) | [Troubleshooting](Troubleshooting.md) | [Medicare Pipeline](WorkedExamples/MedicarePipeline.md) | [KPI Pipeline](WorkedExamples/KPIPipeline.md)


## Purpose

This guide explains how the active metadata tables control workbook layout, formulas, formatting, repeating report blocks, subtotals, and grand totals in the MetadataDrivenReportingEngine.

The active metadata tables are:

- `tblGoodFieldName_Header` for the one-time worksheet-level report header
- `tblGoodFieldName_Report` for ordinary repeating grouped report blocks
- `tblGoodFieldName_G_Totals` for the final report-level grand-total block

The worksheet header is documented separately because it is rendered once above the repeating report bands and requires sufficient reserved worksheet rows. See the [Report Header Metadata Reference](ReportHeaderReference.md) for its complete field definitions, formulas, merged ranges, source hyperlinks, formatting, and `beginSheet` prerequisite.

The guide reflects the current Medicare reference renderer and the earlier KPI renderer. Where the schema contains fields that the current code does not consume, those fields are identified as reserved, legacy, or unimplemented rather than presented as active features.

## 1. Worksheet-Level Header Metadata

`tblGoodFieldName_Header` controls the one-time header area at the top of each worksheet.

It can create and format:

- Worksheet titles and subtitles
- Static labels and explanatory text
- Excel formulas
- Public-source hyperlinks
- Generated-by information
- Generated date and time values
- Merged cell ranges
- Fonts, fills, row heights, and borders
- Freeze-below behavior

The worksheet-level header is different from the H1 zone in `tblGoodFieldName_Report`:

```text
tblGoodFieldName_Header
    -> rendered once at the top of the worksheet

tblGoodFieldName_Report H1
    -> rendered once for every grouped report band
```
The Medicare implementation requires enough reserved rows above the first generated report band. In `fn_FindBreaks`, the generated report area must begin at row 10 or later:

```vb
Const original = 2
Const beginSheet = 10
```

## 2. Before Editing Metadata

Metadata in this project behaves like executable configuration. A small change can affect the physical Excel layout, formula placement, inserted rows, merged ranges, or the visible boundaries of a report section.

Use this workflow when creating or changing a report:

1. Copy a working metadata configuration.
2. Substitute a small representative source, preferably the 99-row test source used by the Medicare demo.
3. Change one related group of settings at a time.
4. Render the workbook.
5. Inspect headers, detail ranges, formulas, footers, and totals.
6. Correct the metadata.
7. Repeat until the small test is stable.
8. Restore the full source and perform a volume test.

Do not begin by populating every metadata field. Start with the minimum required column definitions, then add headers, formulas, footers, and shared formatting in stages.

## 3. Report-Band Anatomy

A standard generated report band contains five logical zones:

```text
H1   Primary group header
B    Body or column-heading row
B_D  Detail-data rows
F1   Primary subtotal or summary row
F2   Extended footer or additional summary rows
```

A simplified band looks like this:

```text
+--------------------------------------------------+
| H1: Primary group header                         |
|     Static labels and values copied from data    |
+--------------------------------------------------+
| B:  Column headings                              |
+--------------------------------------------------+
| B_D: Detail record 1                             |
|      Detail record 2                             |
|      ...                                         |
|      Detail record n                             |
+--------------------------------------------------+
| F1: Primary subtotal or summary                  |
+--------------------------------------------------+
| F2 position 0: Extended summary row              |
| F2 position 1: Extended summary row              |
| F2 position 2: Extended summary row              |
+--------------------------------------------------+
```

The physical row numbers are calculated at runtime from the sorting tables, record counts, default band allocations, and any valid header or footer overrides.

## 4. Metadata Has Two Dimensions

A metadata record usually corresponds to an Excel column, but not every metadata field is column-specific. The table must be read in two dimensions.

### 4.1 Vertical dimension: one record per Excel column

The `Letter` field identifies the destination Excel column. Many settings are processed separately for each metadata record.

Examples include:

- `Letter`
- `ColumnWidth`
- `Column_Hide`
- `FieldName_B`
- B_D detail formatting
- `Range_B_D`
- `FieldName_F1`
- `Range_F1`

This is the normal column-oriented interpretation:

```text
Metadata row for K -> configure Excel column K
Metadata row for L -> configure Excel column L
Metadata row for M -> configure Excel column M
```

### 4.2 Horizontal dimension: one value controls a report section

Some settings describe an entire H1, F2, body boundary, or other range spanning multiple Excel columns.

For these operations, the renderer may:

1. Find the first visible Excel column.
2. Find the last visible Excel column.
3. Build one range across the visible report width.
4. Move the metadata recordset back to its first record.
5. Read the shared setting from the first record in `Sort` order.

The visible boundaries and the authoritative settings record are related but not identical:

- The section range is generally bounded by the first and last visible Excel columns.
- The shared property is often read from the first metadata record in `Sort` order.

The safest practice is to make the first sorted metadata record a visible column and place section-wide settings there.

### 4.3 Metadata behavior classes

This guide uses four behavior labels:

- **C, Column-specific:** evaluated separately for each Excel column
- **I, Repeatable item:** multiple positioned label/value items may be defined
- **S, Section-wide:** controls a multi-column report-zone range
- **U, Unused or unconfirmed:** present in the schema but not consumed by the current reference renderer

Examples:

```text
C  Letter, ColumnWidth, Column_Hide, Range_B_D, Range_F1
I  FieldName_H1_A, Position_H1_A, FieldName_F2_A, Position_F2_A
S  Merge_Center_H1, ColorIndex_H1, H1 and F2 border settings
U  List_B, ErrorTitle_B, ErrorMessage_B, Lock_B
```

## 5. Core Identification and Column Fields

Each metadata record requires enough information to associate the record with a report, worksheet, processing order, and Excel destination.

### `Sort`

- **Behavior:** C
- **Purpose:** Controls metadata processing order.
- **Recommended type:** Long Integer
- **Important:** The first record in `Sort` order can become the authoritative record for section-wide settings.

### `Report`

- **Behavior:** C
- **Purpose:** Identifies the report or workbook definition.
- **Example:** `Monthly_Report.XLSX`

### `Tab`

- **Behavior:** C
- **Purpose:** Identifies the worksheet configuration.
- **Example:** `KPI_Monthly`

### `Letter`

- **Behavior:** C
- **Purpose:** Identifies the destination Excel column.
- **Example:** `K`, `L`, `AA`
- **Important:** The value must be a valid Excel column label.

### `ColumnWidth`

- **Behavior:** C
- **Purpose:** Sets the width of the destination Excel column.
- **Default behavior:** The renderer may use a default width when the metadata value is null.

### `Column_Hide`

- **Behavior:** C
- **Purpose:** Hides the entire Excel column after it is created.
- **Recommended use:** Hide helper columns with `Column_Hide` rather than setting the width to zero.
- **Important:** Hidden columns are excluded when the renderer calculates the visible left and right boundaries of section-wide ranges.

## 6. Common Accepted Values

### 6.1 Yes/No fields

Use the Access Yes/No data type and checkbox values. Do not store the words `Yes` and `No` in text fields unless the schema explicitly requires text.

### 6.2 Horizontal alignment

Common supported values are:

- `Left`
- `Center`
- `Right`

### 6.3 Vertical alignment

Common supported values are:

- `Top`
- `Center`
- `Bottom`

### 6.4 Line weight

Common supported values are:

- `Thin`
- `Medium`
- `Thick`

### 6.5 Font styles

Common supported values are:

- `Regular`
- `Bold`
- `Italic`
- `Bold Italic`

### 6.6 Colors

Colors explicitly handled in current formatting branches include:

- `None`, where supported
- `Black`
- `Red`
- `Green`
- `Blue`
- `Gold`
- `Light Gold`
- `Light Brown`

Unsupported text values generally fall into a default branch. Verify every new color against the current VBA before relying on it.

### 6.7 Excel number formats

Format metadata is passed to Excel's `NumberFormat` property. Examples include:

```text
General
0
#,##0
$#,##0.00
0.00%
m/d/yyyy
```

Test custom formats in the generated workbook because Access stores the format as text while Excel interprets the final pattern.

## 7. H1 Primary Header Metadata

The H1 zone provides the primary header for each report band.

H1 contains two types of behavior:

1. Repeatable label/value items, evaluated from individual metadata records
2. Section-wide styling and merge behavior, usually controlled by the first sorted metadata record

### 7.1 Repeatable H1 item fields

#### `HeaderCol_H1_A`

- **Behavior:** I
- **Purpose:** Enables the H1 label/value behavior for the current metadata record.
- **Use:** Set to Yes only on records intended to contribute an H1 item.

#### `FieldName_H1_A`

- **Behavior:** I
- **Purpose:** Supplies the static H1 label.
- **Example:** `Medicare Report for TX:`

#### `Add_Name_From_Data_H1_A`

- **Behavior:** I
- **Purpose:** Identifies the source Excel column from which the H1 value is copied.
- **Example:** `I`

#### `Position_H1_A`

- **Behavior:** I
- **Purpose:** Positions the H1 item vertically within the allocated H1 rows.
- **Important:** The position must fit inside the primary-header allocation.

#### H1 label formatting

The following fields format the static label for the current H1 item:

- `Alignment_Horiz_H1_A`
- `Alignment_Vert_H1_A`
- `Font_H1_A`
- `Font_Size_H1_A`
- `Font_Style_H1_A`
- `Font_Color_H1_A`
- `Format_H1_A`
- `Wrap_Text_H1_A`
- `Underline_H1_A`

#### H1 copied-value formatting

The `_H1_AD` fields format the value copied from report data:

- `Font_H1_AD`
- `Font_Size_H1_AD`
- `Font_Style_H1_AD`
- `Font_Color_H1_AD`
- `Format_H1_AD`
- `Wrap_Text_H1_AD`
- `Underline_H1_AD`

### 7.2 Section-wide H1 fields

The following fields should be treated as section-wide and configured on the first sorted metadata record:

- `Merge_Center_H1`
- `ColorIndex_H1`
- `Row_Height_H1`
- `Repeat_YN_H1`
- `LineWeight_H1`
- `LineTop_H1`
- `LineBottom_H1`
- `LineLeft_H1`
- `LineRight_H1`
- `LineHoriz_H1`
- `LineVert_H1`

#### `Merge_Center_H1`

- **Behavior:** S
- **Purpose:** Combines configured H1 content and merges it across the visible report width.
- **Important:** The current logic evaluates this after repositioning the recordset to its first row. Setting it on a later metadata record may have no effect.

#### H1 fill, height, and borders

These settings apply to an H1 range bounded by the first and last visible Excel columns. Place the intended values on the first sorted metadata record.

## 8. B Column-Heading Metadata

The B zone is the column-heading row immediately above the detail data.

Most B fields are column-specific.

### `FieldName_B`

- **Behavior:** C
- **Purpose:** Supplies the visible column heading.
- **Example:** `Service Count`

### `Add_Name_From_Data_B`

- **Behavior:** C
- **Purpose:** Supplies an alternate heading value from report data when configured.

### B formatting fields

- `Alignment_Horiz_B`
- `Alignment_Vert_B`
- `Format_B`
- `ColorIndex_B`
- `Font_B`
- `Font_Size_B`
- `Font_Style_B`
- `Font_Color_B`
- `Wrap_Text_B`
- `Underline_B`
- `Row_Height_B`
- `LineWeight_B`
- `LineTop_B`
- `LineBottom_B`
- `LineLeft_B`
- `LineRight_B`

### B boundary fields

The `_B_Bline` fields control the outer boundary spanning the visible heading range:

- `LineWeight_B_Bline`
- `LineTop_B_Bline`
- `LineBottom_B_Bline`
- `LineLeft_B_Bline`
- `LineRight_B_Bline`

These boundary settings are section-oriented. Configure them consistently on the authoritative first metadata record.

## 9. B_D Detail-Region Metadata

The B_D zone formats the data rows for the current column.

### 9.1 Detail formatting fields

- `Alignment_Horiz_B_D`
- `Alignment_Vert_B_D`
- `Format_B_D`
- `ColorIndex_B_D`
- `Font_B_D`
- `Font_Size_B_D`
- `Font_Style_B_D`
- `Font_Color_B_D`
- `Wrap_Text_B_D`
- `Underline_B_D`
- `Row_Height_B_D`
- `LineWeight_B_D`
- `LineTop_B_D`
- `LineBottom_B_D`
- `LineLeft_B_D`
- `LineRight_B_D`
- `LineHoriz_B_D`

Most of these fields are applied to the detail range in the current metadata record's Excel column.

### 9.2 `Range_B_D`

- **Behavior:** C
- **Purpose:** Stores a detail formula template.
- **Important:** Store the formula without a leading equals sign.
- **Example:** `IFERROR(M?/L?,0)`

The renderer replaces `?` with the first detail row and assigns the resulting formula to the selected detail range. Excel adjusts relative references through the range.

### 9.3 `MergeCenter_B_D`

- **Behavior:** S within each detail row
- **Purpose:** Defines a row-relative merge range.
- **Example:** `A?:C?`

The renderer loops through the detail rows, replaces `?` with the current row number, and merges each resulting range.

Use this feature cautiously. Merged detail cells can complicate sorting, filtering, copying, and later formula maintenance.

### 9.4 Detail boundary fields

The `_B_Bline_D` fields control the outer boundary around the full visible detail block:

- `LineWeight_B_Bline_D`
- `LineTop_B_Bline_D`
- `LineBottom_B_Bline_D`
- `LineLeft_B_Bline_D`
- `LineRight_B_Bline_D`

These should be treated as section-wide boundary settings.

## 10. F1 Primary Footer Metadata

The F1 zone is the primary subtotal or summary row for each report band.

### `FieldName_F1`

- **Behavior:** C
- **Purpose:** Supplies a static subtotal label or value.
- **Example:** `Totals:`

### `Add_Name_From_Data_F1`

- **Behavior:** C
- **Purpose:** Copies a value from a configured source column rather than using a static label.

### `Range_F1`

- **Behavior:** C
- **Purpose:** Stores a footer formula template.
- **Important:** Store the formula without a leading equals sign.
- **Examples:**

```text
SUM(K@:K#)
IFERROR(M?/L?,0)
```

### F1 formatting fields

- `Alignment_Horiz_F1`
- `Alignment_Vert_F1`
- `Format_F1`
- `ColorIndex_F1`
- `Font_F1`
- `Font_Size_F1`
- `Font_Style_F1`
- `Font_Color_F1`
- `Wrap_Text_F1`
- `Underline_F1`
- `Row_Height_F1`
- `LineWeight_F1`
- `LineTop_F1`
- `LineBottom_F1`
- `LineLeft_F1`
- `LineRight_F1`

### F1 boundary fields

- `LineWeight_F1_Bline`
- `LineTop_F1_Bline`
- `LineBottom_F1_Bline`
- `LineLeft_F1_Bline`
- `LineRight_F1_Bline`

Treat the boundary fields as section-wide settings.

## 11. F2 Extended-Footer Metadata

The F2 zone supports additional positioned summary rows below F1.

### 11.1 Repeatable F2 item fields

#### `HeaderCol_F2_A`

- **Behavior:** I
- **Purpose:** Enables the F2 label/value behavior for the current metadata record.

#### `FieldName_F2_A`

- **Behavior:** I
- **Purpose:** Supplies the static F2 label.
- **Example:** `Run. Count of Samples`

#### `Add_Name_From_Data_F2_A`

- **Behavior:** I
- **Purpose:** Identifies the source Excel column from which the displayed F2 value is retrieved.

#### `GetFromSum_F2_A`

- **Behavior:** I
- **Purpose:** Selects which summary context supplies the copied value.
- **Important:** Verify the exact result against the current renderer when adapting this field.

#### `Position_F2_A`

- **Behavior:** I
- **Purpose:** Selects the vertical position within the allocated F2 rows.
- **Important:** The position must fit inside the available footer allocation.

#### F2 label formatting

- `Alignment_Horiz_F2_A`
- `Alignment_Vert_F2_A`
- `Font_F2_A`
- `Font_Size_F2_A`
- `Font_Style_F2_A`
- `Font_Color_F2_A`
- `Format_F2_A`
- `Wrap_Text_F2_A`
- `Underline_F2_A`

#### F2 copied-value formatting

- `Font_F2_AD`
- `Font_Size_F2_AD`
- `Font_Style_F2_AD`
- `Font_Color_F2_AD`
- `Format_F2_AD`
- `Wrap_Text_F2_AD`
- `Underline_F2_AD`

### 11.2 Section-wide F2 fields

Treat the following as section-wide settings on the first sorted metadata record:

- `Merge_Center_F2`
- `ColorIndex_F2`
- `Row_Height_F2`
- `Repeat_YN_F2`
- `LineWeight_F2`
- `LineTop_F2`
- `LineBottom_F2`
- `LineLeft_F2`
- `LineRight_F2`

The renderer finds the visible left and right boundaries before applying shared F2 fill and border behavior.

## 12. Formula Tokens

The formula-token system allows one metadata expression to work in report bands with different physical row numbers.

### 12.1 General rule

Store formulas without the leading equals sign:

```text
SUM(K@:K#)
```

The VBA adds `=` when assigning the formula to Excel.

### 12.2 `?` in `Range_B_D`

`?` is replaced with the first detail row, `Begin1`.

```text
IFERROR(M?/L?,0)
```

If the detail block begins at row 18, Excel initially receives:

```excel
=IFERROR(M18/L18,0)
```

Excel then adjusts relative references through the selected detail range.

### 12.3 `?` in `Range_F1`

In the F1 routine, `?` resolves to the row context defined by that footer calculation. Test the output when adapting a formula because the footer context differs from the B_D detail context.

### 12.4 `?` in `MergeCenter_B_D`

`?` is replaced with each current detail-row number while the renderer loops through the band.

```text
A?:C?
```

becomes:

```text
A18:C18
A19:C19
A20:C20
```

### 12.5 `@` and `#` in `Range_F1`

- `@` represents the first detail row, `Begin1`.
- `#` represents the final detail row, `End1`.

```text
SUM(K@:K#)
```

For rows 18 through 35, Excel receives:

```excel
=SUM(K18:K35)
```

## 13. Hidden Helper Columns

A metadata record may define an Excel column that does not exist in the SQL source. The column can receive a formula and then be hidden with `Column_Hide`.

Suitable uses include:

- Ratios and rates
- Running values
- Sort ranks
- Conditional flags
- Display labels
- Chart-support values
- Intermediate formula results

A recommended pattern is:

```text
Prepared SQL fields
    -> metadata-generated helper formula
    -> hidden helper column
    -> visible F1 or F2 result
```

Do not place business-critical filtering or grouping logic only in a hidden Excel column. Values required to define the recordset or group order belong in SQL.

## 14. KPI Running-Count Pattern

The KPI demo uses hidden helper columns to produce running values exposed in the extended footer.

A representative workbook formula is:

```excel
=SUMIFS(
    $K$1:INDEX(K:K,ROW()-1),
    $J$1:INDEX(J:J,ROW()-1),
    "<>Totals:"
)
```

The formula works because:

- `ROW()-1` ends the range immediately above the formula.
- `INDEX` creates a dynamic endpoint.
- The criterion excludes rows labeled `Totals:`.
- Detail values are counted while inserted subtotal rows are ignored.

This pattern supports report-wide accumulation across dynamically sized bands. A variant can use tokenized starting rows when a running value should restart within a band.

Always verify the actual value and criterion columns used by the current report. The formula above illustrates the pattern and may require different column letters in another worksheet.

## 15. Header and Footer Allocation

The renderer defines default allocations for primary and secondary headers and footers.

In the current reference implementation:

- A null `Over_ride_head1` uses the default primary-header allocation.
- A null `Over_ride_foot1` uses the default primary-footer allocation.
- Secondary header and footer allocations retain their configured defaults.

An override must provide enough rows for every configured position. For example, an F2 item at a later position cannot fit inside an override that allocates too few footer rows.

Before reducing an allocation:

1. Review every H1 or F2 position used by the worksheet metadata.
2. Review merged ranges.
3. Review row heights.
4. Review formula and copied-value placement.
5. Test against several report bands.

## 16. Grand-Total Metadata

`tblGoodFieldName_G_Totals` defines the final report-level summary block.

Grand-total metadata is separate because the final block summarizes the completed report rather than an individual grouping band.

Depending on the active configuration, the table may control:

- Grand-total labels
- Final formulas
- Number formats
- Fonts and colors
- Borders and fills
- Final summary calculations

The renderer creates the grand-total block after processing the ordinary sorting groups and then recursively calls the formatting logic using the grand-total metadata.

When adapting grand totals:

1. Confirm `HasTotals` is enabled in the worksheet/report mapping.
2. Confirm the grand-total metadata uses the correct `Report` and `Tab` values.
3. Confirm every formula references the intended final worksheet rows.
4. Validate that hidden helper columns contain calculated values before the grand totals are evaluated.

## 17. Active, Reserved, and Unimplemented Fields

A field's presence in the table schema does not prove that the current VBA consumes it.

The following fields are present in metadata or related schemas but are not confirmed as implemented in the current reference renderer:

- `PasteSpecial`
- `PasswordProtect`
- `Password`
- `List_B`
- `ErrorTitle_B`
- `ErrorMessage_B`
- `Lock_B`

Treat these fields as **U, unused or unconfirmed**.

Do not document them as active features until the renderer includes explicit code to read, apply, and test them.

## 18. Common Configuration Mistakes

### 18.1 Putting a section-wide setting on a later row

**Symptom:** The expected H1 merge, fill, row height, or border does not appear.

**Cause:** The renderer reads the shared setting from the first metadata record after determining the visible boundaries.

**Correction:** Move the intended section-wide value to the first record in `Sort` order.

### 18.2 Hiding the authoritative first record

**Symptom:** The visible boundaries and shared properties appear inconsistent.

**Cause:** Visible range discovery and authoritative-record selection are separate operations.

**Correction:** Prefer a visible first sorted record and use later hidden records for helper formulas.

### 18.3 Storing formulas with `=`

**Symptom:** Invalid or duplicated formula syntax.

**Cause:** The renderer already adds the leading equals sign.

**Correction:** Store `SUM(K@:K#)`, not `=SUM(K@:K#)`.

### 18.4 Using the wrong token for the zone

**Symptom:** A formula points to the wrong row.

**Cause:** `?` is context-specific, while `@` and `#` identify the beginning and ending detail rows in supported footer formulas.

**Correction:** Test the formula in the specific metadata field where it will be used.

### 18.5 Insufficient header or footer allocation

**Symptom:** Overlapping labels, missing footer values, or misplaced formulas.

**Cause:** A position exceeds the allocated H1 or F2 rows.

**Correction:** Increase the allocation or reduce the configured position.

### 18.6 Using unimplemented metadata fields

**Symptom:** A populated field has no effect.

**Cause:** The schema contains a reserved or legacy field that the renderer does not read.

**Correction:** Confirm implementation in `clsExport` before relying on the field.

### 18.7 Incorrect source-column letter

**Symptom:** H1 or F2 displays the wrong value.

**Cause:** `Add_Name_From_Data_*` points to the wrong Excel column after the source is transferred and columns are inserted.

**Correction:** Validate the final worksheet column map, not only the Access field order.

## 19. Recommended Configuration Sequence

Use this order when creating metadata for a new worksheet.

### Phase 1: Core columns

1. Add one metadata row per required Excel column.
2. Set `Sort`, `Report`, `Tab`, and `Letter`.
3. Set `ColumnWidth` and `Column_Hide`.
4. Render a plain worksheet and verify column alignment.

### Phase 2: Body and details

1. Add `FieldName_B` values.
2. Configure B formatting.
3. Configure B_D formats, alignment, and borders.
4. Render and inspect several groups.

### Phase 3: Primary headers

1. Set H1 label/value fields only on required metadata rows.
2. Add H1 positions.
3. Place section-wide H1 settings on the first sorted metadata row.
4. Test merged and unmerged behavior separately.

### Phase 4: Primary footers

1. Add `FieldName_F1` labels.
2. Add `Range_F1` subtotal formulas.
3. Validate `@`, `#`, and `?` replacement.
4. Check number formats and borders.

### Phase 5: Extended footers and helpers

1. Add hidden helper columns if needed.
2. Add formulas to the helper columns.
3. Add F2 label/value items.
4. Verify every F2 position fits inside the allocation.

### Phase 6: Grand totals and charts

1. Configure `tblGoodFieldName_G_Totals`.
2. Enable `HasTotals` where required.
3. Verify calculation order.
4. Verify the chart routines reference final worksheet ranges.

## 20. Validation Checklist

Before accepting a configuration, verify:

- [ ] `Report` matches the selected report definition.
- [ ] `Tab` matches the worksheet mapping.
- [ ] `Sort` values are unique and intentional.
- [ ] Every `Letter` is a valid Excel column.
- [ ] The first sorted metadata record contains section-wide settings.
- [ ] Visible left and right boundaries are correct.
- [ ] Hidden helper columns are intentionally hidden.
- [ ] H1 positions fit inside the header allocation.
- [ ] F2 positions fit inside the footer allocation.
- [ ] Formula metadata omits the leading equals sign.
- [ ] `?`, `@`, and `#` resolve correctly.
- [ ] Subtotal rows do not double-count prior totals.
- [ ] Grand totals calculate after helper formulas.
- [ ] The workbook is fully recalculated before save.
- [ ] Reserved fields are not being treated as active features.

## 21. Practical Formula Examples

### Detail ratio

```text
IFERROR(M?/L?,0)
```

### Detail payment rate

```text
IFERROR(N?/M?,0)
```

### Detail reduction percentage

```text
IFERROR(1-(M?/L?),0)
```

### Band sum

```text
SUM(K@:K#)
```

### Band currency sum

```text
SUM(L@:L#)
```

### Row-relative merge

```text
A?:C?
```

### Running total excluding subtotal rows

```text
SUMIFS($K$1:INDEX(K:K,ROW()-1),$J$1:INDEX(J:J,ROW()-1),"<>Totals:")
```

These examples are patterns, not universal formulas. Replace column letters and criteria with values appropriate to the target worksheet.

```markdown ## 22. Related Documentation - ../README.md - [Getting Started](GettingStarted.md) - [Architectural White Paper](WhitePaper.md) - [Report Header Metadata Reference](ReportHeaderReference.md) - [`tblReportSequence` Reference](TblReportSequenceReference.md) - [Troubleshooting Guide](Troubleshooting.md) - WorkedExamples/MedicarePipeline.md - WorkedExamples/KPIPipeline.md
