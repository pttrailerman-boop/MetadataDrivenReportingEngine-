# Metadata Configuration Guide

## Purpose

This guide explains how the metadata tables control workbook layout, formulas, formatting, grouping zones, subtotals, and grand totals in the MetadataDrivenReportingEngine.

The two active configuration tables are:

- `tblGoodFieldName_Report`, which controls worksheet columns and the H1, B, B_D, F1, and F2 zones.
- `tblGoodFieldName_G_Totals`, which controls final grand-total output.

The implementation described here is based on the current `clsExport` rendering logic and the sample `KPI_Monthly` configuration.

## 1. Report Block Anatomy

Each grouping block is rendered as a set of logical worksheet zones:

```text
+--------------------------------------------------+
| H1: Primary group header                        |
| Static labels and values copied from data       |
+--------------------------------------------------+
| B: Body heading row                             |
| Column labels                                   |
+--------------------------------------------------+
| B_D: Detail-data rows                           |
| Record 1                                        |
| Record 2                                        |
| ...                                             |
| Record n                                        |
+--------------------------------------------------+
| F1: Primary subtotal or summary row             |
+--------------------------------------------------+
| F2 position 0: Nested-summary row 1             |
| F2 position 1: Nested-summary row 2             |
| F2 position 2: Nested-summary row 3             |
+--------------------------------------------------+
```

Field-name suffixes identify the affected zone:

| Suffix | Zone | Purpose |
|---|---|---|
| `_H1_A` | H1 label | Primary-header label and placement |
| `_H1_AD` | H1 data | Primary-header value copied from data |
| `_B` | B heading | Body or column-heading row |
| `_B_D` | B_D detail | Detail-data region |
| `_F1` | F1 footer | Primary subtotal or summary row |
| `_F2_A` | F2 label | Nested-footer label and placement |
| `_F2_AD` | F2 data | Nested-footer value copied from data |
| `_F2` | F2 panel | Overall nested-footer panel formatting |

## 2. Formula Tokens

Formula metadata is stored without a leading equals sign because the VBA adds `=` before assigning the formula to Excel.

### `?` token

The meaning of `?` depends on the metadata field.

#### In `Range_B_D`

`?` is replaced with the current block's first detail row, `begin1`.

Metadata:

```text
A?+B?
```

If the block begins at row 18, Excel receives:

```text
=A18+B18
```

When assigned across the detail range, Excel adjusts the relative references for later rows.

A block-based running calculation can use:

```text
SUMIFS($J$?:INDEX(J:J,ROW()-1),$I$?:INDEX(I:I,ROW()-1),"<>Totals:")
```

If the block begins at row 18, the starting references become `$J$18` and `$I$18`.

#### In `Range_F1`

`?` is replaced with `End1 + 1`, the row immediately following the last detail row.

Metadata:

```text
J?+K?
```

If `End1` is 35, Excel receives:

```text
=J36+K36
```

#### In `MergeCenter_B_D`

`?` is replaced with each current detail-row number while the code loops through the block.

Metadata:

```text
A?:C?
```

The engine produces ranges such as `A18:C18`, `A19:C19`, and `A20:C20`, merging each row separately.

### `@` token

In `Range_F1`, `@` represents the first detail row, `begin1`.

### `#` token

In `Range_F1`, `#` represents the last detail row, `End1`.

Metadata:

```text
SUM(J@:J#)
```

For a block running from row 18 through row 35, Excel receives:

```text
=SUM(J18:J35)
```

## 3. Common Accepted Values

### Yes/No fields

Use the Microsoft Access `Yes/No` data type with a checkbox.

- `No` is the normal default for optional behavior.
- Do not store the words `Yes` or `No` in a text field.

### Horizontal alignment

Accepted values used by the VBA:

- `Left`
- `Center`
- `Right`

### Vertical alignment

Accepted values used by the VBA:

- `Top`
- `Center`
- `Bottom`

### Line weight

Accepted values used by the VBA:

- `Thin`
- `Medium`
- `Thick`

### Colors

Accepted named values currently handled by the VBA include:

- `None`, where supported
- `Black`
- `Red`
- `Green`
- `Blue`
- `Gold`
- `Light Brown`
- `Light Gold`

Adding a new color name requires extending the applicable `Select Case` blocks in `clsExport`.

### Fonts and styles

Typical values include:

- Font: `Arial`
- Font style: `Regular`, `Bold`, `Italic`, or `Bold Italic`
- Font size: numeric points such as `10`, `12`, or `14`

### Excel number formats

Examples:

- `General`
- `0`
- `0.0%`
- `m/d/yyyy`
- Currency or custom Excel number-format strings

## 4. Core Record Fields

Each `tblGoodFieldName_Report` record generally represents one output column.

| Field | Suggested Access type | Purpose | Example |
|---|---|---|---|
| `Sort` | Number, Long Integer | Processing and display order | `10` |
| `Report` | Short Text | Workbook/report identifier | `Monthly_Report.XLSX` |
| `Tab` | Short Text | Worksheet name or tab identifier | `KPI_Monthly` |
| `Letter` | Short Text | Excel destination column | `J` |
| `Column_Hide` | Yes/No | Hides the worksheet column | `No` |
| `ColumnWidth` | Number, Double | Excel column width; Null normally defaults to 9 | `17.14` |

`Letter` must contain a valid Excel column label. A blank or invalid value produces an invalid whole-column reference. Use `Column_Hide` rather than a width of zero when a helper column should be hidden.

## 5. H1 Primary Header Fields

H1 fields control the primary header area above each major block.

| Field | Suggested type | Purpose | Example |
|---|---|---|---|
| `HeaderCol_H1_A` | Yes/No | Enables label/data header behavior | `Yes` |
| `FieldName_H1_A` | Short Text | Static header label | `Monthly KPI:` |
| `Add_Name_From_Data_H1_A` | Short Text | Source Excel column for the accompanying value | `Q` |
| `Position_H1_A` | Number, Integer | Vertical position within the H1 area | `0` |
| `Alignment_Horiz_H1_A` | Short Text/lookup | Label horizontal alignment | `Center` |
| `Alignment_Vert_H1_A` | Short Text/lookup | Label vertical alignment | `Top` |
| `Font_H1_A` | Short Text/lookup | Label font | `Arial` |
| `Font_Size_H1_A` | Number | Label font size | `12` |
| `Font_Style_H1_A` | Short Text/lookup | Label style | `Regular` |
| `Font_Color_H1_A` | Short Text/lookup | Label font color | `Black` |
| `Format_H1_A` | Short Text | Label number format | `General` |
| `Wrap_Text_H1_A` | Yes/No | Wraps label text | `Yes` |
| `Underline_H1_A` | Yes/No | Underlines label text | `No` |
| `Font_H1_AD` | Short Text/lookup | Data-value font | `Arial` |
| `Font_Size_H1_AD` | Number | Data-value font size | `14` |
| `Font_Style_H1_AD` | Short Text/lookup | Data-value style | `Bold` |
| `Font_Color_H1_AD` | Short Text/lookup | Data-value color | `Black` |
| `Format_H1_AD` | Short Text | Data-value number format | `General` |
| `Wrap_Text_H1_AD` | Yes/No | Wraps the copied value | `No` |
| `Underline_H1_AD` | Yes/No | Underlines the copied value | `No` |
| `Merge_Center_H1` | Yes/No | Merges and centers configured H1 content | `Yes` |
| `ColorIndex_H1` | Short Text/lookup | H1 background color | `Light Brown` |
| `Row_Height_H1` | Number | H1 row height | `24` |
| `Repeat_YN_H1` | Yes/No | Metadata flag for repeated header behavior | `No` |
| `LineWeight_H1` | Short Text/lookup | H1 border weight | `Medium` |
| `LineTop_H1` | Yes/No | Top border | `Yes` |
| `LineBottom_H1` | Yes/No | Bottom border | `Yes` |
| `LineLeft_H1` | Yes/No | Left border | `No` |
| `LineRight_H1` | Yes/No | Right border | `Yes` |
| `LineHoriz_H1` | Yes/No | Inside horizontal borders | `No` |
| `LineVert_H1` | Yes/No | Inside vertical borders | `No` |

## 6. B Body-Heading Fields

B fields control the column-heading row immediately above the detail data.

| Field | Suggested type | Purpose | Example |
|---|---|---|---|
| `FieldName_B` | Short Text | Displayed column heading | `Count Of Late WO` |
| `Add_Name_From_Data_B` | Short Text | Alternate column from which a heading value may be copied | Null |
| `Alignment_Horiz_B` | Short Text/lookup | Horizontal alignment | `Center` |
| `Alignment_Vert_B` | Short Text/lookup | Vertical alignment | `Center` |
| `Format_B` | Short Text | Heading number format | `General` |
| `ColorIndex_B` | Short Text/lookup | Heading background color | `None` |
| `Font_B` | Short Text/lookup | Heading font | `Arial` |
| `Font_Size_B` | Number | Heading font size | `10` |
| `Font_Style_B` | Short Text/lookup | Heading font style | `Bold Italic` |
| `Font_Color_B` | Short Text/lookup | Heading font color | `Black` |
| `Wrap_Text_B` | Yes/No | Wraps heading text | `No` |
| `Underline_B` | Yes/No | Underlines heading | `Yes` |
| `Row_Height_B` | Number | Body-heading row height | `17.25` |
| `LineWeight_B` | Short Text/lookup | Border weight | `Thin` |
| `LineTop_B` | Yes/No | Top border | `No` |
| `LineBottom_B` | Yes/No | Bottom border | `Yes` |
| `LineLeft_B` | Yes/No | Left border | `Yes` |
| `LineRight_B` | Yes/No | Right border | `Yes` |
| `LineWeight_B_Bline` | Short Text/lookup | Special body-heading boundary weight | `Medium` |
| `LineTop_B_Bline` | Yes/No | Boundary top line | `No` |
| `LineBottom_B_Bline` | Yes/No | Boundary bottom line | `No` |
| `LineLeft_B_Bline` | Yes/No | Boundary left line | `Yes` |
| `LineRight_B_Bline` | Yes/No | Boundary right line | `Yes` |

## 7. B_D Detail-Region Fields

B_D fields control formatting and formulas for the detail rows.

| Field | Suggested type | Purpose | Example |
|---|---|---|---|
| `Alignment_Horiz_B_D` | Short Text/lookup | Detail horizontal alignment | `Center` |
| `Alignment_Vert_B_D` | Short Text/lookup | Detail vertical alignment | `Center` |
| `Format_B_D` | Short Text | Detail Excel number format | `0.0%` |
| `ColorIndex_B_D` | Short Text/lookup | Detail background color | `None` |
| `Font_B_D` | Short Text/lookup | Detail font | `Arial` |
| `Font_Size_B_D` | Number | Detail font size | `10` |
| `Font_Style_B_D` | Short Text/lookup | Detail font style | `Regular` |
| `Font_Color_B_D` | Short Text/lookup | Detail font color | `Black` |
| `Wrap_Text_B_D` | Yes/No | Wraps detail values | `No` |
| `Underline_B_D` | Yes/No | Underlines detail values | `No` |
| `Row_Height_B_D` | Number | Detail row height | `12.75` |
| `LineWeight_B_D` | Short Text/lookup | Detail border weight | `Thin` |
| `LineTop_B_D` | Yes/No | Top border | `Yes` |
| `LineBottom_B_D` | Yes/No | Bottom border | `Yes` |
| `LineLeft_B_D` | Yes/No | Left border | `Yes` |
| `LineRight_B_D` | Yes/No | Right border | `Yes` |
| `LineHoriz_B_D` | Yes/No | Inside horizontal borders | `Yes` |
| `MergeCenter_B_D` | Short Text | Row-relative merge range using `?` | `A?:C?` |
| `LineWeight_B_Bline_D` | Short Text/lookup | Detail-region outer boundary weight | `Medium` |
| `LineTop_B_Bline_D` | Yes/No | Boundary top line | `Yes` |
| `LineBottom_B_Bline_D` | Yes/No | Boundary bottom line | `Yes` |
| `LineLeft_B_Bline_D` | Yes/No | Boundary left line | `Yes` |
| `LineRight_B_Bline_D` | Yes/No | Boundary right line | `Yes` |
| `Range_B_D` | Long Text | Excel formula applied across the detail range | `A?+B?` |
| `List_B` | Long Text | Validation-list configuration, where used | Project-specific |
| `ErrorTitle_B` | Short Text | Validation error title | `Invalid value` |
| `ErrorMessage_B` | Long Text | Validation error message | `Select a listed value.` |
| `Lock_B` | Yes/No | Metadata flag for cell protection | `No` |

### Running-total example

To calculate a cumulative sum of J while excluding rows whose I value is `Totals:`:

```text
SUMIFS($J$1:INDEX(J:J,ROW()-1),$I$1:INDEX(I:I,ROW()-1),"<>Totals:")
```

To restart the calculation with each block, use `?` as the block start:

```text
SUMIFS($J$?:INDEX(J:J,ROW()-1),$I$?:INDEX(I:I,ROW()-1),"<>Totals:")
```

A hidden helper column such as T can hold this formula and F2 can copy the final helper value into a visible footer.

## 8. F1 Primary Footer Fields

F1 fields control the primary summary row for each group.

| Field | Suggested type | Purpose | Example |
|---|---|---|---|
| `FieldName_F1` | Short Text | Static subtotal label or value | `Totals:` |
| `Add_Name_From_Data_F1` | Short Text | Copies a value from `End1` in another column | `T` |
| `Alignment_Horiz_F1` | Short Text/lookup | Footer horizontal alignment | `Right` |
| `Alignment_Vert_F1` | Short Text/lookup | Footer vertical alignment | `Center` |
| `Format_F1` | Short Text | Footer Excel number format | `0.0%` |
| `ColorIndex_F1` | Short Text/lookup | Footer background color | `Gold` |
| `Font_F1` | Short Text/lookup | Footer font | `Arial` |
| `Font_Size_F1` | Number | Footer font size | `10` |
| `Font_Style_F1` | Short Text/lookup | Footer font style | `Bold` |
| `Font_Color_F1` | Short Text/lookup | Footer font color | `Black` |
| `Wrap_Text_F1` | Yes/No | Wraps footer content | `No` |
| `Underline_F1` | Yes/No | Underlines footer content | `No` |
| `Row_Height_F1` | Number | Footer row height | `12.75` |
| `LineWeight_F1` | Short Text/lookup | Footer border weight | `Thin` |
| `LineTop_F1` | Yes/No | Top border | `Yes` |
| `LineBottom_F1` | Yes/No | Bottom border | `Yes` |
| `LineLeft_F1` | Yes/No | Left border | `Yes` |
| `LineRight_F1` | Yes/No | Right border | `Yes` |
| `LineWeight_F1_Bline` | Short Text/lookup | Footer-area boundary weight | `Medium` |
| `LineTop_F1_Bline` | Yes/No | Boundary top border | `Yes` |
| `LineBottom_F1_Bline` | Yes/No | Boundary bottom border | `No` |
| `LineLeft_F1_Bline` | Yes/No | Boundary left border | `Yes` |
| `LineRight_F1_Bline` | Yes/No | Boundary right border | `Yes` |
| `Range_F1` | Long Text | Footer formula using `?`, `@`, and `#` | `SUM(J@:J#)` |

## 9. F2 Nested Footer Fields

F2 controls the extended nested-summary panel. The implementation allocates three F2 positions numbered 0 through 2.

| Field | Suggested type | Purpose | Example |
|---|---|---|---|
| `HeaderCol_F2_A` | Yes/No | Enables paired F2 label and data output | `Yes` |
| `FieldName_F2_A` | Short Text | F2 label | `Runn Tot:` |
| `Add_Name_From_Data_F2_A` | Short Text | Source Excel column for the copied F2 value | `T` |
| `GetFromSum_F2_A` | Yes/No | `No` reads `End1`; `Yes` reads `End1 + 1` | `No` |
| `Position_F2_A` | Number, Integer | F2 panel position: 0, 1, or 2 | `0` |
| `Alignment_Horiz_F2_A` | Short Text/lookup | Label horizontal alignment | `Right` |
| `Alignment_Vert_F2_A` | Short Text/lookup | Label vertical alignment | `Top` |
| `Font_F2_A` | Short Text/lookup | Label font | `Arial` |
| `Font_Size_F2_A` | Number | Label font size | `10` |
| `Font_Style_F2_A` | Short Text/lookup | Label font style | `Bold` |
| `Font_Color_F2_A` | Short Text/lookup | Label font color | `Black` |
| `Format_F2_A` | Short Text | Label number format | `General` |
| `Wrap_Text_F2_A` | Yes/No | Wraps label text | `No` |
| `Underline_F2_A` | Yes/No | Underlines label | `No` |
| `Font_F2_AD` | Short Text/lookup | Copied-value font | `Arial` |
| `Font_Size_F2_AD` | Number | Copied-value font size | `10` |
| `Font_Style_F2_AD` | Short Text/lookup | Copied-value style | `Bold` |
| `Font_Color_F2_AD` | Short Text/lookup | Copied-value color | `Black` |
| `Format_F2_AD` | Short Text | Copied-value number format | `0` |
| `Wrap_Text_F2_AD` | Yes/No | Wraps copied value | `No` |
| `Underline_F2_AD` | Yes/No | Underlines copied value | `No` |
| `Merge_Center_F2` | Yes/No | Merges configured F2 rows across visible columns | `No` |
| `ColorIndex_F2` | Short Text/lookup | F2 panel background color | `Light Gold` |
| `Row_Height_F2` | Number | F2 row height | `8` |
| `Repeat_YN_F2` | Yes/No | Metadata flag for repeated F2 behavior | `No` |
| `LineWeight_F2` | Short Text/lookup | F2 panel border weight | `Medium` |
| `LineTop_F2` | Yes/No | Top border | `No` |
| `LineBottom_F2` | Yes/No | Bottom border | `Yes` |
| `LineLeft_F2` | Yes/No | Left border | `Yes` |
| `LineRight_F2` | Yes/No | Right border | `Yes` |

### `Position_F2_A`

- `0`: first F2 row
- `1`: second F2 row
- `2`: third F2 row

Use the Access `Number` data type with the `Integer` field size and a default value of `0`.

### `GetFromSum_F2_A`

Use an Access `Yes/No` field with a checkbox and a default of `No`.

- `No`: copy from the last detail row, `End1`.
- `Yes`: copy from `End1 + 1`, normally the generated subtotal row.

## 10. Grand-Total Metadata

`tblGoodFieldName_G_Totals` is reserved for report-level grand-total presentation and calculations after ordinary grouping blocks have been processed.

Use it for:

- Final report labels
- Final summary formulas
- Grand-total number formats
- Fonts, colors, borders, and line weights
- Values summarized across completed groups

The exact field-by-field dictionary should mirror the schema actually distributed in `ReportKPI.accdb`. The current supplied source confirms that the table is active, but the complete table-definition export was not included in the available documentation. Contributors should treat the Access table design as authoritative until its schema is exported into this guide.

## 11. Worked Configuration Example

The following pattern creates a hidden running-total helper in T and displays its final value in an F2 panel.

### Hidden helper column T

```text
Letter: T
Column_Hide: Yes
Range_B_D: SUMIFS($J$1:INDEX(J:J,ROW()-1),$I$1:INDEX(I:I,ROW()-1),"<>Totals:")
Format_B_D: 0
```

### Visible F2 output

```text
Letter: I
HeaderCol_F2_A: Yes
FieldName_F2_A: Runn Tot:
Position_F2_A: 0
Add_Name_From_Data_F2_A: T
GetFromSum_F2_A: No
Format_F2_AD: 0
```

The helper formula runs in T. The F2 configuration copies the value from T at the last detail row and places it in the nested footer associated with column I.

## 12. Troubleshooting

### Failure while setting `ColumnWidth`

Check:

- `Letter` contains a valid Excel column label.
- `ColumnWidth` is numeric or Null.
- An empty text value is not being used where the code expects Null.
- A width of zero is not being used as a substitute for `Column_Hide`.

### Formula displays as text

- Store the metadata expression without a leading equals sign.
- Confirm the VBA assigns the value through `.Formula`.
- Confirm that the destination is not preformatted as Text.

### F2 label appears but the value does not

Check:

- `HeaderCol_F2_A` is `Yes` when using paired label/data behavior.
- `Add_Name_From_Data_F2_A` contains a valid Excel column letter.
- The source value exists at `End1` or `End1 + 1`, depending on `GetFromSum_F2_A`.
- `Position_F2_A` is 0, 1, or 2.

### Running total includes subtotal rows

Use criteria that excludes the subtotal label row, for example:

```text
SUMIFS($J$1:INDEX(J:J,ROW()-1),$I$1:INDEX(I:I,ROW()-1),"<>Totals:")
```

### Formula does not restart for each block

Replace fixed starting rows with `?` in `Range_B_D`:

```text
SUMIFS($J$?:INDEX(J:J,ROW()-1),$I$?:INDEX(I:I,ROW()-1),"<>Totals:")
```

### Metadata field is missing

If VBA references a field that is absent from the table, add the field using the type expected by the code. For `GetFromSum_F2_A`, use `Yes/No` with a default of `No`.

## 13. Recommended Metadata Validation

Before rendering, validate that:

- Every record has `Report`, `Tab`, `Sort`, and `Letter`.
- `Letter` contains a valid Excel column label.
- Column widths and row heights are numeric or Null.
- Position fields use supported values.
- Yes/No settings use the Access Boolean type.
- Alignment, color, and line-weight values match values handled by the VBA.
- Formula metadata does not include a leading equals sign.
- Formula tokens are used only in fields whose code replaces them.
- Source-column fields contain valid Excel letters.

## 14. Maintenance Notes

When extending the engine:

- Add new lookup values only when corresponding VBA handling exists.
- Keep new fields beside the fields they modify.
- Use hidden helper columns for complex intermediate calculations.
- Keep chart formatting report-specific rather than forcing every chart property into metadata.
- Update this guide whenever a new metadata field, token, lookup value, or rendering behavior is introduced.
