# Metadata Configuration Guide

## Purpose

This guide explains how the active metadata tables control workbook layout, formulas, formatting, report blocks, subtotals, and grand totals in the MetadataDrivenReportingEngine.

The active tables are:

- `tblGoodFieldName_Report` for ordinary grouped report blocks.
- `tblGoodFieldName_G_Totals` for the final report-level grand-total block.

The guide is based on the current `clsExport` rendering logic, the supplied sample configuration, and the documented `tblGoodFieldName_G_Totals` schema.

## 1. Report Block Anatomy

```text
+--------------------------------------------------+
| H1: Primary group header                        |
| Static labels and values copied from data       |
+--------------------------------------------------+
| B: Body or column-heading row                   |
+--------------------------------------------------+
| B_D: Detail-data rows                           |
| Record 1                                        |
| Record 2                                        |
| ...                                             |
| Record n                                        |
+--------------------------------------------------+
| F1: Primary subtotal or summary row             |
+--------------------------------------------------+
| F2 position 0: Extended-summary row 1           |
| F2 position 1: Extended-summary row 2           |
| F2 position 2: Extended-summary row 3           |
+--------------------------------------------------+
```

| Suffix | Zone | Purpose |
|---|---|---|
| `_H1_A` | H1 label | Header label and placement |
| `_H1_AD` | H1 data | Header value copied from data |
| `_B` | B heading | Column-heading row |
| `_B_D` | B_D detail | Detail region |
| `_F1` | F1 footer | Primary subtotal row |
| `_F2_A` | F2 label | Extended-footer label and placement |
| `_F2_AD` | F2 data | Extended-footer value copied from data |
| `_F2` | F2 panel | Overall extended-footer formatting |

## 2. Formula Tokens

Store formula metadata without a leading equals sign because the VBA adds `=` when assigning the formula to Excel.

### `?` in `Range_B_D`

`?` is replaced with the current block's first detail row, `begin1`.

```text
A?+B?
```

If the block begins at row 18, Excel receives:

```text
=A18+B18
```

Excel adjusts the relative references as the formula is filled down the selected detail range.

### `?` in `Range_F1`

`?` is replaced with `End1 + 1`, the row immediately following the final detail row.

```text
J?+K?
```

If `End1` is 35, Excel receives:

```text
=J36+K36
```

### `?` in `MergeCenter_B_D`

`?` is replaced with each current detail-row number while the engine loops through the block.

```text
A?:C?
```

This creates ranges such as `A18:C18`, `A19:C19`, and `A20:C20`.

### `@` and `#` in `Range_F1`

- `@` represents the first detail row, `begin1`.
- `#` represents the last detail row, `End1`.

```text
SUM(J@:J#)
```

For a block running from row 18 through row 35, Excel receives:

```text
=SUM(J18:J35)
```

## 3. Common Accepted Values

### Yes/No

Use the Access `Yes/No` data type with a checkbox. Do not store the words `Yes` and `No` in a text field.

### Horizontal alignment

- `Left`
- `Center`
- `Right`

### Vertical alignment

- `Top`
- `Center`
- `Bottom`

### Line weight

- `Thin`
- `Medium`
- `Thick`

### Fonts and styles

Documented font choices include:

- `MS Sans Serif`
- `Arial`

Documented styles include:

- `Regular`
- `Bold`
- `Italic`
- `Bold Italic`

### Colors

Color lists vary slightly by zone. Values used by the current metadata include:

- `None`, where supported
- `Black`
- `Red`
- `Green`
- `Blue`
- `Gold`
- `Light Gold`
- `Light Brown`

### Excel number formats

Format fields use the lookup query:

```text
SELECT tblFormat.Format, tblFormat.Description FROM tblFormat;
```

Examples include `General`, `0`, `0.0%`, and `m/d/yyyy`.

## 4. Core `tblGoodFieldName_Report` Fields

Each record generally represents one output column.

| Field | Recommended Access type | Purpose | Example |
|---|---|---|---|
| `Sort` | Long Integer | Processing order | `10` |
| `Report` | Short Text | Report identifier | `Monthly_Report.XLSX` |
| `Tab` | Short Text | Worksheet identifier | `KPI_Monthly` |
| `Letter` | Short Text | Excel destination column | `J` |
| `Column_Hide` | Yes/No | Hides a worksheet column | `No` |
| `ColumnWidth` | Decimal/Double | Excel column width | `17.14` |

`Letter` must contain a valid Excel column label. Use `Column_Hide` rather than a width of zero when a helper column should be hidden.

## 5. H1 Primary Header Fields

| Field | Type | Purpose | Example |
|---|---|---|---|
| `HeaderCol_H1_A` | Yes/No | Enables paired header label/data behavior | `Yes` |
| `FieldName_H1_A` | Short Text | Static header label | `Monthly KPI:` |
| `Add_Name_From_Data_H1_A` | Short Text | Source Excel column for the copied value | `Q` |
| `Position_H1_A` | Long Integer | Vertical placement within H1 | `0` |
| `Alignment_Horiz_H1_A` | Lookup text | Label alignment | `Center` |
| `Alignment_Vert_H1_A` | Lookup text | Label vertical alignment | `Top` |
| `Font_H1_A` | Lookup text | Label font | `Arial` |
| `Font_Size_H1_A` | Number | Label font size | `12` |
| `Font_Style_H1_A` | Lookup text | Label style | `Regular` |
| `Font_Color_H1_A` | Lookup text | Label color | `Black` |
| `Format_H1_A` | Lookup text | Label number format | `General` |
| `Wrap_Text_H1_A` | Yes/No | Wraps label text | `Yes` |
| `Underline_H1_A` | Yes/No | Underlines label | `No` |
| `Font_H1_AD` | Lookup text | Copied-value font | `Arial` |
| `Font_Size_H1_AD` | Number | Copied-value size | `14` |
| `Font_Style_H1_AD` | Lookup text | Copied-value style | `Bold` |
| `Font_Color_H1_AD` | Lookup text | Copied-value color | `Black` |
| `Format_H1_AD` | Lookup text | Copied-value format | `General` |
| `Wrap_Text_H1_AD` | Yes/No | Wraps copied value | `No` |
| `Underline_H1_AD` | Yes/No | Underlines copied value | `No` |
| `Merge_Center_H1` | Yes/No | Merges configured H1 cells | `Yes` |
| `ColorIndex_H1` | Lookup text | H1 fill color | `Light Brown` |
| `Row_Height_H1` | Decimal | H1 row height | `24` |
| `Repeat_YN_H1` | Yes/No | Repeated-header metadata flag | `No` |
| `LineWeight_H1` | Lookup text | Border weight | `Medium` |
| `LineTop_H1` | Yes/No | Top border | `Yes` |
| `LineBottom_H1` | Yes/No | Bottom border | `Yes` |
| `LineLeft_H1` | Yes/No | Left border | `No` |
| `LineRight_H1` | Yes/No | Right border | `Yes` |
| `LineHoriz_H1` | Yes/No | Inside horizontal borders | `No` |
| `LineVert_H1` | Yes/No | Inside vertical borders | `No` |

## 6. B Body-Heading Fields

| Field | Type | Purpose | Example |
|---|---|---|---|
| `FieldName_B` | Short Text 255 | Displayed heading | `Count Of Late WO` |
| `Add_Name_From_Data_B` | Short Text 50 | Alternate source for heading value | Blank |
| `Alignment_Horiz_B` | Lookup text | Horizontal alignment | `Center` |
| `Alignment_Vert_B` | Lookup text | Vertical alignment | `Center` |
| `Format_B` | Lookup text | Number format | `General` |
| `ColorIndex_B` | Lookup text | Fill color | `None` |
| `Font_B` | Lookup text | Font | `Arial` |
| `Font_Size_B` | Long Integer | Font size | `10` |
| `Font_Style_B` | Lookup text | Font style | `Bold Italic` |
| `Font_Color_B` | Lookup text | Font color | `Black` |
| `Wrap_Text_B` | Yes/No | Wraps heading | `No` |
| `Underline_B` | Yes/No | Underlines heading | `Yes` |
| `Row_Height_B` | Decimal | Heading row height | `17.25` |
| `LineWeight_B` | Lookup text | Border weight | `Thin` |
| `LineTop_B` | Yes/No | Top border | `No` |
| `LineBottom_B` | Yes/No | Bottom border | `Yes` |
| `LineLeft_B` | Yes/No | Left border | `Yes` |
| `LineRight_B` | Yes/No | Right border | `Yes` |
| `LineWeight_B_Bline` | Lookup text | Boundary weight | `Medium` |
| `LineTop_B_Bline` | Yes/No | Boundary top line | `No` |
| `LineBottom_B_Bline` | Yes/No | Boundary bottom line | `No` |
| `LineLeft_B_Bline` | Yes/No | Boundary left line | `Yes` |
| `LineRight_B_Bline` | Yes/No | Boundary right line | `Yes` |

## 7. B_D Detail Fields

| Field | Type | Purpose | Example |
|---|---|---|---|
| `Alignment_Horiz_B_D` | Lookup text | Detail alignment | `Center` |
| `Alignment_Vert_B_D` | Lookup text | Detail vertical alignment | `Center` |
| `Format_B_D` | Lookup text | Excel number format | `0.0%` |
| `ColorIndex_B_D` | Lookup text | Detail fill color | `None` |
| `Font_B_D` | Lookup text | Detail font | `Arial` |
| `Font_Size_B_D` | Long Integer | Detail font size | `10` |
| `Font_Style_B_D` | Lookup text | Detail style | `Regular` |
| `Font_Color_B_D` | Lookup text | Detail font color | `Black` |
| `Wrap_Text_B_D` | Yes/No | Wraps values | `No` |
| `Underline_B_D` | Yes/No | Underlines values | `No` |
| `Row_Height_B_D` | Decimal | Detail row height | `12.75` |
| `LineWeight_B_D` | Lookup text | Detail border weight | `Thin` |
| `LineTop_B_D` | Yes/No | Top border | `Yes` |
| `LineBottom_B_D` | Yes/No | Bottom border | `Yes` |
| `LineLeft_B_D` | Yes/No | Left border | `Yes` |
| `LineRight_B_D` | Yes/No | Right border | `Yes` |
| `LineHoriz_B_D` | Yes/No | Inside horizontal borders | `Yes` |
| `MergeCenter_B_D` | Short Text 50 | Row-relative merge range | `A?:C?` |
| `LineWeight_B_Bline_D` | Lookup text | Outer-boundary weight | `Medium` |
| `LineTop_B_Bline_D` | Yes/No | Boundary top line | `Yes` |
| `LineBottom_B_Bline_D` | Yes/No | Boundary bottom line | `Yes` |
| `LineLeft_B_Bline_D` | Yes/No | Boundary left line | `Yes` |
| `LineRight_B_Bline_D` | Yes/No | Boundary right line | `Yes` |
| `Range_B_D` | Short Text 50 | Detail formula | `A?+B?` |
| `List_B` | Short Text 255 | Validation-list configuration | Project-specific |
| `ErrorTitle_B` | Short Text 50 | Validation error title | `Invalid value` |
| `ErrorMessage_B` | Short Text 255 | Validation error message | `Select a listed value.` |
| `Lock_B` | Yes/No | Cell-protection flag | `No` |

### Running-total example

```text
SUMIFS($J$1:INDEX(J:J,ROW()-1),$I$1:INDEX(I:I,ROW()-1),"<>Totals:")
```

To restart at each block's first detail row:

```text
SUMIFS($J$?:INDEX(J:J,ROW()-1),$I$?:INDEX(I:I,ROW()-1),"<>Totals:")
```

## 8. F1 Primary Footer Fields

| Field | Type | Purpose | Example |
|---|---|---|---|
| `FieldName_F1` | Short Text 255 | Static subtotal label | `Totals:` |
| `Add_Name_From_Data_F1` | Short Text 50 | Source column at `End1` | `T` |
| `Alignment_Horiz_F1` | Lookup text | Horizontal alignment | `Right` |
| `Alignment_Vert_F1` | Lookup text | Vertical alignment | `Center` |
| `Format_F1` | Lookup text | Number format | `0.0%` |
| `ColorIndex_F1` | Lookup text | Fill color | `Gold` |
| `Font_F1` | Lookup text | Font | `Arial` |
| `Font_Size_F1` | Long Integer | Font size | `10` |
| `Font_Style_F1` | Lookup text | Font style | `Bold` |
| `Font_Color_F1` | Lookup text | Font color | `Black` |
| `Wrap_Text_F1` | Yes/No | Wraps footer value | `No` |
| `Underline_F1` | Yes/No | Underlines footer value | `No` |
| `Row_Height_F1` | Decimal | Footer row height | `12.75` |
| `LineWeight_F1` | Lookup text | Border weight | `Thin` |
| `LineTop_F1` | Yes/No | Top border | `Yes` |
| `LineBottom_F1` | Yes/No | Bottom border | `Yes` |
| `LineLeft_F1` | Yes/No | Left border | `Yes` |
| `LineRight_F1` | Yes/No | Right border | `Yes` |
| `LineWeight_F1_Bline` | Lookup text | Footer boundary weight | `Medium` |
| `LineTop_F1_Bline` | Yes/No | Boundary top | `Yes` |
| `LineBottom_F1_Bline` | Yes/No | Boundary bottom | `No` |
| `LineLeft_F1_Bline` | Yes/No | Boundary left | `Yes` |
| `LineRight_F1_Bline` | Yes/No | Boundary right | `Yes` |
| `Range_F1` | Short Text 50 | Footer formula | `SUM(J@:J#)` |

## 9. F2 Extended Footer Fields

| Field | Type | Purpose | Example |
|---|---|---|---|
| `HeaderCol_F2_A` | Yes/No | Enables paired label/data behavior | `Yes` |
| `FieldName_F2_A` | Short Text 255 | F2 label | `Runn Tot:` |
| `Add_Name_From_Data_F2_A` | Short Text 50 | Source Excel column | `T` |
| `GetFromSum_F2_A` | Yes/No | No reads `End1`; Yes reads `End1 + 1` | `No` |
| `Position_F2_A` | Long Integer | F2 row position, 0 through 2 | `0` |
| `Alignment_Horiz_F2_A` | Lookup text | Label alignment | `Right` |
| `Alignment_Vert_F2_A` | Lookup text | Label vertical alignment | `Top` |
| `Font_F2_A` | Short Text 75 | Label font | `Arial` |
| `Font_Size_F2_A` | Decimal | Label font size | `10` |
| `Font_Style_F2_A` | Lookup text | Label style | `Bold` |
| `Font_Color_F2_A` | Lookup text | Label color | `Black` |
| `Format_F2_A` | Lookup text | Label format | `General` |
| `Wrap_Text_F2_A` | Yes/No | Wraps label | `No` |
| `Underline_F2_A` | Yes/No | Underlines label | `No` |
| `Font_F2_AD` | Short Text 75 | Copied-value font | `Arial` |
| `Font_Size_F2_AD` | Decimal | Copied-value size | `10` |
| `Font_Style_F2_AD` | Lookup text | Copied-value style | `Bold` |
| `Font_Color_F2_AD` | Lookup text | Copied-value color | `Black` |
| `Format_F2_AD` | Lookup text | Copied-value format | `0` |
| `Wrap_Text_F2_AD` | Yes/No | Wraps copied value | `No` |
| `Underline_F2_AD` | Yes/No | Underlines copied value | `No` |
| `Merge_Center_F2` | Yes/No | Merges configured F2 rows | `No` |
| `ColorIndex_F2` | Lookup text | Panel fill color | `Light Gold` |
| `Row_Height_F2` | Decimal | F2 row height | `8` |
| `Repeat_YN_F2` | Yes/No | Repeated-F2 metadata flag | `No` |
| `LineWeight_F2` | Lookup text | Panel border weight | `Medium` |
| `LineTop_F2` | Yes/No | Top border | `No` |
| `LineBottom_F2` | Yes/No | Bottom border | `Yes` |
| `LineLeft_F2` | Yes/No | Left border | `Yes` |
| `LineRight_F2` | Yes/No | Right border | `Yes` |

### F2 position values

- `0`: first F2 row
- `1`: second F2 row
- `2`: third F2 row

## 10. `tblGoodFieldName_G_Totals`

The grand-total metadata table contains the same core H1, B, B_D, F1, and F2 field groups described above. This lets the final grand-total block use separate labels, formulas, colors, borders, and formatting from ordinary report blocks.

When grand totals are enabled, the engine creates a final synthetic grouping record labeled `Grand Totals`, then renders that block with grand-total metadata.

### Additional workbook-control fields

| Field | Access type | Purpose | Example |
|---|---|---|---|
| `ID` | AutoNumber | Primary key | Automatically assigned |
| `PasteSpecial` | Yes/No | Grand-total workbook-control flag | `No` |
| `PasswordProtect` | Yes/No | Enables protection behavior | `No` |
| `Password` | Short Text 50 | Password when protection is enabled | Blank |

### Exact core field types

| Field | Access type |
|---|---|
| `Sort` | Long Integer |
| `Report` | Short Text 255 |
| `Tab` | Short Text 255 |
| `Letter` | Short Text 2 |
| `Column_Hide` | Yes/No |
| `ColumnWidth` | Decimal, precision 4, scale 2 |

The documented default for `ColumnWidth` is `0`. A positive width is safer for visible columns; use `Column_Hide` for intentionally hidden columns.

### Exact specialized text fields

| Field | Access type |
|---|---|
| `FieldName_H1_A`, `FieldName_B`, `FieldName_F1`, `FieldName_F2_A` | Short Text 255 |
| `Add_Name_From_Data_H1_A`, `Add_Name_From_Data_B`, `Add_Name_From_Data_F1`, `Add_Name_From_Data_F2_A` | Short Text 50 |
| `Range_B_D`, `Range_F1`, `MergeCenter_B_D` | Short Text 50 |
| `List_B`, `ErrorMessage_B` | Short Text 255 |
| `ErrorTitle_B` | Short Text 50 |
| Font-name fields | Short Text 75 |
| Alignment, style, color, format, and line-weight fields | Short Text 50 |

### Numeric field patterns

| Category | Access type |
|---|---|
| H1 and F2 font sizes | Decimal, one decimal place |
| B, B_D, and F1 font sizes | Long Integer |
| Row heights | Decimal, two decimal places |
| H1 and F2 positions | Long Integer |

### Lookup sources

| Category | Values or source |
|---|---|
| Report | `SELECT tblReports.[Report Name] FROM tblReports ORDER BY tblReports.[Report Name];` |
| Horizontal alignment | `Left;Center;Right` |
| Vertical alignment | `Top;Center;Bottom` |
| Fonts | `MS Sans Serif;Arial` |
| Font styles | `Regular;Bold;Italic;Bold Italic` |
| Line weights | `Thin;Medium;Thick` |
| Formats | `SELECT tblFormat.Format, tblFormat.Description FROM tblFormat;` |

### Indexes

| Index | Fields | Properties |
|---|---|---|
| `PrimaryKey` | `ID` | Primary, unique, required |
| `letter` | `Letter` | Non-unique |
| `report` | `Report`, `Tab` | Non-unique compound index |
| `sort` | `Sort` | Non-unique |
| `tab` | `Tab` | Non-unique |

## 11. Worked Example: Hidden Running Total

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

The helper formula runs in T. F2 copies the T value from the last detail row into the extended footer associated with column I.

## 12. Troubleshooting

### Failure while setting `ColumnWidth`

Check that:

- `Letter` contains a valid Excel column.
- `ColumnWidth` is numeric or Null.
- An empty string is not being used where the code expects Null.
- Zero is not being used instead of `Column_Hide`.

### Formula displays as text

- Store the expression without a leading equals sign.
- Confirm the destination is not formatted as Text.
- Confirm the code assigns the expression through `.Formula`.

### F2 label appears but its value does not

Check that:

- `HeaderCol_F2_A` is `Yes` for paired label/data output.
- `Add_Name_From_Data_F2_A` contains a valid source column.
- The source value exists at `End1` or `End1 + 1`.
- `Position_F2_A` is `0`, `1`, or `2`.

### Running total includes subtotal rows

Use criteria excluding the subtotal label:

```text
SUMIFS($J$1:INDEX(J:J,ROW()-1),$I$1:INDEX(I:I,ROW()-1),"<>Totals:")
```

### Formula does not restart for each block

Use `?` for the block's first detail row:

```text
SUMIFS($J$?:INDEX(J:J,ROW()-1),$I$?:INDEX(I:I,ROW()-1),"<>Totals:")
```

## 13. Validation Checklist

Before rendering, validate that:

- Every metadata record has a report, tab, sort order, and valid Excel letter.
- Widths, row heights, font sizes, and positions are numeric.
- F2 positions are limited to 0, 1, and 2.
- Yes/No settings use Access Boolean fields.
- Lookup values match VBA `Select Case` values.
- Formula metadata has no leading equals sign.
- Formula tokens are used only in fields whose code replaces them.
- Source-column fields contain valid Excel letters.
- Hidden helper columns use `Column_Hide`.

## 14. Maintenance Notes

- Add new lookup values only when corresponding VBA handling exists.
- Keep new fields beside the fields they modify.
- Use hidden helper columns for complex intermediate calculations.
- Keep chart formatting report-specific rather than forcing all chart properties into metadata.
- Update this guide whenever fields, tokens, lookups, or rendering behavior change.
