# Getting Started

[Back to project README](../README.md) | [Architectural White Paper](WhitePaper.md) | [Report Header Reference](ReportHeaderReference.md) | [Metadata Configuration Guide](MetadataConfigurationGuide.md) | [`tblReportSequence` Reference](TblReportSequenceReference.md) | [Troubleshooting](Troubleshooting.md)

## Purpose

This guide explains how to configure a new worksheet for the Metadata-Driven Excel Reporting Engine. It focuses on the proven materialized-table pattern used by the Medicare and KPI demonstrations.

The goal is to move from a tested Access record source to a generated Excel worksheet with dynamic report bands, metadata-controlled formatting, formulas, totals, and optional charts.

## 1. Understand the Required Artifacts

A basic worksheet configuration requires:

1. A report-ready Access source
2. A final worksheet source
3. A sorting/grouping table
4. Four ordered rows in `tblReportSequence`
5. Optional worksheet-level header metadata in `tblGoodFieldName_Header`
6. Repeating report-band metadata in `tblGoodFieldName_Report`
7. Optional grand-total metadata in `tblGoodFieldName_G_Totals`
8. A report registration row in `tblReports`
9. Launcher and renderer modules

The standard flow is:

```text
Prepared source
    -> final worksheet table
    -> sorting table
    -> tblReportSequence
    -> ADO recordset transfer
    -> one-time worksheet header
    -> break calculation
    -> repeating report-band metadata
    -> grand totals and charts
    -> standalone XLSX workbook
```

## 2. Start from a Working Demo

The safest development method is to copy a working report definition and change it incrementally.

Use the Medicare demo when the new report resembles:

- One large detail worksheet
- One primary grouping level
- Report-level totals
- A current renderer baseline
- Explicit workbook recalculation before save

Use the KPI demo when the new report resembles:

- Multiple worksheets
- Multiple sorting tables
- Multiple chart sheets
- Hidden helper columns
- Running footer calculations
- Broader workbook orchestration

Do not begin by creating every table and metadata row from scratch unless necessary.

## 3. Prepare the Report-Ready Source

The host application should prepare a source containing the records and fields required by Excel.

The source may begin as:

- A local Access table
- A linked table
- A saved select query
- A sequence of Access action queries
- A result imported from an external source

The final report-ready source should already contain fields required for:

- Filtering
- Grouping
- Sorting
- Display labels
- Business calculations shared by multiple outputs

Presentation-only formulas can be added later through metadata-generated Excel columns.

### 3.1 Choose the final field order

The final worksheet source should expose fields in the order expected by the metadata.

Create a simple field map before continuing:

```text
Access field                         Excel column
State                                A
ProviderType                         B
PlaceOfServiceDescription            C
RBCSCategory                         E
RBCSSubcategory                      F
ServiceCount                         J
EstimatedSubmittedCharges            K
EstimatedAllowedAmount               L
EstimatedMedicarePayment             M
```

This map becomes the basis for:

- Metadata `Letter` values
- Header and footer source-column references
- Formula templates
- Hidden helper columns
- Chart source ranges

### 3.2 Add explicit sort fields

Do not rely on user-facing text alone when a stable sort key is needed.

Examples include:

- Numeric place-of-service sort
- Uppercase trimmed provider-type sort
- Group-band key
- Month sequence
- Fiscal-year sequence

A reliable report source should make ordering explicit.

## 4. Materialize the Final Worksheet Table

The demonstrated pattern uses a make-table query stored in `tblReportSequence`.

Example:

```sql
SELECT
    Q.State,
    Q.ProviderType,
    Q.PlaceOfServiceDescription,
    Q.PlaceOfServiceSort,
    Q.RBCSCategory,
    Q.RBCSSubcategory,
    Q.RBCSBandKey,
    Q.RBCSBandLabel,
    Q.ProviderTypeSort,
    Q.ServiceCount,
    Q.EstimatedSubmittedCharges,
    Q.EstimatedAllowedAmount,
    Q.EstimatedMedicarePayment
INTO Medicare_ReportFinal
FROM qryMedicare_ReportFinal AS Q
ORDER BY
    Q.RBCSBandKey,
    Q.EstimatedMedicarePayment DESC,
    Q.ProviderTypeSort,
    Q.PlaceOfServiceSort;
```

The `ORDER BY` sequence matters because the sorting table's group counts must align with the physical record order copied into Excel.

### 4.1 Test the SQL independently

Before storing SQL in `tblReportSequence`:

1. Create a normal Access query.
2. Run it manually.
3. Verify the destination table is created.
4. Verify field names and order.
5. Verify record order.
6. Verify there are no unresolved parameters.
7. Delete the test destination table.

Only then copy the SQL into the sequence table.

## 5. Create the Sorting Table

The sorting table summarizes the final worksheet source by report band.

Its first fields must provide:

1. Ordered group sequence or sortable key
2. Group description
3. Record count

A Medicare-style example is:

```sql
SELECT
    1 + (
        SELECT Count(*)
        FROM (
            SELECT X.RBCSBandKey
            FROM Medicare_ReportFinal AS X
            GROUP BY X.RBCSBandKey
        ) AS PriorBands
        WHERE PriorBands.RBCSBandKey < M.RBCSBandKey
    ) AS BandSequence,
    M.RBCSBandKey,
    Count(*) AS RecCount
INTO tblSort_MedicareFinal
FROM Medicare_ReportFinal AS M
GROUP BY M.RBCSBandKey
ORDER BY M.RBCSBandKey;
```

A KPI-style example is:

```sql
SELECT
    [Year] & Left([YearNumber],2) AS SpecialSort,
    [MonthTxt] & " " & [Year] AS Month_Year,
    Count(KPI_Results_Monthly.WorkOrdr) AS CountOfWorkOrdr
INTO tblSorKPI_MonthlyReport
FROM KPI_Results_Monthly
GROUP BY
    [Year] & Left([YearNumber],2),
    [MonthTxt] & " " & [Year]
ORDER BY
    [Year] & Left([YearNumber],2);
```

### 5.1 Reconcile record counts

Before continuing, verify:

```text
Sum of sorting-table record counts
    = number of rows in final worksheet source
```

If the counts do not reconcile, the break coordinates will not align with the copied records.

### 5.2 Preserve required leading fields

`fn_FindBreaks` depends on the expected first fields and headers. Do not insert unrelated fields before the grouping key, description, and count fields.

## 6. Configure `tblReportSequence`

Use the standard four-row pattern.

### Row 1: final-output table marker

```text
Report:       Your_Report.XLSX
Sort:         1
QueryName:    tblFinalOutput
FinalOutput:  True
Tab:          ReportSheet
Tab Order:    1
isTable?:     True
Exec:         False
```

This row deletes the previous final-output table and identifies the rebuilt table as the worksheet source.

### Row 2: final-output build QueryDef

```text
Report:       Your_Report.XLSX
Sort:         2
QueryName:    qryBuildFinalOutput
isTable?:     False
SQL:          SELECT ... INTO tblFinalOutput ...;
Exec:         True
```

This row creates and executes the QueryDef that materializes the final worksheet table.

### Row 3: sorting-table marker

```text
Report:         Your_Report.XLSX
Sort:           3
QueryName:      tblSortFinalOutput
isTable?:       True
Spacing:        ReportSheet
Spacing Order:  1
HasTotals:      True or False
```

This row deletes the previous sorting table and associates the sorting table with the worksheet.

### Row 4: sorting-table build QueryDef

```text
Report:       Your_Report.XLSX
Sort:         4
QueryName:    qryBuildSortFinalOutput
isTable?:     False
SQL:          SELECT ... INTO tblSortFinalOutput ...;
Exec:         True
```

This row creates and executes the QueryDef that materializes the sorting table.

### 6.1 Leave overrides blank initially

Leave `Over_ride_head1` and `Over_ride_foot1` null until the worksheet works with default allocations.

Add overrides only when the metadata requires a different number of primary header or footer rows.

## 7. Register the Report in `tblReports`

Add a row that makes the report selectable in `frmReportLauncher`.

The exact schema may vary by demo, but the report name must match the `Report` value used in `tblReportSequence`.

Example:

```text
Report Name: Your_Report.XLSX
Type:        Your_Report
```

Open the launcher and confirm the report appears in the dropdown before testing the workbook.

## 8. Create Basic Worksheet Metadata

Add one row to `tblGoodFieldName_Report` for each Excel column required by the worksheet.

Begin with only the core fields:

- `Sort`
- `Report`
- `Tab`
- `Letter`
- `ColumnWidth`
- `Column_Hide`
- `FieldName_B`
- Basic B and B_D formatting

Example conceptual rows:

```text
Sort  Report            Tab          Letter  FieldName_B       Hidden
1     Your_Report.XLSX  ReportSheet  A       State             No
2     Your_Report.XLSX  ReportSheet  B       Provider Type     No
3     Your_Report.XLSX  ReportSheet  C       Place of Service  No
4     Your_Report.XLSX  ReportSheet  D       Sort Helper       Yes
```

### 8.1 Use the first metadata row carefully

The first row in `Sort` order is often authoritative for section-wide settings such as:

- H1 merge behavior
- H1 fill and borders
- F2 fill and borders
- Outer heading/detail/footer boundaries

Use a visible worksheet column as the first sorted metadata row whenever possible.

### 8.2 Render the plain worksheet first

Before adding headers, formulas, and totals, verify that:

- Every Access field lands in the intended Excel column.
- Column headings are correct.
- Number and date formats are correct.
- Hidden fields are hidden.
- Group boundaries align with the sorting table.

### 8.3 Configure the worksheet-level header

If the worksheet requires a one-time report title or introductory area, copy a working `tblGoodFieldName_Header` configuration.

The worksheet-level header can contain:

- A report title and subtitle
- Static explanatory text
- Excel formulas
- Public-source hyperlinks
- Generated-by information
- A generated date and time
- Merged ranges
- Fonts, fills, row heights, and borders
- Freeze-below behavior

The worksheet-level header is different from the repeating H1 report-band header:

```text
tblGoodFieldName_Header
    -> rendered once at the top of the worksheet

tblGoodFieldName_Report H1
    -> rendered once for every report band
```

To configure a new worksheet header:

1. Copy a working `tblGoodFieldName_Header` configuration.
2. Change `Report` to the new report identifier.
3. Change `Tab` to the target worksheet name.
4. Replace the title, subtitle, formulas, hyperlinks, and generated-report information.
5. Verify the configured header row numbers.
6. Verify every merged range.
7. Confirm that the report-band starting row is below the final worksheet-header row.
8. Render the worksheet and validate the header before configuring the repeating H1 zone.

The Medicare reference implementation requires report-band construction to begin at row 10 or later:

```vb
Const original = 2
Const beginSheet = 10
```

`original` identifies the original starting position of the exported records before report-band rows are inserted.

`beginSheet` identifies where the first generated report band begins. If `beginSheet` is too low, report-band construction can occupy the rows reserved for `tblGoodFieldName_Header`.

Maintain this rule:

```text
beginSheet > final worksheet-header row
```
For complete field definitions and Medicare examples, see the [Report Header Reference](ReportHeaderReference.md).
  
## 9. Configure the Repeating H1 Band Header

Add repeating H1 content only after the plain worksheet and optional worksheet-level header are stable. The H1 zone is generated once for every report band and is controlled by `tblGoodFieldName_Report`.

For each H1 label/value item, configure:

- `HeaderCol_H1_A = True`
- `FieldName_H1_A`
- `Add_Name_From_Data_H1_A`, when copying a source value
- `Position_H1_A`
- H1 label formatting
- H1 copied-value formatting

Place section-wide H1 settings on the first sorted metadata row:

- `Merge_Center_H1`
- `ColorIndex_H1`
- `Row_Height_H1`
- H1 border settings

Verify that every `Position_H1_A` fits inside the allocated H1 rows.

## 10. Configure Detail Formulas and Helper Columns

Presentation-specific calculations can be added with `Range_B_D`.

Example:

```text
IFERROR(M?/L?,0)
```

Store formulas without the leading equals sign.

To add a hidden helper column:

1. Add a metadata row with an unused Excel `Letter`.
2. Leave the source field heading blank if appropriate.
3. Put the formula template in `Range_B_D`.
4. Set the desired number format.
5. Set `Column_Hide = True`.
6. Reference the helper result from F1, F2, grand totals, or chart code.

Keep filtering, grouping, and reusable business calculations in SQL. Use helper columns for workbook-specific presentation logic.

## 11. Configure F1 Subtotals

For each subtotal column, use:

- `FieldName_F1` for a static label such as `Totals:`
- `Add_Name_From_Data_F1` to copy a value from the report band
- `Range_F1` for a formula template

Formula tokens include:

- `@` for the first detail row
- `#` for the final detail row
- `?` for the footer context used by that routine

Example:

```text
SUM(K@:K#)
```

The renderer converts the formula into the physical detail range for each band.

## 12. Configure F2 Extended Footers

Use F2 for additional positioned metrics such as:

- Running record counts
- Running late counts
- Additional percentages
- Values retrieved from hidden helper columns

Configure:

- `HeaderCol_F2_A = True`
- `FieldName_F2_A`
- `Add_Name_From_Data_F2_A`
- `GetFromSum_F2_A`
- `Position_F2_A`
- F2 label and copied-value formatting

Place shared F2 merge, fill, row-height, and border settings on the first sorted metadata row.

## 13. Configure Grand Totals

If the sorting-table marker has `HasTotals = True`, create matching rows in `tblGoodFieldName_G_Totals` for the same report and worksheet.

Verify:

- Grand-total labels use the correct columns.
- Formulas reference the required final ranges.
- Helper formulas are calculated before the grand totals are evaluated.
- Number formats match the detail and subtotal sections.

Grand totals are separate from ordinary F1 and F2 rows because they summarize the completed worksheet.

## 14. Run the Report

1. Open `frmReportManager`.
2. Verify the prepared source data.
3. Open `frmReportLauncher`.
4. Select the report definition.
5. Leave **Render Sheet?** unchecked for a normal background run, or select it when observing the render for debugging.
6. Select **Create Spreadsheets**.
7. Choose an output filename and folder.
8. Wait for the launcher to display the workbook hyperlink.
9. Open the generated workbook.

The exporter creates worksheet tabs, opens each configured final source through ADO, and copies the returned records beginning at cell `A2`. It then passes the workbook to the formatter for break calculation and metadata rendering.

## 15. Validate the Generated Workbook

Review the workbook in this order.

### 15.1 Worksheet structure

- Correct worksheet names
- Correct worksheet order
- Correct raw field-to-column alignment
- Correct hidden columns
- Worksheet-level title and subtitle appear once
- Header formulas and hyperlinks work
- Header merged ranges do not overlap
- The first generated report band begins below the worksheet header
- Freeze-below behavior uses the intended row

### 15.2 Report bands

- Correct number of groups
- Correct detail-row counts
- Correct group descriptions
- Correct H1, B, F1, and F2 placement

### 15.3 Formulas

- Detail formulas filled through every band
- `@`, `#`, and `?` replaced correctly
- No `#REF!`, `#VALUE!`, or divide-by-zero errors
- Running values exclude subtotal rows where intended

### 15.4 Totals and charts

- Band subtotals reconcile to detail rows
- Grand totals reconcile to band totals
- Chart labels and values use the completed worksheet ranges

### 15.5 Saved values

Close and reopen the workbook. Verify that formula results and charts remain correct after the workbook is reopened.

## 16. Use the 99-Row Testing Strategy

Before rendering a large production-size source, test with approximately 99 representative rows.

The small source should include:

- Several report groups
- Different group sizes
- Zero values
- Nonzero values
- Rows that exercise formula conditions
- Enough records to produce several inserted headers and footers

A 99-row test is large enough to exercise the report engine while remaining fast enough for repeated inspection.

Recommended sequence:

```text
Test SQL
    -> test sorting counts
    -> render 99 rows
    -> inspect formulas and bands
    -> correct metadata
    -> repeat
    -> restore full source
    -> run volume test
```

## 17. Troubleshooting Quick Reference

### Worksheet-level header does not appear

Confirm matching `tblGoodFieldName_Header` records exist for the selected `Report` and `Tab`.

For the Medicare reference implementation, also confirm that `fn_FindBreaks` uses:

```vb
Const original = 2
Const beginSheet = 10
```

### Destination table already exists

The table-marker row is missing, uses the wrong name, or appears after the make-table query.

### QueryDef exists but table is missing

The action-query row probably has `Exec = False`.

### ADO cannot open the final source

The final table or saved query was not created, is misspelled, or has unresolved parameters.

### Report bands do not align with records

The sorting-table counts or sort order do not match the final worksheet source.

### H1 or F2 shared formatting has no effect

The section-wide value may be on a later metadata row instead of the first sorted row.

### Formula begins with `==`

The metadata formula includes a leading equals sign. Remove it.

### Footer content overlaps

The configured position exceeds the available header or footer allocation.

### Workbook opens with stale formula values

Ensure the current renderer restores automatic calculation and performs a full calculation before saving.

## 18. Completion Checklist

- [ ] Final source SQL runs successfully in Access.
- [ ] Final source has the intended field order.
- [ ] Final source has an explicit stable sort order.
- [ ] Sorting-table counts reconcile to final-source rows.
- [ ] Sorting table satisfies the three-field contract.
- [ ] Four sequence rows are in the correct order.
- [ ] `Tab` and `Spacing` match exactly.
- [ ] Report is registered in `tblReports`.
- [ ] Basic metadata rows align with source fields.
- [ ] First sorted metadata row contains shared section settings.
- [ ] Detail formulas omit the leading equals sign.
- [ ] H1 and F2 positions fit inside their allocations.
- [ ] Grand-total metadata exists when `HasTotals = True`.
- [ ] The 99-row test renders correctly.
- [ ] The full workbook recalculates and saves correctly.
- [ ] Worksheet-header metadata exists when a one-time header is required.
- [ ] Worksheet-header `Report` and `Tab` values match the report configuration.
- [ ] Header row numbers and merge ranges are valid.
- [ ] Header formulas and hyperlinks render correctly.
- [ ] `beginSheet` is at least 10 for the Medicare reference implementation.
- [ ] `beginSheet` is greater than the final worksheet-header row.
- [ ] The first report band begins below the worksheet header.


## 19. Next Steps

After the first worksheet works:

1. Add presentation-specific helper columns.
2. Add grand totals.
3. Add chart routines or chart-ready helper ranges.
4. Add a second worksheet by repeating the four-row sequence pattern with a new `Tab Order`.
5. Document the completed pipeline as a worked example.

## 20. Related Documentation

- [Project README](../README.md)
- [Architectural White Paper](WhitePaper.md)
- [Report Header Reference](ReportHeaderReference.md)
- [Metadata Configuration Guide](MetadataConfigurationGuide.md)
- [`tblReportSequence` Reference](TblReportSequenceReference.md)
- [Troubleshooting Guide](Troubleshooting.md)
- [Medicare Pipeline](WorkedExamples/MedicarePipeline.md)
- [KPI Pipeline](WorkedExamples/KPIPipeline.md)
