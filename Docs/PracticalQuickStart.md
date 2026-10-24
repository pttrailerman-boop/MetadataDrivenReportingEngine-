# Practical Quick Start

[Back to project README](../README.md) | [Getting Started](GettingStarted.md)

## Purpose

This guide presents the fastest practical path to creating a report with the Metadata-Driven Excel Reporting Engine.

The complete documentation explains the engine's internal architecture, but a developer does not need to understand every rendering routine before creating a useful report. The recommended approach is to copy a working configuration, substitute a small representative dataset, and adapt one component at a time.

> Start with working SQL, a valid sorting table, and copied metadata. Generate the workbook, inspect the result, adjust one or two settings, and repeat.

The goal is not to design every metadata record from scratch. The goal is to begin with a proven report definition and turn metadata configuration into a controlled tuning exercise.

## Who This Guide Is For

This guide is intended for analysts and Access developers who regularly produce structured Excel workbooks through a mixture of custom SQL, exports, macros, manual formatting, formulas, subtotals, and chart preparation.

If every reporting cycle requires rebuilding or repairing many of the same workbook elements, the initial effort required to learn this engine can be worthwhile. The engine preserves the reusable reporting machinery so that recurring and related reports can begin from a working SQL, sequencing, metadata, and rendering pattern rather than from another one-off workbook automation project.

The learning process is practical rather than purely theoretical. Expect to:

- Begin with a working demonstration
- Copy existing sequence and metadata records
- Substitute a small representative dataset
- Generate an imperfect first workbook
- Inspect how the output responds to the configuration
- Adjust one or two settings at a time
- Repeat until the report matches the intended design

The architecture usually becomes much easier to understand after the first complete report has been configured. Developers do not need to master every internal routine before experimenting, but they should preserve the documented data, sorting, sequencing, and metadata contracts.

> The engine does not eliminate programming. It eliminates repeatedly programming the same reporting machinery.

## 1. Begin with a Working Demonstration

## 1. Begin with a Working Demonstration

Choose the supplied demonstration that most closely resembles the report you want to create.

Use the Medicare demonstration when the new report resembles:

- One large detail worksheet
- One primary grouping level
- A one-time worksheet header
- Repeating report bands
- Report-level grand totals
- A chart based on the completed worksheet

Use the KPI demonstration when the new report resembles:

- Multiple worksheets
- Multiple sorting tables
- Multiple chart sheets
- Hidden helper columns
- Running footer calculations
- Broader workbook orchestration

Do not begin by creating every sequence row and metadata record from a blank table unless necessary. Copying a working report definition reduces the number of variables that must be solved at the same time.

## 2. Prepare a Small Report-Ready Source

Begin with the smallest useful dataset you can produce. A test source containing approximately 50 to 100 representative rows is usually enough to exercise several report bands, formulas, subtotals, grand totals, and worksheet-header behavior.

The source data can originate from:

- Local Access tables
- Linked Access tables
- SQL Server or another relational database
- Pass-through queries
- Stored procedures
- Imported spreadsheets or text files
- Public datasets
- Any other source the host application can retrieve and stage

A form such as `frmReportManager` can collect the user's parameters through:

- Text boxes
- Dropdown lists
- Date ranges
- Checkboxes
- Option groups
- Other Access controls

Use those parameters to retrieve and prepare the first version of the report data.

If the result is small enough to materialize locally, land it in an Access table. If the source is large, perform the heavy filtering, joining, and aggregation in the source database before bringing the report-ready result into Access.

The final worksheet source should contain:

- The detail or summarized records to display
- All visible report columns
- Explicit sorting fields
- Grouping keys and labels
- SQL-generated business calculations
- Helper fields needed by formulas, headers, footers, totals, or charts

The final source may be a detail query or a grouped summary query. What matters is that it produces the exact worksheet grain required by the report.

Most importantly, sort the final source in the precise order in which the records must appear in Excel. The sorting table described in the next section must summarize the records in that same order.

## 3. Create the Sorting Table

The sorting table tells the break engine how the final worksheet records are divided into report bands.

Its first three fields must contain these concepts in this order:

1. A numeric or otherwise reliably sortable band sequence
2. The display name or description of the band
3. The number of detail records in the band

A simplified result looks like this:

```text
BandSequence    BandDescription                 RecordCount
1               Evaluation and Management       18
2               Imaging                         12
3               Procedures                      24
```

The first field should sort from low to high in the same order as the corresponding groups in the final worksheet source.

The second field contains the label used to identify the group.

The third field contains the number of final-source records belonging to that group.

A typical sorting query uses `GROUP BY` and `Count(*)` against the final worksheet source. It should produce one row per report band.

Before continuing, verify:

```text
Sum of sorting-table RecordCount values
    = total number of records in the final worksheet source
```

If the totals or ordering do not match, the inserted headers, detail rows, subtotals, and footers will not align correctly.

For the full contract and troubleshooting cautions, see the [Getting Started guide](GettingStarted.md) and [`tblReportSequence` Reference](TblReportSequenceReference.md).

## 4. Register the Report in `tblReports`

Open `tblReports` and add the new workbook definition.

The report name should normally include the `.XLSX` extension:

```text
My_New_Report.XLSX
```

The value must match the `Report` value that will be used in:

- `tblReportSequence`
- `tblGoodFieldName_Report`
- `tblGoodFieldName_Header`
- `tblGoodFieldName_G_Totals`

Registering the report first makes the new identifier available in metadata dropdown lists and in `frmReportLauncher`.

## 5. Copy the Four-Row `tblReportSequence` Pattern

Do not build the sequence from an empty table unless necessary. Copy the four rows from the working Medicare demonstration and replace the report-specific values.

The standard sequence is:

```text
Row 1    Final worksheet table marker
Row 2    QueryDef that creates the final worksheet table
Row 3    Sorting-table marker and worksheet association
Row 4    QueryDef that creates the sorting table
```

### Row 1: Final worksheet table marker

This row identifies the table or saved source that will be transferred to Excel.

Configure:

```text
Report:       My_New_Report.XLSX
Sort:         1
QueryName:    tblMyFinalOutput
FinalOutput:  True
Tab:          MyReportTab
Tab Order:    1
isTable?:     True
Exec:         False
```

`QueryName` identifies the final worksheet source.

`Tab` becomes the Excel worksheet name.

`Tab Order` determines the order in which workbook tabs are processed.

`FinalOutput = True` tells the exporter that this source should be copied to a worksheet.

`isTable? = True` tells the preparation routine to delete the previous disposable output table before rebuilding it.

### Row 2: Final worksheet build query

This row creates and executes the QueryDef that builds the table identified in Row 1.

Configure:

```text
Report:       My_New_Report.XLSX
Sort:         2
QueryName:    qryBuildMyFinalOutput
isTable?:     False
SQL:          SELECT ... INTO tblMyFinalOutput ...;
Exec:         True
```

Develop and test this SQL as an ordinary Access query before placing it in `tblReportSequence`.

The SQL may:

- Select prepared detail records
- Group and summarize extracted data
- Add grouping keys and labels
- Add explicit sort fields
- Add reusable business calculations
- Materialize the result with `SELECT ... INTO`

Make sure the SQL creates the same table named in Row 1 and sorts it in the required final worksheet order.

### Row 3: Sorting-table marker

This row identifies the sorting table and associates it with the worksheet defined in Row 1.

Configure:

```text
Report:         My_New_Report.XLSX
Sort:           3
QueryName:      tblSortMyFinalOutput
isTable?:       True
Spacing:        MyReportTab
Spacing Order:  1
HasTotals:      True
```

The critical relationship is:

```text
Row 1 Tab = Row 3 Spacing
```

For example:

```text
Tab:      MyReportTab
Spacing:  MyReportTab
```

`Spacing Order` controls the grouping level. Use `1` for a single-level report.

`HasTotals` determines whether the renderer creates the report-level grand-total area.

Leave `Over_ride_head1` and `Over_ride_foot1` blank initially. Blank override values allow the renderer to use its established defaults.

### Row 4: Sorting-table build query

This row creates and executes the QueryDef that builds the sorting table identified in Row 3.

Configure:

```text
Report:       My_New_Report.XLSX
Sort:         4
QueryName:    qryBuildSortMyFinalOutput
isTable?:     False
SQL:          SELECT ... INTO tblSortMyFinalOutput ...;
Exec:         True
```

The SQL should create the required leading structure:

```text
Numeric sequence
Text band description
Numeric record count
```

It must group and count the same records contained in the final worksheet table and return the groups in the same order.

## 6. Test the Data Contract Before Formatting

Before copying metadata, run the report preparation and inspect the two generated tables.

For the final worksheet table, confirm:

- The expected records are present.
- The fields appear in the required order.
- The records are sorted correctly.
- Each visible and helper field has the expected value.
- Null, zero, and blank values have been tested.

For the sorting table, confirm:

- The first three fields satisfy the sorting-table contract.
- The bands appear in the same order as the final worksheet source.
- Each band count is correct.
- The sum of the band counts equals the final worksheet row count.

Fix data and sorting problems before attempting to correct them through metadata.

## 7. Copy a Working Report Metadata Configuration

Do not create `tblGoodFieldName_Report` from scratch.

Find the working report that most closely resembles the new layout. Sort its metadata by `Sort`, select the complete block of rows, copy it, and append the copied block back into the same metadata table.

Update the copied rows so that:

```text
Report = My_New_Report.XLSX
Tab    = MyReportTab
```

Freezing the columns from `Sort` through `ColumnWidth` can make this work easier because the report, tab, order, letter, visibility, and width fields remain visible while reviewing the remaining metadata.

### Preserve specialized rows initially

At first, leave rows containing specialized header and footer behavior mostly unchanged.

Examples include:

- `HeaderCol_H1_A`
- `FieldName_H1_A`
- `Add_Name_From_Data_H1_A`
- `Position_H1_A`
- `Range_F1`
- `HeaderCol_F2_A`
- `Add_Name_From_Data_F2_A`
- `Position_F2_A`
- Shared H1 and F2 formatting fields

These settings define behavior that is easier to modify after the first workbook renders.

If the new worksheet requires more columns than the copied report, duplicate a generic metadata row that does not contain specialized header or footer instructions. Then adjust:

- `Sort`
- `Letter`
- `FieldName_B`
- `ColumnWidth`
- `Column_Hide`
- Number formats
- Alignment
- Formula fields
- Border settings

Use `Column_Hide = True` for SQL or Excel helper columns that the renderer needs but report users should not see.

For field-level definitions, see the [Metadata Configuration Guide](MetadataConfigurationGuide.md).

## 8. Understand the First Metadata Row

The first record in `Sort` order can be authoritative for section-wide settings.

These may include:

- H1 merge behavior
- H1 fill and borders
- F2 fill and borders
- Row heights
- Visible report boundaries
- Other multi-column formatting

Whenever possible, use an ordinary visible report column as the first metadata record.

Do not assume that placing a shared property on a later metadata row will override the first row.

## 9. Use Helper Columns as a Calculation Workspace

Helper columns can contain values needed by:

- Group headers
- Extended footers
- Running calculations
- Grand totals
- Chart categories
- Chart values
- Intermediate formulas
- Stable sorting

A helper column may come directly from the final SQL source or be generated through metadata.

Header and footer metadata can retrieve values from these columns.

Conceptually, the renderer treats each report band like a small deck of records:

- H1 header values are normally taken from the first detail record in the band.
- Footer values are normally taken from the final detail record or calculated summary position.
- A helper value repeated throughout the band can therefore be projected into a visible header or footer.

For example:

```text
Add_Name_From_Data_H1_A = I
```

This instructs the H1 metadata to retrieve its configured value from Excel column `I`.

This permits SQL-prepared labels and Excel-generated helper calculations to appear in visible report zones without exposing their supporting columns.

## 10. Copy Grand-Total Metadata

If the Row 3 sequence record has:

```text
HasTotals = True
```

copy the closest working configuration from:

```text
tblGoodFieldName_G_Totals
```

Update its `Report` and `Tab` values to match the new worksheet.

Initially preserve the working:

- Labels
- Formula patterns
- Number formats
- Fonts and fills
- Borders
- Merge behavior

After the first workbook renders, adjust the grand-total fields and formulas one at a time.

## 11. Copy Worksheet-Header Metadata When Needed

If the report requires a one-time title and information area, copy a working block from:

```text
tblGoodFieldName_Header
```

Update:

```text
Report = My_New_Report.XLSX
Tab    = MyReportTab
```

The worksheet-level header is separate from the repeating H1 report-band header:

```text
tblGoodFieldName_Header
    -> rendered once at the top of the worksheet

tblGoodFieldName_Report H1
    -> rendered once for every report band
```

The worksheet-header metadata may control:

- Report title
- Subtitle
- Source hyperlink
- Generated-by information
- Generated date and time
- Merged ranges
- Fonts and fills
- Row heights
- Borders
- Freeze-below behavior

For the Medicare reference implementation, report-band construction begins at row 10 so the worksheet-header rows remain reserved:

```vb
Const beginSheet = 10
```

The required relationship is:

```text
beginSheet > final worksheet-header row
```

If no one-time worksheet header is required, do not create matching `tblGoodFieldName_Header` records for that report and tab.

For field-level definitions, see the [Report Header Reference](ReportHeaderReference.md).

## 12. Render the First Small Workbook

Run the report using approximately 50 to 100 representative rows.

The first result may not look correct. That is expected. The purpose of the first render is to prove that:

- The SQL pipeline executes.
- The final worksheet table is created.
- The sorting table is created.
- The worksheet is added.
- The recordset transfers successfully.
- The report bands are inserted.
- The metadata is being read.
- The workbook is saved.

Inspect the output in this order:

1. Raw field-to-column alignment
2. Report-band order
3. Detail-row counts
4. Column headings
5. Hidden helper columns
6. H1 values and positioning
7. Detail formulas
8. F1 subtotals
9. F2 extended footer values
10. Grand totals
11. Worksheet-level header
12. Charts
13. Saved formula values after reopening

## 13. Adjust One or Two Things at a Time

Treat metadata configuration as an iterative tuning process.

A productive loop is:

```text
Change one or two metadata values
    -> generate the workbook
    -> inspect the affected report zone
    -> compare the result with the metadata guide
    -> correct the configuration
    -> repeat
```

Avoid changing many unrelated fields between test runs. Small changes make it much easier to identify the setting responsible for a result or runtime error.

If a runtime error occurs, review:

- The final source field order
- The sorting-table contract
- The `Tab` and `Spacing` association
- Metadata `Sort` values
- Excel column letters
- Header and footer positions
- Formula tokens
- Null, zero, and blank values
- Worksheet-header row reservations

The detailed documentation is most useful at this stage. Consult the guide for the specific report zone or field being adjusted rather than attempting to absorb the entire architecture at once.

See the [Troubleshooting Guide](Troubleshooting.md) for common failure patterns.

## Primary Contracts to Preserve

A developer can initially treat the renderer as a working black box, but these contracts must be respected:

- The final worksheet source must contain the expected fields in the required order.
- The final worksheet records must be sorted in band order.
- The sorting table must begin with the required sequence, description, and record-count fields.
- Sorting-table counts must reconcile to the final worksheet source.
- `tblReportSequence` must associate the worksheet `Tab` with the correct `Spacing` value.
- The first sorted metadata record may control section-wide formatting.
- Worksheet-header rows and report-band starting positions must remain compatible.
- Null, zero, and blank values must be tested because some legacy routines expect specific values.
- Formula metadata must use the supported tokens and should normally omit the leading equals sign.

## Practical Development Model

Creating the first report in a new family requires SQL and metadata work. After that pattern is established, related reports become largely configuration-driven.

The practical development model is:

```text
Prepare the source data
    -> build the sorting contract
    -> register the four-row sequence
    -> copy the closest metadata configuration
    -> render a small test
    -> inspect and tune
    -> expand to the full dataset
```

It is not necessary to understand every internal coordinate calculation before producing a useful report. Start with the demonstrated contracts and working examples. Use the detailed architecture and reference guides when troubleshooting a specific behavior or extending the renderer itself.

## Related Documentation

- [Project README](../README.md)
- [Getting Started](GettingStarted.md)
- [Architectural White Paper](WhitePaper.md)
- [Metadata Configuration Guide](MetadataConfigurationGuide.md)
- [Report Header Reference](ReportHeaderReference.md)
- [`tblReportSequence` Reference](TblReportSequenceReference.md)
- [Troubleshooting Guide](Troubleshooting.md)
- [Medicare Pipeline](WorkedExamples/MedicarePipeline.md)
- [KPI Pipeline](WorkedExamples/KPIPipeline.md)
