# Metadata-Driven Excel Reporting Engine
## Architectural White Paper

[Back to project README](../README.md) | [Getting Started](GettingStarted.md) | [Report Header Reference](ReportHeaderReference.md) | [Metadata Configuration Guide](MetadataConfigurationGuide.md) | [`tblReportSequence` Reference](TblReportSequenceReference.md) | [Troubleshooting](Troubleshooting.md) | [Medicare Pipeline](WorkedExamples/MedicarePipeline.md) | [KPI Pipeline](WorkedExamples/KPIPipeline.md)

**Author:** Alex F. 
**Technology:** Microsoft Access, Access SQL, DAO, ADO, VBA, Excel COM automation  
**Reference implementations:** Medicare reporting demo and KPI reporting demo

## Abstract

The Metadata-Driven Excel Reporting Engine is a Microsoft Access and VBA framework for generating recurring, presentation-ready Excel workbooks from SQL-backed data. It addresses a common business-reporting gap: reports that are too structured, repetitive, and error-prone to assemble manually, but that do not require a specialized financial, scientific, statistical, or engineering modeling platform.

The engine separates application-specific data preparation from reusable workbook construction. Access SQL, saved queries, action queries, and prepared tables produce report-ready records and compact sorting tables. An ordered manifest in `tblReportSequence` controls object cleanup, QueryDef creation, action-query execution, worksheet selection, grouping associations, totals, and optional band-size overrides. VBA and Excel COM automation then create the workbook, transfer source records, calculate physical report-band coordinates, apply formatting metadata, generate formulas, add totals and charts, and save a standalone `.xlsx` file.

The architecture originated in 2009 from a practical need to create Excel reports that were more than data dumps. It has since been reused for multiple reporting problems. The public repository includes two complementary demonstrations: the current Medicare reference implementation and an earlier KPI implementation that demonstrates multiple worksheets, multiple charts, hidden formula columns, running footer calculations, and broader orchestration.

## 1. Relational Principles Applied to Spreadsheet Construction

The engine applies relational database principles to spreadsheet construction.

SQL defines the business result and prepares the records in their required order. Sorting and grouping tables describe the logical structure of the report. Microsoft Access calculates the physical worksheet coordinates required for each report band. Excel performs the final calculation, presentation, charting, and workbook rendering.

The workbook is therefore an output of a database-controlled report model rather than the primary location of the reporting logic.

This distinction shaped the architecture from the beginning. A conventional Excel-centered solution might start with a workbook template containing formulas, named ranges, tables, formatting, and macros. Over time, the workbook can become both the application and the report.

This engine takes the opposite approach:

```text
Relational data and SQL
    |
    | define the business result and record order
    v
Sorting and grouping tables
    |
    | describe the logical report structure
    v
Break calculations
    |
    | determine physical worksheet coordinates
    v
Report metadata
    |
    | defines formulas, formatting, and report zones
    v
Excel
    |
    | calculates and presents the final business artifact
    v
Standalone XLSX workbook
```
The worksheet is treated as a numeric coordinate space. The break engine determines the vertical row coordinates, while metadata and column-conversion utilities determine the horizontal column coordinates. Once both dimensions are known, the renderer can address and control any cell, range, report band, helper region, subtotal, grand-total block, or chart source within a worksheet.

The workbook adds another coordinate to the model:

```text
Workbook
    -> worksheet
    -> report band
    -> report zone
    -> row
    -> column
```

## 2. The Reporting Problem

Business users frequently need an Excel workbook because Excel is familiar, portable, inspectable, and easy to distribute. The source facts may already exist in Access, SQL Server, linked tables, or public datasets, but the expected deliverable is not a raw recordset.

A finished operational report may require:

- Repeating report sections with different numbers of detail rows
- Group headers and descriptive labels
- Column-heading rows
- Detail formatting by data type
- Subtotals and report-level grand totals
- Rates, percentages, variances, and running values
- Hidden helper columns
- Multiple worksheets
- Summary and chart sheets
- Stable, repeatable formatting
- A standalone workbook suitable for distribution

A simple `TransferSpreadsheet` operation or `CopyFromRecordset` call can move data into Excel, but it does not construct a report. At the other extreme, advanced financial models, scientific workbooks, simulations, and engineering calculation systems require specialized mathematics and domain logic that are outside this engine's purpose.

The intended niche is therefore:

> **Beyond a data dump, short of a modeling platform.**

The engine is strongest when the primary challenge is constructing a complex recurring workbook from prepared data rather than performing advanced analytical computation.

## 3. Design Goals

The framework was designed around several practical goals.

### 3.1 Preserve the host application's data architecture

The engine does not require an application to abandon its existing tables, saved queries, forms, or SQL process. The host application prepares a final record source using the approach appropriate to its business problem.

### 3.2 Separate data preparation from presentation

SQL should define the report-ready facts, grouping keys, classifications, and sort values. Metadata and VBA should define workbook structure, formulas, and presentation behavior.

### 3.3 Reuse one renderer across multiple reports

A new report should be able to reuse the same workbook-generation components by supplying:

- An ordered sequence of SQL and object instructions
- One or more final worksheet sources
- One or more sorting/grouping sources
- Report metadata
- Optional chart-specific code

### 3.4 Support variable-length report bands

The renderer must construct groups whose detail sections contain different numbers of records while still placing headers, subtotals, footers, and formulas correctly.

### 3.5 Produce a standalone workbook

The output should be a normal `.xlsx` file that can be opened without Microsoft Access after generation.

## 4. Architectural Overview

The architecture has three major layers.

```text
Host application
    |
    | parameters, retrieval, application-specific SQL
    v
Report-ready Access sources
    |
    | ordered instructions in tblReportSequence
    v
Report orchestrator
    |
    | final worksheet sources + sorting tables
    v
ADO transfer into Excel
    |
    | calculated break coordinates
    v
Repeating report-band renderer
    |
    | formulas, subtotals, and grand totals
    v
Worksheet-header metadata
    |
    | chart creation, recalculation, save
    v
Standalone XLSX workbook
```

### 4.1 Layer 1: Host application

The host application owns source-specific behavior:

- User parameters
- Retrieval from local, linked, or external data sources
- Application-specific joins and calculations
- Data cleanup and normalization
- Preparation of report-ready records
- Preview of prepared results

The public forms demonstrate where this logic belongs. Their **Get Data** buttons are instructional because the public demos include prepared sample data rather than live private retrieval logic.

### 4.2 Layer 2: Report orchestrator

`frmReportLauncher` is the bridge between the host application and the reusable renderer. It:

- Validates the selected report
- Reads ordered rows from `tblReportSequence`
- Clears disposable output tables
- Creates or replaces saved QueryDefs
- Executes configured action queries
- Creates worksheet and spacing views
- Maps sorting/grouping sources to worksheets
- Calls `createSpreadsheetReport`

### 4.3 Layer 3: Workbook generator

`modExportToExcel` and `clsExport` perform the workbook work:

- Prompt for an output path
- Create or replace the workbook
- Create and name worksheets
- Open configured sources through ADO
- Copy records beginning at cell `A2`
- Calculate report breaks
- Insert rows for report zones
- Apply repeating report-band formatting from `tblGoodFieldName_Report`
- Generate formulas
- Create subtotals and grand totals
- Hide helper columns
- Apply the one-time worksheet header from `tblGoodFieldName_Header`
- Create charts
- Recalculate and save the workbook

## 5. `tblReportSequence` as an Execution Manifest

`tblReportSequence` is the central orchestration table. It is not merely a list of reports. It describes an ordered workflow.

Important fields include:

- `Report`: workbook or report definition
- `Sort`: execution order
- `Version`: optional configuration version
- `QueryName`: table or QueryDef associated with the row
- `SQL`: Access SQL used to define or execute the step
- `Exec`: whether a created QueryDef is executed
- `isTable?`: whether the row identifies a disposable table rather than a QueryDef
- `FinalOutput`: whether the named source is copied into a worksheet
- `Tab`: worksheet identifier
- `Tab Order`: worksheet-processing order
- `Spacing`: sorting/grouping source associated with a worksheet
- `Spacing Order`: grouping-level order
- `HasTotals`: whether report-level grand totals are added
- `Over_ride_head1`: optional primary-header allocation
- `Over_ride_foot1`: optional primary-footer allocation
- `Notes`: implementation guidance

### 5.1 Object lifecycle

Rows follow one of two preparation paths.

```text
isTable? = True
    -> QueryName identifies a disposable output/cache table
    -> delete the table before recreation

isTable? = False
    -> QueryName identifies a QueryDef
    -> delete and recreate the QueryDef
    -> execute it only when Exec = True
```

This distinction matters because the pipeline can contain both object definitions and output-table markers.

### 5.2 Action queries and saved select queries

`clsCreateRecord` creates a saved QueryDef from the SQL stored in the sequence row.

- With `Exec = True`, the QueryDef is executed as an action query.
- With `Exec = False`, the QueryDef is saved without execution.

The demonstrated pipelines use action QueryDefs containing `SELECT ... INTO` statements to materialize final worksheet tables and sorting tables.

### 5.3 Final worksheet transfer

The final transfer is recordset-oriented rather than table-specific:

```vb
TableString = Trim(rsTab![QueryName])
rsa.Open TableString, conn
oSheet.Range("A" & RECORDSET_START_ROW).CopyFromRecordset rsa
```
`RECORDSET_START_ROW` is a public constant shared by the ADO transfer routine and the break-calculation logic. In the current implementation, its value is `2`. Using one shared value prevents the transfer routine and `fn_FindBreaks` from developing different assumptions about where the original records begin.

ADO opens the configured named source and copies the returned records into Excel. The public configurations use materialized local tables as final worksheet sources. The implementation also supports creating a saved select QueryDef with `Exec = False`; a saved query can therefore serve as a final source when it returns the required fields, preserves the required ordering, and has no unresolved parameters. This direct-query path is supported by the code but is not represented by the supplied demo configurations.

## 6. Worksheet and Spacing Association

The launcher builds separate views of `tblReportSequence` for:

- The full ordered report sequence
- Worksheet definitions
- Spacing or grouping definitions
- The worksheet-to-spacing association

It creates temporary saved queries named:

- `spcngStrToTab_1`
- `spcngStrToTab_2`
- `spcngStrToTab`

The final association joins the spacing definition to the worksheet definition where `Spacing = Tab`. This mapping tells the renderer which sorting or grouping source belongs to each worksheet and preserves settings such as `HasTotals`, `Over_ride_head1`, and `Over_ride_foot1`.

This design permits one workbook definition to coordinate multiple worksheets and multiple grouping sources without hard-coding each worksheet inside the transfer routine.

## 7. The Sorting-Table Contract

The worksheet source contains the records displayed in Excel. The sorting table contains a compact inventory of the groups into which those records will be divided.

The primary sorting input must provide three required concepts in the expected first columns:

1. A sortable sequence or grouping key
2. A display description or group label
3. A record count for the group

A simplified example is:

```text
SortKey    Description       RecordCount
202407     July 2024         8
202408     August 2024       12
202409     September 2024    17
```

The sorting table does not contain the detail records. It tells the break engine how many records belong to each ordered group.

### 7.1 Why field order matters

`fn_FindBreaks` depends on the required headers and positions. A sorting table may contain correct values and still fail if the required first fields are missing, misnamed, or arranged incorrectly.

The sorting table should be treated as a formal interface contract, not as an informal helper query.

## 8. Break Calculation and Physical Row Planning

Before the engine formats the workbook, it must convert logical group counts into physical Excel coordinates.

### 8.1 `tblNameHolder`

The renderer creates or clears `tblNameHolder` and adds the name of each grouping source associated with the current worksheet. This table supplies the ordered grouping levels used by the recursive break calculation.

### 8.2 `fn_FindBreaks`

`fn_FindBreaks` recursively evaluates the grouping sources and writes the calculated layout into `tbl_temp_sorting_grouping`.

The resulting rows contain values such as:

- `Head_B_1`: first row of the primary header allocation
- `Head_E_1`: final row of the primary header allocation
- `Begin1`: first detail row
- `End1`: final detail row
- `Sum_Beg_1`: first primary-summary row
- `Sum_End_1`: final footer allocation row
- Group description
- Group level
- Inserted-row counts

These calculated coordinates become the physical map for the workbook.

### 8.3 `fn_break`

`fn_break` inserts the required rows into Excel so the copied recordset becomes a report layout. The engine then formats the inserted header, body, detail, and footer regions using the calculated coordinates.

The order is important:

```text
Copy source records
    -> calculate logical breaks
    -> insert physical rows
    -> apply formulas and formatting
```
### 8.4 Reserved worksheet-header rows

The one-time worksheet header occupies reserved rows above the first generated report band. This header is controlled by `tblGoodFieldName_Header`, not by the repeating H1 metadata.

```text
tblGoodFieldName_Header
    -> rendered once at the top of the worksheet

tblGoodFieldName_Report H1
    -> rendered once for every report band
```
The Medicare reference implementation requires `fn_FindBreaks` to begin the generated report-band area at row 10 or later:

```vb
Const original = 2
Const beginSheet = 10
```
`original` identifies the initial row position of the exported records before report-band rows are inserted.

`beginSheet` identifies where the first generated report band begins. If `beginSheet` is too low, report-band construction can occupy the rows reserved for the worksheet header.

The required relationship is:

```text
beginSheet > final worksheet-header row
```
For complete header-field definitions and Medicare examples, see the [Report Header Reference](ReportHeaderReference.md).

## 9. Report-Band Anatomy

The report-band zones described in this section repeat for each grouping band. They are separate from the one-time worksheet header defined in `tblGoodFieldName_Header`.

An ordinary report block contains the following zones:

```text
H1   Primary group header
B    Body or column-heading row
B_D  Detail-data rows
F1   Primary subtotal or summary row
F2   Extended footer or additional summary rows
```

### 9.1 H1: Repeating primary group header

The H1 zone can contain:

- Static labels
- Values copied from report data
- Group descriptions
- Positioned label/value pairs
- Merged header content
- Section-wide fill and border formatting

### 9.2 B: Column-heading row

The B zone provides visible headings for the detail columns. A heading may be supplied directly from `FieldName_B` or obtained from another configured source.

### 9.3 B_D: Detail-data region

The B_D zone applies formatting and formulas to the report records. It controls alignment, number formats, fonts, colors, row height, borders, merge instructions, and metadata-generated formulas.

### 9.4 F1: Primary subtotal row

The F1 zone commonly contains:

- `Totals:` labels
- Band sums
- Ratios or rates
- Values copied from the final detail row
- Formula-token expressions

### 9.5 F2: Extended footer

The F2 zone can contain multiple positioned rows. It is suitable for additional metrics, running values, summary panels, or values retrieved from hidden helper columns.

### 9.6 Grand totals

`tblGoodFieldName_G_Totals` controls report-level totals after ordinary report bands have been completed. Grand totals are separate from F1 and F2 because they summarize the full report rather than a single group.

## 10. Metadata Must Be Read in Two Dimensions

One of the least obvious aspects of the design is that the metadata table must be interpreted both vertically and horizontally.

### 10.1 Vertical dimension: metadata record as Excel column

Each metadata record normally represents an Excel destination column. Fields such as these are evaluated for each record:

- `Letter`
- `ColumnWidth`
- `Column_Hide`
- `FieldName_B`
- B_D formatting
- `Range_B_D`
- `FieldName_F1`
- `Range_F1`

This is the conventional interpretation:

```text
Metadata row for K -> configure Excel column K
Metadata row for L -> configure Excel column L
Metadata row for M -> configure Excel column M
```

### 10.2 Horizontal dimension: metadata fields as report-zone instructions

Some fields do not describe one cell in one output column. They describe a shared H1, F2, body boundary, or multi-column range.

The renderer often:

1. Finds the first visible Excel column.
2. Finds the last visible Excel column.
3. Constructs one range spanning those boundaries.
4. Returns the metadata recordset to its first row.
5. Reads shared formatting settings from the first row in `Sort` order.

The first sorted metadata record should therefore be treated as the authoritative configuration row for section-wide properties.

### 10.3 Four metadata behavior classes

A useful documentation classification is:

- **C: Column-specific**: evaluated independently for each Excel column
- **I: Repeatable item**: multiple positioned header or footer label/value items may be defined
- **S: Section-wide**: controls a range spanning the visible report width
- **U: Unused or unconfirmed**: present in the schema but not consumed by the current renderer

Examples:

```text
C  Letter, ColumnWidth, Column_Hide, Range_B_D, Range_F1
I  FieldName_H1_A, Position_H1_A, FieldName_F2_A, Position_F2_A
S  Merge_Center_H1, ColorIndex_H1, H1 borders, F2 borders
U  List_B, ErrorTitle_B, ErrorMessage_B, Lock_B
```

This classification prevents a common mistake: placing a section-wide setting on a later metadata row and expecting it to override the first row.

## 11. Formula Templates and Tokens

The formula-token system allows metadata to store reusable Excel formula templates without knowing the final worksheet row numbers.

Formula metadata is stored without a leading equals sign. The VBA renderer adds `=` when assigning the formula to Excel.

### 11.1 `?`: context row

In `Range_B_D`, `?` is replaced with the first detail row before Excel fills the formula through the selected range.

```text
IFERROR(M?/L?,0)
```

If the detail block begins at row 18, the initial formula becomes:

```excel
=IFERROR(M18/L18,0)
```

Excel then adjusts relative references as the formula is filled through the detail range.

In `Range_F1`, `?` refers to the footer context used by that routine.

### 11.2 `@`: first detail row

In footer formulas, `@` represents the first detail row of the current band.

### 11.3 `#`: last detail row

In footer formulas, `#` represents the final detail row of the current band.

```text
SUM(K@:K#)
```

For a band covering rows 18 through 35, the engine produces:

```excel
=SUM(K18:K35)
```

### 11.4 Engine-resolved and workbook-resolved logic

The formula system combines two calculation layers:

```text
Engine-resolved logic
    -> replace report-position tokens with physical row numbers

Workbook-resolved logic
    -> let Excel evaluate SUMIFS, INDEX, ROW, IF, IFERROR, and other functions
```

This makes the formula layer substantially more flexible than a collection of fixed cell references.

## 12. Hidden Helper Columns as a Calculation Workspace

Not every output column must come from the SQL source. A developer can add a metadata record for an additional Excel column, place a formula in `Range_B_D` or another supported formula field, hide the column, and reuse the result elsewhere in the workbook.

Hidden columns can hold:

- Ratios and rates
- Running values
- Sort ranks
- Conditional flags
- Display labels
- Error-handled calculations
- Chart categories and values
- Intermediate formula components

The visible report can then retrieve or reference only the presentation-ready result.

This creates a second calculation layer:

```text
Prepared SQL facts
    -> hidden metadata-generated calculations
    -> visible report fields and footers
    -> charts and summaries
```

## 13. KPI Running-Count Example

The KPI demo provides a concrete example of hidden helper calculations feeding visible footer values.

A running formula uses `SUMIFS`, `INDEX`, and `ROW()` to sum qualifying values above the current footer while excluding inserted subtotal rows labeled `Totals:`.

A representative formula is:

```excel
=SUMIFS(
    $K$1:INDEX(K:K,ROW()-1),
    $J$1:INDEX(J:J,ROW()-1),
    "<>Totals:"
)
```

The formula works without knowing the final footer row in advance:

- `ROW()-1` ends the range immediately above the current formula.
- `INDEX` creates a dynamic range endpoint.
- The `"<>Totals:"` criterion excludes subtotal rows and avoids double counting.
- The hidden result can be exposed through an F2 label such as **Run. Count of Samples** or **Run. Count of Late WO**.

This example demonstrates four reusable ideas:

1. Metadata-generated helper columns
2. Position-independent Excel formulas
3. Formulas that distinguish data rows from structural rows
4. Selective projection of hidden results into visible footers

## 14. Choosing Where a Calculation Belongs

The framework supports several places for calculations. The correct location depends on the purpose of the value.

### 14.1 Calculate in Access SQL when the value is:

- Required for filtering
- Required for grouping
- Required for sorting
- Shared by multiple outputs
- Part of the business-data definition

### 14.2 Calculate in a hidden Excel helper column when the value is:

- Presentation-specific
- Needed only by the workbook
- Conveniently expressed through relative cell references
- Used by another Excel formula
- Intended to feed a footer or chart

### 14.3 Calculate directly in a report zone when the value is:

- A subtotal
- A grand total
- A band-level ratio
- A short label/value summary
- Specific to the physical layout of a report band

This separation prevents the two most common extremes: forcing every presentation calculation into SQL or placing core business logic entirely inside Excel.

## 15. Default and Override Behavior

The renderer defines default row allocations for primary and secondary headers and footers. In the current Medicare implementation:

- Blank `Over_ride_head1` values use the primary-header default.
- Blank `Over_ride_foot1` values use the primary-footer default.
- Secondary header and footer allocations retain their configured defaults.

This behavior is important because a blank override should mean “use the framework default,” not “allocate zero rows.”

When custom overrides are supplied, they must still provide enough rows for the metadata positions and formulas used by the report.

## 16. Workbook Rendering Lifecycle

The overall workbook lifecycle is:

```text
1. Prompt for output path
2. Read ordered tblReportSequence rows
3. Delete disposable tables
4. Create or replace QueryDefs
5. Execute configured action queries
6. Create workbook and worksheets
7. Open each final source through ADO
8. Copy each recordset to its worksheet
9. Build worksheet-to-spacing associations
10. Calculate break coordinates
11. Insert report-band rows
12. Apply repeating report-band metadata
13. Generate formulas
14. Create subtotals and grand totals
15. Apply one-time worksheet-header metadata
16. Create charts
17. Recalculate workbook formulas
18. Save and close workbook
19. Display output hyperlink
```

The current Medicare renderer explicitly restores automatic calculation, forces full calculation, runs a full calculation rebuild, and processes pending events before saving. This prevents a workbook constructed under manual calculation mode from being saved with stale formula values.

## 17. Reference Implementation: Medicare Demo

The Medicare demo is the current reference implementation.

### 17.1 Source data

The demo uses:

- CMS 2024 Medicare Physician and Other Practitioners data imported as `Medicare_PHYS_24`
- CMS RY2026 RBCS taxonomy imported as `RBCS_Taxonomy_Final_RY26_20260731`

### 17.2 Taxonomy preparation

`qryRBCS_Normalized_Base` trims and renames taxonomy fields and converts analysis dates and the first release year.

`qryRBCS_MaxReleaseYear` identifies the maximum first release year for each HCPCS code.

A current-RBCS query joins normalized rows to the maximum release year by HCPCS code and release year, producing `tblRBCS_Current`.

### 17.3 Medicare aggregation

The Medicare source is left-joined to the current RBCS table by HCPCS code, filtered to Texas, and grouped by:

- State
- Provider type
- Place of service
- RBCS category
- RBCS subcategory

Estimated totals are calculated by multiplying service counts by the source averages. The result is written to `tblMedicare_ReportSummary`.

### 17.4 Final report source

The final report query adds:

- Facility and non-facility descriptions
- `RBCSBandKey`
- `RBCSBandLabel`
- Uppercase trimmed `ProviderTypeSort`
- Numeric `PlaceOfServiceSort`

The final source is paired with a sorting table and passed to the renderer.

### 17.5 Workbook output

The Medicare output demonstrates:

- A large grouped detail report
- Financial totals and calculated rates
- A worksheet report header
- Formula-generated columns
- Grand totals
- A Medicare Top 10 chart
- Explicit recalculation before save

### 17.6 Worksheet-level header

The Medicare worksheet uses `tblGoodFieldName_Header` to create a one-time report header above the repeating RBCS bands.

The header demonstrates:

- The title `Medicare Report for Texas, 2024 Data.`
- A public CMS source hyperlink
- The subtitle `Metadata Reporting Engine Demonstration`
- The generated-by label `Report Generated by A.F.`
- A generated date and time formula
- Merged title and information ranges
- Fonts, fills, row heights, and borders
- Freeze-below behavior

The generated timestamp follows this Excel expression pattern:

```excel
"Report Generated at: " & TEXT(NOW(),"mm/dd/yyyy h:mm AM/PM")
```

The one-time worksheet header is independent of the repeating H1 zone. The header is created once, while H1 is rendered for every RBCS report band.

## 18. Reference Implementation: KPI Demo

The KPI demo is an earlier implementation retained because it demonstrates broader workbook composition.

### 18.1 Prepared source

The object named `qryKPI_Results` is a local prepared table despite its query-style prefix. It contains the common KPI fields used by both branches.

### 18.2 Monthly and Special branches

Saved select queries filter the prepared source into Monthly and Special result sets and attach the beginning and ending reporting years.

`tblReportSequence` then runs action SQL that materializes:

- `KPI_Results_Monthly`
- `tblSorKPI_MonthlyReport`
- `KPI_Results_Special`
- `tblSorKPI_SpecialReport`

### 18.3 Workbook output

The KPI workbook demonstrates:

- Two report worksheets
- Two chart worksheets
- Variable-length monthly bands
- Default header and footer allocations
- Hidden calculated columns
- Running footer counts
- Grand totals included in chart behavior
- Multiple formula-token patterns

The KPI and Medicare demos therefore serve different purposes. Medicare represents the current engine baseline, while KPI demonstrates broader multi-sheet orchestration.

## 19. Chart Generation

Chart creation in the public implementations is code-based rather than a completely generic chart-metadata subsystem.

- The KPI renderer calls chart functions after formatting each worksheet and adds grand-total content to the charts.
- The Medicare renderer creates a Top 10 chart from prepared worksheet data.

Charts are created after the report bands, formulas, and totals exist because the chart routines depend on the final worksheet structure.

A future version could move more chart behavior into metadata, but the existing code demonstrates that the workbook pipeline can prepare chart-ready hidden columns, labels, summary ranges, and final chart sheets.

## 20. Reserved and Unimplemented Metadata Fields

The metadata schema contains fields that are not consumed by the current reference renderer. Populating a field does not activate behavior unless the VBA reads and applies that field.

Currently unconfirmed or unimplemented fields include:

- `PasteSpecial`
- `PasswordProtect`
- `Password`
- `List_B`
- `ErrorTitle_B`
- `ErrorMessage_B`
- `Lock_B`

These fields should be described as reserved, legacy, or not implemented in the current reference renderer. They should not be presented as working features until code and test coverage are added.

## 21. Development and Debugging Strategy

Metadata behaves like executable configuration. A small error in a field name, Excel letter, formula token, position, or sorting-table contract can produce a runtime error or a structurally incorrect workbook.

A practical development loop is:

```text
Copy a working metadata configuration
    -> substitute a small representative source
    -> change one setting
    -> render the workbook
    -> inspect the result
    -> correct metadata
    -> repeat
    -> restore the full source
    -> run a volume test
```

### 21.1 The 99-row Medicare test strategy

The Medicare demo includes a 99-row test source. This is large enough to exercise multiple report bands and formulas while remaining small enough for rapid rendering and manual review.

Recommended validation order:

1. Verify source field names and order.
2. Verify sorting-table headers and record counts.
3. Verify `tblReportSequence` order.
4. Verify worksheet and spacing associations.
5. Verify matching `tblGoodFieldName_Header` records.
6. Verify worksheet-header row numbers and merge ranges.
7. Verify `beginSheet` is at least 10 and greater than the final worksheet-header row.
8. Verify title, formulas, hyperlinks, and freeze behavior.
9. Verify header and footer allocations.
10. Verify hidden helper columns.
11. Verify formula-token replacement.
12. Verify subtotals and grand totals.
13. Verify chart source ranges.
14. Restore full data and validate workbook recalculation.

## 22. Error-Prone Areas

Developers adapting the engine should pay particular attention to:

### 22.1 Sorting-table headers

`fn_FindBreaks` expects a specific leading structure. Incorrect field names or positions can break the report map.

### 22.2 Metadata sort order

The first sorted metadata record can be authoritative for section-wide formatting. Moving or hiding that record can change behavior unexpectedly.

### 22.3 Formula syntax

Formula metadata should normally omit the leading equals sign because the renderer supplies it.

### 22.4 Band allocations

Header and footer positions must fit within the allocated number of rows. A custom override that is too small can cause overlap or misplaced output.

### 22.5 Worksheet-header row reservation

The worksheet-level header requires reserved rows above the first generated report band.

For the Medicare reference implementation:

```vb
Const original = 2
Const beginSheet = 10
```

If `beginSheet` is less than 10, or is not greater than the final worksheet-header row, report-band construction can overwrite, displace, or omit the worksheet header.

### 22.6 Parameterized queries

A final saved query opened through ADO must have all parameters resolved before the transfer stage.

### 22.7 Calculation mode

The renderer may disable automatic calculation for performance. The workbook must be recalculated before save when formulas were created under manual calculation mode.

### 22.8 Excel process cleanup

Errors during COM automation can leave an Excel process open. Development testing should include cleanup and recovery procedures.

## 23. Appropriate Use Cases

The engine is well suited to recurring operational and analytical reports such as:

- Billing and payment summaries
- Healthcare utilization reports
- Compliance and audit workbooks
- Reconciliation and exception reports
- Project-status reports
- Laboratory turnaround metrics
- Service-level and timeliness KPIs
- Inventory and purchasing summaries
- Quality-control reports
- Claims or transaction analysis
- Government and public-data reporting
- Vendor-performance reports
- Monthly management workbooks

It can also present finalized results from scientific, engineering, or financial systems when those results already exist in a suitable data source. It is not intended to perform the specialized underlying analysis itself.

## 24. Out-of-Scope Use Cases

The engine is not intended to replace:

- Advanced financial modeling platforms
- Scientific computing environments
- Statistical analysis packages
- Engineering simulation systems
- General-purpose business-intelligence platforms
- No-code report designers
- Every native Excel feature

These boundaries strengthen the project by defining a realistic and useful scope.

## 25. Reuse and Adoption Model

The project does not require all adopters to use the entire database unchanged. Reuse can occur at several levels:

1. Adopt the full reporting engine inside an Access application.
2. Adopt `tblReportSequence` as an execution-manifest pattern.
3. Adopt the sorting-table and break-calculation contract.
4. Adopt `tblGoodFieldName_Header` as a one-time worksheet-header configuration layer.
5. Adopt the H1, B, B_D, F1, and F2 report-zone model.
6. Adopt the formula-token system.
7. Adopt hidden helper columns as a metadata-controlled calculation workspace.
8. Use the Medicare SQL pipeline as a public-data reporting example.
9. Use the KPI demo as a multi-sheet and multi-chart orchestration example.

Partial adoption is a valid outcome. A developer may reuse one architectural idea without adopting the complete renderer.

## 26. Design Tradeoffs

### 26.1 Benefits

- Keeps source-specific SQL outside the renderer
- Allows report layouts to change without rewriting every VBA statement
- Supports multiple worksheets and grouping sources
- Handles variable-length report bands
- Produces standalone Excel output
- Encourages repeatable testing and configuration reuse

### 26.2 Costs

- Metadata is dense and requires documentation
- Some settings are not obvious from the table alone
- Section-wide settings depend on record order
- Debugging requires understanding both Access and Excel coordinates
- COM automation is slower than a simple export
- Chart code remains implementation-specific
- The current engine has legacy and reserved fields

The framework trades initial configuration complexity for reuse across recurring reports.

## 27. Future Documentation and Extension Points

The repository documentation should continue with:

- A field-by-field metadata configuration guide
- A complete `tblReportSequence` reference
- A worked Medicare SQL pipeline
- A worked KPI pipeline
- A formula-pattern cookbook
- A troubleshooting guide
- Annotated screenshots of forms, report bands, helper columns, and charts

Potential code extensions include:

- Metadata validation before rendering
- Explicit metadata capability flags
- Generic chart metadata
- Query parameter resolution
- Improved COM cleanup
- Logging and timing instrumentation
- A formal schema version for metadata tables

## Related Documentation

- [Project README](../README.md)
- [Getting Started](GettingStarted.md)
- [Report Header Reference](ReportHeaderReference.md)
- [Metadata Configuration Guide](MetadataConfigurationGuide.md)
- [`tblReportSequence` Reference](TblReportSequenceReference.md)
- [Troubleshooting Guide](Troubleshooting.md)
- [Medicare Pipeline](WorkedExamples/MedicarePipeline.md)
- [KPI Pipeline](WorkedExamples/KPIPipeline.md)
  
## 28. Conclusion

The Metadata-Driven Excel Reporting Engine addresses a broad but often underserved business need: recurring Excel reports that are too complex for routine manual assembly or a flat export, but not specialized enough to justify a modeling platform.

Its central design achievement is the separation of concerns:

```text
SQL prepares the facts and grouping order.
Metadata describes the report.
VBA compiles that description into Excel.
Excel evaluates the completed formulas.
```

The Medicare demo proves the current architecture against a substantial public dataset. The KPI demo proves that the same design can coordinate multiple worksheets, multiple charts, hidden calculations, running footer metrics, and broader report sequencing.

The result is best understood not as a single report and not merely as an Access-to-Excel export routine, but as a reusable report compiler for structured business workbooks.
