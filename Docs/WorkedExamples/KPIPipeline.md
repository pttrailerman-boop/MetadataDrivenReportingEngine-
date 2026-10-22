# KPI Pipeline Worked Example

[Back to project README](../../README.md) | [Getting Started](../GettingStarted.md) | [`tblReportSequence` Reference](../TblReportSequenceReference.md) | [Metadata Configuration Guide](../MetadataConfigurationGuide.md) | [Medicare Pipeline](MedicarePipeline.md)

## Purpose

This worked example traces the KPI demo from its prepared Access data through the Monthly and Special branches, `tblReportSequence`, sorting tables, ADO worksheet transfer, metadata-generated formulas, grand totals, running footer values, and chart generation.

The KPI demo complements the Medicare reference implementation. Medicare shows the current single-detail-report baseline, while KPI demonstrates broader orchestration with two report worksheets, two sorting sources, hidden calculation columns, report-level totals, and multiple charts.

## 1. End-to-End Pipeline

```text
Prepared local KPI data
    |
    v
qryKPI_Results local table
    |
    +-----------------------------+
    |                             |
    v                             v
Monthly filter                Special filter
    |                             |
    v                             v
KPI_Results_Monthly          KPI_Results_Special
    |                             |
    v                             v
tblSorKPI_MonthlyReport      tblSorKPI_SpecialReport
    |                             |
    +-------------+---------------+
                  |
                  v
          tblReportSequence
                  |
                  v
      ADO transfer to two worksheets
                  |
                  v
       Break calculation by month
                  |
                  v
 Metadata formatting and helper formulas
                  |
                  v
    Subtotals, running values, grand totals
                  |
                  v
             Chart sheets
```

## 2. Prepared Source Object

The source object named `qryKPI_Results` is a local Access table despite its query-style name.

Its fields include:

- `ProjName`
- `WorkOrdr`
- `SmplPerE`
- `Month`
- `CountOfSampleId`
- `SampReleaseDate`
- `LastSampRecDate`
- `ReportDueDate`
- `Category`
- `Late`
- `CountOfLate`
- `Year`
- `monthtxt`
- `YearNumber`

The table contains the common fields used by both KPI report branches.

### Why the source is prepared first

The workbook renderer should not contain application-specific retrieval logic. The host application prepares a stable local source, and the reporting sequence transforms that source into worksheet-specific outputs.

## 3. Split the Source into Monthly and Special Branches

Two saved select queries divide the prepared source by `Category`.

### 3.1 Monthly branch

`qryKPI_Results_Monthly` selects KPI records whose category is stored as `Monthy` in the supplied demo configuration. The spelling is part of the existing sample data and query definition, so it must remain consistent unless both are corrected together.

The query also includes values from:

- `qryBegin_Yr`
- `qryEnd_yr`

The records are ordered by `SmplPerE`.

Conceptually:

```sql
SELECT
    qryKPI_Results.ProjName,
    qryKPI_Results.WorkOrdr,
    qryKPI_Results.SmplPerE,
    qryKPI_Results.Month,
    qryKPI_Results.CountOfSampleId,
    qryKPI_Results.SampReleaseDate,
    qryKPI_Results.LastSampRecDate,
    qryKPI_Results.ReportDueDate,
    qryKPI_Results.Category,
    qryKPI_Results.Late,
    qryKPI_Results.CountOfLate,
    qryKPI_Results.Year,
    qryKPI_Results.monthtxt,
    qryKPI_Results.YearNumber,
    qryBegin_Yr.Begin_Yr,
    qryEnd_yr.End_Yr
FROM
    qryKPI_Results,
    qryBegin_Yr,
    qryEnd_yr
WHERE qryKPI_Results.Category = "Monthy"
ORDER BY qryKPI_Results.SmplPerE;
```

### 3.2 Special branch

`qryKPI_Results_Specials` selects records whose category is `Special`, includes the same year-boundary values, and orders the records by `SmplPerE`.

Conceptually:

```sql
SELECT
    qryKPI_Results.ProjName,
    qryKPI_Results.WorkOrdr,
    qryKPI_Results.SmplPerE,
    qryKPI_Results.Month,
    qryKPI_Results.CountOfSampleId,
    qryKPI_Results.SampReleaseDate,
    qryKPI_Results.LastSampRecDate,
    qryKPI_Results.ReportDueDate,
    qryKPI_Results.Category,
    qryKPI_Results.Late,
    qryKPI_Results.CountOfLate,
    qryKPI_Results.Year,
    qryKPI_Results.monthtxt,
    qryKPI_Results.YearNumber,
    qryBegin_Yr.Begin_Yr,
    qryEnd_yr.End_Yr
FROM
    qryKPI_Results,
    qryBegin_Yr,
    qryEnd_yr
WHERE qryKPI_Results.Category = "Special"
ORDER BY qryKPI_Results.SmplPerE;
```

## 4. Materialize the Monthly Final Source

The Monthly build row creates `qryKPI_Results_Monthly_2` and executes it as a make-table query.

The query writes `KPI_Results_Monthly` and adds workbook-ready measures:

- `countofWorkOrder`, assigned a value of 1 per source row
- `PercentWoOnTime`
- `CountofLateSamples`
- `PercentSamplesOnTime`
- `Year_Range`

Key calculations are:

```text
PercentWoOnTime
    = (CountOfWorkOrder - CountOfLate) / CountOfWorkOrder

CountOfLateSamples
    = CountOfSampleId when CountOfLate > 0, otherwise 0

PercentSamplesOnTime
    = (CountOfSampleId - CountOfLateSamples) / CountOfSampleId
```

The final action SQL follows this pattern:

```sql
SELECT
    qryKPI_Results_Monthly.ProjName,
    qryKPI_Results_Monthly.WorkOrdr,
    qryKPI_Results_Monthly.SmplPerE,
    qryKPI_Results_Monthly.Month,
    qryKPI_Results_Monthly.SampReleaseDate,
    qryKPI_Results_Monthly.LastSampRecDate,
    qryKPI_Results_Monthly.ReportDueDate,
    qryKPI_Results_Monthly.Category,
    qryKPI_Results_Monthly.Late,
    qryKPI_Results_Monthly.CountOfLate,
    1 AS countofWorkOrder,
    ([Countofworkorder]-[countofLate])/[Countofworkorder]
        AS PercentWoOnTime,
    qryKPI_Results_Monthly.CountOfSampleId,
    IIf([countoflate]>0,[Countofsampleid],0)
        AS CountofLateSamples,
    ([countofsampleid]-[countoflatesamples])/[countofsampleid]
        AS PercentSamplesOnTime,
    qryKPI_Results_Monthly.Year,
    qryKPI_Results_Monthly.monthtxt,
    qryKPI_Results_Monthly.YearNumber,
    [Begin_Yr] & " To " & [End_yr] AS Year_Range
INTO KPI_Results_Monthly
FROM qryKPI_Results_Monthly
ORDER BY qryKPI_Results_Monthly.SmplPerE;
```

## 5. Build the Monthly Sorting Table

The Monthly sorting query creates `tblSorKPI_MonthlyReport`.

```sql
SELECT
    [Year] & Left([YearNumber],2) AS SpecialSort,
    [monthtxt] & " " & [Year] AS Month_Year,
    Count(KPI_Results_Monthly.WorkOrdr) AS CountOfWorkOrdr
INTO tblSorKPI_MonthlyReport
FROM KPI_Results_Monthly
GROUP BY
    [Year] & Left([YearNumber],2),
    [monthtxt] & " " & [Year]
ORDER BY
    [Year] & Left([YearNumber],2);
```

The first three fields satisfy the grouping contract:

```text
SpecialSort       ordered month key
Month_Year        visible month description
CountOfWorkOrdr   detail-record count
```

## 6. Materialize the Special Final Source

The Special branch repeats the same calculation pattern and writes `KPI_Results_Special` from `qryKPI_Results_Specials`.

It creates the same workbook-ready fields:

- Count of work orders
- Percent of work orders on time
- Count of samples
- Count of late samples
- Percent of samples on time
- Year range

Keeping the Monthly and Special outputs structurally compatible allows both worksheets to use the same renderer architecture while retaining separate metadata and sorting sources.

## 7. Build the Special Sorting Table

The Special sorting query creates `tblSorKPI_SpecialReport`.

```sql
SELECT
    [Year] & Left([YearNumber],2) AS SpecialSort,
    [monthtxt] & " " & [Year] AS Month_Year,
    Count(KPI_Results_Special.WorkOrdr) AS CountOfWorkOrdr
INTO tblSorKPI_SpecialReport
FROM KPI_Results_Special
GROUP BY
    [Year] & Left([YearNumber],2),
    [monthtxt] & " " & [Year]
ORDER BY
    [Year] & Left([YearNumber],2);
```

The supplied QueryDef name is `qryInsertInto_SorKPI_SpeicalReport`. The word `Speical` is misspelled in that existing object name, so the sequence row and saved object must remain consistent unless both are renamed together.

## 8. Configure the Eight `tblReportSequence` Rows

The KPI workbook uses two four-row branches under one report definition, `Monthly_Report.XLSX`.

### 8.1 Monthly branch

#### Sort 1: Monthly final-output marker

```text
QueryName:    KPI_Results_Monthly
FinalOutput:  True
Tab:          KPI_Monthly
Tab Order:    1
isTable?:     True
HasTotals:    False on this final-output row
```

#### Sort 2: Monthly final-output build query

```text
QueryName:    qryKPI_Results_Monthly_2
isTable?:     False
Exec:         True
```

#### Sort 3: Monthly sorting-table marker

```text
QueryName:     tblSorKPI_MonthlyReport
isTable?:      True
Spacing:       KPI_Monthly
Spacing Order: 1
HasTotals:     True
```

#### Sort 4: Monthly sorting-table build query

```text
QueryName:    qryInsertInto_SorKPI_MonthlyReport
isTable?:     False
Exec:         True
```

### 8.2 Special branch

#### Sort 5: Special final-output marker

```text
QueryName:    KPI_Results_Special
FinalOutput:  True
Tab:          KPI_Special
Tab Order:    2
isTable?:     True
```

#### Sort 6: Special final-output build query

```text
QueryName:    qryKPI_Results_Special_2
isTable?:     False
Exec:         True
```

#### Sort 7: Special sorting-table marker

```text
QueryName:     tblSorKPI_SpecialReport
isTable?:      True
Spacing:       KPI_Special
Spacing Order: 2
HasTotals:     True
```

#### Sort 8: Special sorting-table build query

```text
QueryName:    qryInsertInto_SorKPI_SpeicalReport
isTable?:     False
Exec:         True
```

## 9. Identify the Worksheet and Sorting Sources

```text
Worksheet 1
    Tab:            KPI_Monthly
    Final source:   KPI_Results_Monthly
    Sorting table:  tblSorKPI_MonthlyReport
    Tab Order:      1

Worksheet 2
    Tab:            KPI_Special
    Final source:   KPI_Results_Special
    Sorting table:  tblSorKPI_SpecialReport
    Tab Order:      2
```

Both worksheet branches belong to `Monthly_Report.XLSX`, so one launcher selection generates one workbook containing both report tabs.

## 10. Transfer Both Sources to Excel

The exporter reads final-output rows in tab order.

For each row, it:

1. Creates or selects the Excel worksheet.
2. Opens the configured `QueryName` through ADO.
3. Copies the recordset beginning at `A2`.
4. Advances to the next worksheet.

Conceptually:

```vb
TableString = Trim(rsTab![QueryName])
rsa.Open TableString, conn
oSheet.Range("A2").CopyFromRecordset rsa
```

The first pass produces the raw `KPI_Monthly` and `KPI_Special` worksheets. Formatting occurs afterward.

## 11. Calculate Monthly Report Bands

The renderer associates each sorting table with its worksheet through `Spacing`.

`fn_FindBreaks` converts month descriptions and record counts into physical Excel coordinates. The generated `tbl_temp_sorting_grouping` contains values including:

- Month sort key
- Sorting-table level
- Month description
- Detail count
- First and last detail rows
- Header boundaries
- Footer boundaries
- Inserted-row counts

Each month becomes a separate report band with its own heading, detail region, subtotal, and extended footer.

## 12. Use Default Header and Footer Allocations

The KPI sequence leaves `Over_ride_head1` and `Over_ride_foot1` blank.

The renderer therefore uses its default primary allocations:

- Primary header allocation: 5 rows
- Primary footer allocation: 4 rows
- Secondary header allocation: 3 rows
- Secondary footer allocation: 3 rows

This demonstrates the intended blank-override behavior: null means use the renderer default, not zero rows.

## 13. Apply KPI Metadata

The two worksheets have separate metadata rows in `tblGoodFieldName_Report` and separate report-level total rows in `tblGoodFieldName_G_Totals`.

The metadata controls:

- Column letters and widths
- Hidden columns
- Month headers
- Column headings
- Detail formatting
- Work-order and sample metrics
- Percent formats
- Monthly subtotals
- Running values
- Grand totals

The first sorted metadata record should hold section-wide settings such as H1 and F2 fills, merged ranges, and shared boundary lines.

## 14. Formula Columns and KPI Metrics

The final output tables already contain several calculated measures, including percentages. The Excel metadata adds another calculation layer for presentation-specific behavior.

Grand-total metadata includes patterns such as:

```text
Sum(K@:K#)
Sum(L@:L#)
(L?-K?)/L?
Sum(N@:N#)
Sum(O@:O#)
(N?-O?)/N?
```

These expressions calculate:

- Total late work orders
- Total work orders
- Percent of work orders on time
- Total samples
- Total late samples
- Percent of samples on time

The engine replaces the row tokens with the appropriate calculated positions and supplies the leading equals sign.

## 15. Hidden Helper Columns

The KPI implementation uses hidden columns as a workbook calculation workspace.

A metadata row can:

1. Target an Excel column that does not come directly from the Access source.
2. Populate `Range_B_D` with a formula template.
3. Apply the formula through each month band.
4. Hide the column with `Column_Hide`.
5. Reuse the calculated value in F2, grand totals, or chart preparation.

This separates visible report design from intermediate workbook calculations.

## 16. Running Footer Calculations

The KPI demo uses running formulas that accumulate values above the current footer while excluding subtotal rows labeled `Totals:`.

A representative pattern is:

```excel
=SUMIFS(
    $K$1:INDEX(K:K,ROW()-1),
    $J$1:INDEX(J:J,ROW()-1),
    "<>Totals:"
)
```

The pattern works because:

- `ROW()-1` ends the range immediately above the formula.
- `INDEX` creates a dynamic endpoint.
- The criterion excludes subtotal rows.
- The formula does not need to know the final footer row in advance.

The hidden result can be exposed in F2 through labels such as:

- `Run. Count of Samples`
- `Run. Count of Late WO`

The exact source and criterion columns must be verified against the worksheet metadata before adapting the pattern.

## 17. Create Grand Totals

Both sorting-table marker rows use `HasTotals = True`.

The renderer therefore creates grand-total blocks for `KPI_Monthly` and `KPI_Special` using `tblGoodFieldName_G_Totals`.

The supplied grand-total metadata includes:

- A `Grand Totals:` label
- Sums for late work orders and work orders
- Work-order on-time percentage
- Sums for samples and late samples
- Sample on-time percentage
- Month and year summary content

Grand totals are created after the monthly report bands because they summarize the completed worksheet.

## 18. Create the Chart Sheets

The KPI renderer adds chart sheets after formatting the report worksheets.

The demo shows:

- Multiple charts
- Chart source ranges based on completed worksheet structure
- Grand-total content included in chart behavior
- Worksheet helper values supporting chart generation

Chart creation remains code-based rather than fully metadata-driven. The report bands, formulas, and total rows must therefore exist before the chart functions run.

## 19. Validate Monthly and Special Outputs

### Source checks

- [ ] `qryKPI_Results` contains the required fields.
- [ ] Category values match the Monthly and Special filter queries.
- [ ] `Begin_Yr` and `End_Yr` resolve correctly.
- [ ] Records are ordered by `SmplPerE`.

### Sequence checks

- [ ] All eight rows use `Monthly_Report.XLSX`.
- [ ] Each table marker precedes its make-table query.
- [ ] Each action-query row uses `Exec = True`.
- [ ] `KPI_Monthly` uses `Tab Order = 1`.
- [ ] `KPI_Special` uses `Tab Order = 2`.
- [ ] Each sorting-table row uses the matching `Spacing` value.

### Sorting checks

- [ ] Month keys sort chronologically.
- [ ] Month descriptions are readable.
- [ ] Group counts reconcile to final-source rows.
- [ ] The first three sorting fields satisfy the break-engine contract.

### Workbook checks

- [ ] Both report worksheets are present.
- [ ] Monthly bands align with detail records.
- [ ] Work-order metrics reconcile.
- [ ] Sample metrics reconcile.
- [ ] Hidden helper columns calculate correctly.
- [ ] Running values exclude subtotal rows.
- [ ] Grand totals reconcile to monthly bands.
- [ ] Chart sheets use the intended completed ranges.

## 20. Common KPI Configuration Errors

### Monthly branch returns no rows

Check the actual source value `Monthy`. Correcting the spelling requires changing both the prepared data and the filter query.

### Special sorting QueryDef is not found

The supplied name contains `Speical`. Keep the sequence row and QueryDef name synchronized.

### Month bands are out of order

Verify the `SpecialSort` expression and confirm its text values sort chronologically.

### Running totals double-count

Confirm the criterion column contains `Totals:` on subtotal rows and that the formula excludes it.

### Percentages show division errors

Add or retain appropriate `IFERROR` behavior where zero denominators are possible.

### Charts use incomplete ranges

Run chart creation only after row insertion, formulas, subtotals, and grand totals are complete.

## 21. What the KPI Example Demonstrates

The KPI pipeline proves that one report definition can coordinate:

```text
One prepared source
    -> two filtered branches
    -> two materialized worksheet sources
    -> two sorting tables
    -> two formatted report worksheets
    -> hidden calculation columns
    -> running footer values
    -> two grand-total blocks
    -> multiple charts
```

The broader lesson is that `tblReportSequence` is not limited to a single worksheet. It can orchestrate an entire workbook while the metadata remains worksheet-specific.

## 22. Related Documentation

- [Project README](../../README.md)
- [Getting Started](../GettingStarted.md)
- [Architectural White Paper](../WhitePaper.md)
- [Metadata Configuration Guide](../MetadataConfigurationGuide.md)
- [`tblReportSequence` Reference](../TblReportSequenceReference.md)
- [Medicare Pipeline](MedicarePipeline.md)
