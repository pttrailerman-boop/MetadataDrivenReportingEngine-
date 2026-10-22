# Medicare Pipeline Worked Example

[Back to project README](../../README.md) | [Getting Started](../GettingStarted.md) | [Report Header Reference](../ReportHeaderReference.md) | [`tblReportSequence` Reference](../TblReportSequenceReference.md) | [Metadata Configuration Guide](../MetadataConfigurationGuide.md) | [Troubleshooting](../Troubleshooting.md) | [KPI Pipeline](KPIPipeline.md)

## Purpose

This worked example traces the Medicare demo from imported public data through Access SQL, report sequencing, sorting-table construction, ADO transfer, metadata rendering, and the finished Excel workbook.

The Medicare demo is the current reference implementation for the reporting engine. It demonstrates how source-specific SQL can prepare a report-ready dataset while the reusable renderer controls worksheet construction, grouping, formulas, totals, and charts.

## 1. End-to-End Pipeline

```text
CMS Medicare practitioner data
    +
CMS RBCS taxonomy
    |
    v
Normalize RBCS taxonomy
    |
    v
Find latest RBCS release per HCPCS code
    |
    v
Build current RBCS lookup
    |
    v
Join Medicare records to RBCS
    |
    v
Filter to Texas and aggregate
    |
    v
Create display labels and explicit sort fields
    |
    v
Materialize Medicare_ReportFinal
    |
    v
Build tblSort_MedicareFinal
    |
    v
tblReportSequence orchestration
    |
    v
ADO transfer to Medicare_Report worksheet
    |
    v
Apply one-time worksheet header
    |
    v
Break calculation and repeating-band metadata
    |
    v
Grand totals, Top 10 chart, recalculation, save
```

## 2. Imported Source Tables

The demo begins with two imported source tables:

- `Medicare_PHYS_24`, containing the 2024 Medicare Physician and Other Practitioners data used by the demo
- `RBCS_Taxonomy_Final_RY26_20260731`, containing the RY2026 RBCS taxonomy used by the demo

The Medicare source supplies provider, service, HCPCS, utilization, and average payment fields. The RBCS source supplies a classification hierarchy for HCPCS codes.

The reporting engine does not require these specific datasets. They serve as a public-data reference implementation for the broader architecture.

## 3. Normalize the RBCS Taxonomy

The first saved query is `qryRBCS_Normalized_Base`.

Its responsibilities are:

- Trim text values
- Rename source fields to stable internal names
- Convert analysis dates when valid
- Convert the first RBCS release year to a numeric value when valid

```sql
SELECT
    Trim([R].[HCPCS_Cd]) AS HCPCSCode,
    Trim([R].[RBCS_Id]) AS RBCSID,
    Trim([R].[RBCS_Cat]) AS RBCSCategory,
    Trim([R].[RBCS_Cat_Desc]) AS RBCSCategoryDescription,
    Trim([R].[RBCS_Cat_Subcat]) AS RBCSSubcategory,
    Trim([R].[RBCS_Subcat_Desc]) AS RBCSSubcategoryDescription,
    Trim([R].[RBCS_FamNumb]) AS RBCSFamilyNumber,
    Trim([R].[RBCS_Family_Desc]) AS RBCSFamilyDescription,
    IIf(
        IsDate([R].[RBCS_Analysis_Start_Dt]),
        CDate([R].[RBCS_Analysis_Start_Dt]),
        Null
    ) AS AnalysisStartDate,
    IIf(
        IsDate([R].[RBCS_Analysis_End_Dt]),
        CDate([R].[RBCS_Analysis_End_Dt]),
        Null
    ) AS AnalysisEndDate,
    IIf(
        IsNumeric([R].[First_RBCS_Release_Year]),
        CLng([R].[First_RBCS_Release_Year]),
        Null
    ) AS FirstReleaseYear
FROM RBCS_Taxonomy_Final_RY26_20260731 AS R;
```

### Why normalize first

Normalizing once creates stable field names and types for every downstream query. This avoids repeating `Trim`, `CDate`, and `CLng` logic throughout the pipeline.

## 4. Identify the Current RBCS Release

The taxonomy can contain more than one release record for an HCPCS code. `qryRBCS_MaxReleaseYear` finds the greatest valid first release year per code.

```sql
SELECT
    HCPCSCode,
    Max(FirstReleaseYear) AS MaxFirstReleaseYear
FROM qryRBCS_Normalized_Base
WHERE FirstReleaseYear IS NOT NULL
GROUP BY HCPCSCode;
```

This query produces one maximum release-year value for each normalized HCPCS code.

## 5. Build the Current RBCS Lookup

`qryRBCS_Current` joins the normalized rows to the maximum release-year query by both HCPCS code and release year.

```sql
SELECT
    R.HCPCSCode,
    R.RBCSID,
    R.RBCSCategory,
    R.RBCSCategoryDescription,
    R.RBCSSubcategory,
    R.RBCSSubcategoryDescription,
    R.RBCSFamilyNumber,
    R.RBCSFamilyDescription
FROM
    qryRBCS_Normalized_Base AS R
    INNER JOIN qryRBCS_MaxReleaseYear AS M
        ON R.FirstReleaseYear = M.MaxFirstReleaseYear
       AND R.HCPCSCode = M.HCPCSCode;
```

The current lookup is then materialized as `tblRBCS_Current` for the Medicare aggregation step.

### Why materialize the lookup

The materialized table provides a stable local lookup for the larger Medicare join and makes the intermediate result easy to inspect during debugging.

## 6. Aggregate the Medicare Source

`qryMake_Medicare_ReportSummary` joins Medicare records to the current taxonomy with a left join.

The query:

- Matches `Medicare_PHYS_24.HCPCS_Cd` to `tblRBCS_Current.HCPCSCode`
- Filters to Texas
- Excludes rows with null service totals
- Groups records by state, provider type, place of service, RBCS category, and RBCS subcategory
- Uses `Unclassified` when an RBCS classification is unavailable
- Calculates estimated totals by multiplying service counts by source average amounts
- Writes the result to `tblMedicare_ReportSummary`

```sql
SELECT
    M.Rndrng_Prvdr_State_Abrvtn AS State,
    M.Rndrng_Prvdr_Type AS ProviderType,
    M.Place_Of_Srvc AS PlaceOfService,
    Nz(R.RBCSCategoryDescription,"Unclassified") AS RBCSCategory,
    Nz(R.RBCSSubcategoryDescription,"Unclassified") AS RBCSSubcategory,
    Sum(Nz(M.Tot_Srvcs,0)) AS ServiceCount,
    Sum(Nz(M.Tot_Srvcs,0) * Nz(M.Avg_Sbmtd_Chrg,0))
        AS EstimatedSubmittedCharges,
    Sum(Nz(M.Tot_Srvcs,0) * Nz(M.Avg_Mdcr_Alowd_Amt,0))
        AS EstimatedAllowedAmount,
    Sum(Nz(M.Tot_Srvcs,0) * Nz(M.Avg_Mdcr_Pymt_Amt,0))
        AS EstimatedMedicarePayment
INTO tblMedicare_ReportSummary
FROM
    Medicare_PHYS_24 AS M
    LEFT JOIN tblRBCS_Current AS R
        ON M.HCPCS_Cd = R.HCPCSCode
WHERE
    M.Tot_Srvcs Is Not Null
    AND M.Rndrng_Prvdr_State_Abrvtn = "TX"
GROUP BY
    M.Rndrng_Prvdr_State_Abrvtn,
    M.Rndrng_Prvdr_Type,
    M.Place_Of_Srvc,
    Nz(R.RBCSCategoryDescription,"Unclassified"),
    Nz(R.RBCSSubcategoryDescription,"Unclassified");
```

### Estimated financial fields

The source provides average charge and payment values. The demo estimates totals as:

```text
Service count x source average amount
```

These fields are reporting estimates produced for the demonstration, not audited payment totals.

## 7. Create the Final Report Query

`qryMedicare_ReportFinal` converts the aggregated source into a renderer-ready recordset.

It adds:

- Friendly facility and non-facility descriptions
- `RBCSBandKey` for grouping
- `RBCSBandLabel` for display
- `ProviderTypeSort` for stable provider ordering
- `PlaceOfServiceSort` for numeric place-of-service ordering

```sql
SELECT
    S.State,
    S.ProviderType,
    IIf(
        S.PlaceOfService = "F",
        "Facility",
        IIf(
            S.PlaceOfService = "O",
            "Non-Facility",
            Nz(S.PlaceOfService,"Unknown")
        )
    ) AS PlaceOfServiceDescription,
    S.RBCSCategory,
    S.RBCSSubcategory,
    Nz(S.RBCSCategory,"Unclassified") & "|" &
        Nz(S.RBCSSubcategory,"Unclassified") AS RBCSBandKey,
    Nz(S.RBCSCategory,"Unclassified") & " - " &
        Nz(S.RBCSSubcategory,"Unclassified") AS RBCSBandLabel,
    UCase(Trim(Nz(S.ProviderType,"Unknown"))) AS ProviderTypeSort,
    IIf(
        S.PlaceOfService = "F",
        1,
        IIf(S.PlaceOfService = "O",2,9)
    ) AS PlaceOfServiceSort,
    S.ServiceCount,
    S.EstimatedSubmittedCharges,
    S.EstimatedAllowedAmount,
    S.EstimatedMedicarePayment
FROM tblMedicare_ReportSummary AS S
ORDER BY
    S.RBCSCategory,
    S.RBCSSubcategory,
    S.EstimatedMedicarePayment DESC,
    UCase(Trim(Nz(S.ProviderType,"Unknown"))),
    IIf(S.PlaceOfService = "F",1,IIf(S.PlaceOfService = "O",2,9));
```

### Why the query includes both keys and labels

`RBCSBandKey` provides a stable grouping value for the renderer and sorting table. `RBCSBandLabel` provides a friendlier string for visible report content.

### Why explicit sort fields matter

Display text does not always produce the required business order. Numeric and normalized sort fields make the final order deterministic.

## 8. Configure the Four `tblReportSequence` Rows

The Medicare workbook uses the standard four-row sequence pattern.

### Sort 1: mark the final table for cleanup and output

```text
Report:       Medicare_Report.XLSX
Sort:         1
QueryName:    Medicare_ReportFinal
FinalOutput:  True
Tab:          Medicare_Report
Tab Order:    1
isTable?:     True
```

This row deletes the prior `Medicare_ReportFinal` table and identifies the rebuilt table as the worksheet source.

### Sort 2: build the final worksheet table

```text
Report:       Medicare_Report.XLSX
Sort:         2
QueryName:    qryMedicareReport_Final
isTable?:     False
Exec:         True
```

Stored SQL:

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

This order is the final physical order copied into Excel.

### Sort 3: mark and associate the sorting table

```text
Report:           Medicare_Report.XLSX
Sort:             3
QueryName:        tblSort_MedicareFinal
isTable?:         True
Spacing:          Medicare_Report
Spacing Order:    1
HasTotals:        True
Over_ride_head1:  4
Over_ride_foot1:  3
```

This row deletes the previous sorting table, associates the new sorting table with the worksheet, enables grand totals, and applies custom primary band allocations.

### Sort 4: build the sorting table

```text
Report:       Medicare_Report.XLSX
Sort:         4
QueryName:    qryInsertInto_tblSortMedicareFinal
isTable?:     False
Exec:         True
```

Stored SQL:

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

The first three output fields satisfy the sorting-table contract:

```text
BandSequence
RBCSBandKey
RecCount
```

## 9. Identify the Worksheet Source and Sorting Source

For the Medicare worksheet:

```text
Worksheet tab:       Medicare_Report
Final worksheet source: Medicare_ReportFinal
Sorting table:       tblSort_MedicareFinal
Grouping key:        RBCSBandKey
Grouping level:      1
Grand totals:        enabled
```

The final source contains every detail row. The sorting table contains one row per RBCS band with the number of detail records in that band.

## 10. Pass the Prepared Output to Excel

The export routine reads the final-output rows, creates the workbook and worksheet, then opens the configured source through ADO.

The essential handoff is:

```vb
TableString = Trim(rsTab![QueryName])
rsa.Open TableString, conn
oSheet.Range("A2").CopyFromRecordset rsa
```

For the Medicare report, `TableString` resolves to `Medicare_ReportFinal`.

The records are copied beginning at cell `A2`. The first row remains available for field names or intermediate workbook construction before metadata formatting restructures the sheet.

## 11. Apply the Worksheet-Level Report Header

Before the repeating RBCS report bands are constructed, the Medicare renderer applies the one-time worksheet header defined in `tblGoodFieldName_Header`.

This metadata layer is separate from `tblGoodFieldName_Report`:

```text
tblGoodFieldName_Header
    -> rendered once at the top of the worksheet

tblGoodFieldName_Report H1
    -> rendered once for every RBCS report band
```

The Medicare worksheet header demonstrates:

- The report title `Medicare Report for Texas, 2024 Data.`
- A public CMS source hyperlink
- The subtitle `Metadata Reporting Engine Demonstration`
- The generated-by label `Report Generated by A.F.`
- A generated date and time formula
- Merged title and information ranges
- Header fonts, sizes, styles, colors, and fills
- Header row heights
- Border formatting
- Freeze-below behavior

The visible label for the public-source hyperlink is:

```text
Medicare Physician & Other Practitioners - by Provider and Service
```

The generated timestamp follows this Excel formula pattern:

```excel
"Report Generated at: " & TEXT(NOW(),"mm/dd/yyyy h:mm AM/PM")
```

### Reserved-row prerequisite

The worksheet header requires enough reserved rows above the first generated report band.

In `fn_FindBreaks`, the Medicare reference implementation uses:

```vb
Const original = 2
Const beginSheet = 10
```

`original` identifies the initial row position of the exported records before report-band rows are inserted.

`beginSheet` identifies where the generated report-band area begins. If `beginSheet` is less than 10, report-band construction can occupy the rows reserved for `tblGoodFieldName_Header`, causing the worksheet header to be overwritten, displaced, or omitted.

When expanding the worksheet header, maintain this rule:

```text
beginSheet > final worksheet-header row
```

The worksheet header should be validated before troubleshooting the repeating H1, B, B_D, F1, or F2 zones.

For complete field definitions, see the [Report Header Reference](../ReportHeaderReference.md).

## 12. Generate the Physical Report Bands

After data transfer, the renderer:

1. Maps `tblSort_MedicareFinal` to `Medicare_Report` through the `Spacing` value.
2. Writes the sorting-table name to `tblNameHolder`.
3. Calls `fn_FindBreaks`.
4. Calculates the header, detail, and footer coordinates for each RBCS band.
5. Writes those coordinates to `tbl_temp_sorting_grouping`.
6. Inserts physical rows into Excel.
7. Applies metadata to each generated band.

Conceptually:

```text
RBCS band count
    + header allocation
    + body-heading allocation
    + detail-row count
    + footer allocation
    = physical Excel band size
```

## 13. Apply Repeating Report-Band Metadata

`tblGoodFieldName_Report` controls the ordinary Medicare report bands.

The metadata defines:

- Excel column letters
- Column widths and hidden columns
- H1 report headers
- B column headings
- B_D detail formats and formulas
- F1 subtotals
- F2 extended summaries
- Border, fill, font, and alignment settings

The first metadata row in `Sort` order should contain section-wide H1, F2, and boundary settings.

The left and right edges of shared ranges are based on the first and last visible Excel columns.

## 14. Apply Formula Templates

The renderer replaces formula tokens with calculated row positions.

Examples:

```text
SUM(K@:K#)
```

becomes a band-specific sum such as:

```excel
=SUM(K18:K35)
```

A detail calculation such as:

```text
IFERROR(M?/L?,0)
```

is initialized at the first detail row and filled through the band's detail range.

Store formula metadata without a leading equals sign because the VBA adds it.

## 15. Create Grand Totals

Because the Medicare sorting-table row has `HasTotals = True`, the formatter adds the report-level grand-total block using `tblGoodFieldName_G_Totals`.

Grand totals are created after the standard RBCS bands because the final block summarizes the completed report rather than one group.

Validate that:

- Service counts reconcile to the detail rows
- Estimated submitted charges reconcile
- Estimated allowed amounts reconcile
- Estimated Medicare payments reconcile
- Derived percentages use the intended totals

## 16. Create the Medicare Top 10 Chart

The Medicare renderer creates a Top 10 chart after the worksheet data, report bands, formulas, and totals are available.

Chart generation remains code-based in the reference implementation. The worksheet and any supporting calculation columns must therefore be complete before the chart routine executes.

When adapting the chart:

1. Confirm the chart label range.
2. Confirm the chart value range.
3. Confirm the source rows are calculated before chart creation.
4. Confirm hidden support columns remain available.
5. Reopen the saved workbook and verify the chart values.

## 17. Force Recalculation Before Save

The renderer may disable automatic calculation while formatting for performance. The current Medicare implementation restores calculation and forces a full calculation before saving.

This step prevents the standalone workbook from being saved with stale formula values or stale chart sources.

The final sequence is conceptually:

```text
Restore automatic calculation
    -> calculate workbook
    -> full calculation rebuild
    -> process pending events
    -> save workbook
```

## 18. Run the Medicare Report

1. Open `frmReportManager`.
2. Review the prepared Medicare records.
3. Open `frmReportLauncher`.
4. Select `Medicare_Report.XLSX`.
5. Choose whether to enable **Render Sheet?**.
6. Select **Create Spreadsheets**.
7. Choose the output path and filename.
8. Wait for the launcher to display the completed workbook hyperlink.
9. Open the workbook and validate the worksheet and chart.

## 19. Validate the Pipeline

### Source validation

- [ ] HCPCS values are normalized before joining.
- [ ] Maximum RBCS release year is selected per HCPCS code.
- [ ] Current taxonomy rows join by both code and release year.
- [ ] Unmatched taxonomy values become `Unclassified`.
- [ ] Medicare records are filtered to Texas.
- [ ] Service totals are non-null.

### Summary validation

- [ ] `ServiceCount` is aggregated correctly.
- [ ] Estimated amounts equal service count times source averages before aggregation.
- [ ] Facility and non-facility labels are correct.
- [ ] `RBCSBandKey` is stable and non-null.
- [ ] Provider and place-of-service sorts are explicit.

### Sequence validation

- [ ] Final-table cleanup precedes the build query.
- [ ] Sorting-table cleanup precedes the sorting build query.
- [ ] Both action QueryDefs use `Exec = True`.
- [ ] `Spacing` exactly matches `Medicare_Report`.
- [ ] Sorting-table counts sum to final-source rows.
- [ ] The sorting table's first three fields satisfy the break-engine contract.

### Worksheet-header validation

- [ ] Matching `tblGoodFieldName_Header` records exist for `Medicare_Report.XLSX` and `Medicare_Report`.
- [ ] The Medicare report title appears once at the top of the worksheet.
- [ ] The CMS source hyperlink displays the intended label.
- [ ] The demonstration subtitle appears in the configured merged range.
- [ ] The generated-by label appears in the intended location.
- [ ] The generated date and time formula calculates correctly.
- [ ] Header merge ranges do not overlap.
- [ ] Header row heights, fonts, fills, and borders appear correctly.
- [ ] `beginSheet` is at least 10.
- [ ] `beginSheet` is greater than the final worksheet-header row.
- [ ] The first RBCS report band begins below the worksheet header.
- [ ] Freeze-below behavior uses the intended worksheet row.

### Workbook validation

- [ ] RBCS bands align with the copied records.
- [ ] H1, B, B_D, F1, and F2 zones appear correctly.
- [ ] Formula tokens resolve to valid rows.
- [ ] Grand totals reconcile.
- [ ] The Top 10 chart uses the intended values.
- [ ] Reopened workbook values are current.

## 20. The 99-Row Test Variant

Use the Medicare test source before a full-volume run.

The 99-row strategy is intended to exercise:

- Multiple RBCS bands
- Different band sizes
- Row insertion
- Detail formulas
- Subtotals
- Grand totals
- Chart preparation

Recommended loop:

```text
Run test-source SQL
    -> build test final output
    -> build test sorting table
    -> render workbook
    -> inspect every band
    -> correct SQL or metadata
    -> repeat
    -> restore full source
```

## 21. What This Example Demonstrates

The Medicare pipeline shows that the reporting engine separates five concerns:

```text
Source normalization
    -> handled by Access SQL

Current reference selection
    -> handled by Access SQL

Business aggregation
    -> handled by Access SQL

Workbook orchestration
    -> handled by tblReportSequence

Workbook presentation
    -> handled by metadata, VBA, and Excel
```

That separation is the main reusable pattern. A different application can replace the Medicare-specific SQL while retaining the sequencing, sorting-table contract, metadata model, and Excel renderer.

## 22. Related Documentation

- [Project README](../../README.md)
- [Getting Started](../GettingStarted.md)
- [Architectural White Paper](../WhitePaper.md)
- [Report Header Reference](../ReportHeaderReference.md)
- [Metadata Configuration Guide](../MetadataConfigurationGuide.md)
- [`tblReportSequence` Reference](../TblReportSequenceReference.md)
- [Troubleshooting Guide](../Troubleshooting.md)
- [KPI Pipeline](KPIPipeline.md)
