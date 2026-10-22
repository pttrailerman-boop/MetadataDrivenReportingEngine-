# `tblReportSequence` Reference

[Back to project README](../README.md) | [Architectural White Paper](WhitePaper.md) | [Metadata Configuration Guide](MetadataConfigurationGuide.md)

## Purpose

`tblReportSequence` is the execution manifest for the Metadata-Driven Excel Reporting Engine. It tells the launcher and workbook generator which objects to delete, which QueryDefs to create, which action queries to execute, which record sources become worksheets, which sorting tables define grouping, and whether totals or custom band allocations are required.

This table should be read as an ordered program stored in data.

```text
Report selection
    -> ordered sequence rows
    -> table cleanup
    -> QueryDef creation
    -> action-query execution
    -> worksheet-source selection
    -> sorting-table association
    -> Excel transfer and formatting
```

## 1. Field Reference

### `id`

- **Purpose:** Unique row identifier.
- **Typical type:** AutoNumber or Long Integer.
- **Execution effect:** None by itself.
- **Guidance:** Do not rely on `id` for processing order. Use `Sort`.

### `Report`

- **Purpose:** Identifies the report or workbook definition to which the sequence row belongs.
- **Examples:** `Medicare_Report.XLSX`, `Monthly_Report.XLSX`
- **Used by:** The launcher filters the sequence to the selected report.
- **Important:** Every row required for one workbook must use the same report identifier.

### `Sort`

- **Purpose:** Controls execution order within the selected report.
- **Typical type:** Long Integer.
- **Important:** This is the workflow order, not merely a display order.

A common pattern is:

```text
1  Mark final worksheet table for cleanup and final output
2  Create and execute the action QueryDef that rebuilds it
3  Mark sorting table for cleanup and associate it with the worksheet
4  Create and execute the action QueryDef that rebuilds the sorting table
```

For multiple worksheets, repeat the pattern in the required worksheet order.

### `version`

- **Purpose:** Optional configuration version.
- **Execution effect:** No confirmed active behavior in the current public sequence loop.
- **Guidance:** Use for human-readable configuration tracking unless the renderer is extended to enforce it.

### `QueryName`

- **Purpose:** Names the Access table or saved QueryDef associated with the row.
- **Meaning depends on the row:**
  - When `[isTable?] = True`, it identifies a table to delete before recreation.
  - When `[isTable?] = False`, it identifies the QueryDef to create or replace.
  - When `FinalOutput = True`, it identifies the record source copied into Excel.
  - When `Spacing` is populated, it identifies the sorting/grouping source.

A single `QueryName` can participate in more than one role, but each sequence row should have one clear purpose.

### `FinalOutput`

- **Purpose:** Marks a named source as a worksheet source.
- **Typical type:** Yes/No.
- **Effect:** Rows with a worksheet definition are included in the tab query passed to the exporter.
- **Important:** The final source must return fields in the order expected by the metadata and sorting logic.

The demonstrated configurations use local tables as final outputs. The transfer code opens the configured name through ADO, so a saved select QueryDef can also serve as a final source when it returns a normal recordset and has no unresolved parameters.

### `Tab`

- **Purpose:** Identifies the Excel worksheet associated with a final output.
- **Examples:** `Medicare_Report`, `KPI_Monthly`, `KPI_Special`
- **Important:** The worksheet name and metadata `Tab` value must agree.
- **Behavior:** Spaces may be replaced with underscores by the formatter when selecting the worksheet.

### `Tab Order`

- **Purpose:** Controls worksheet creation and processing order.
- **Typical type:** Long Integer.
- **Important:** Use unique, intentional values for final worksheet rows.

### `[isTable?]`

- **Purpose:** Selects the table-cleanup route instead of the QueryDef-creation route.
- **Typical type:** Yes/No.

#### When True

The sequence loop treats `QueryName` as a disposable table:

```vb
If TableOrQueryExists_TSB("", tmpString) = True Then
    DoCmd.DeleteObject acTable, tmpString
End If
```

Despite the helper function's broad name, the deletion command explicitly uses `acTable`.

#### When False

The row is passed to `clsCreateRecord`:

```vb
.prpName = Trim(rs1(fldQname))
.prpSqlString = Trim(rs1(fldString))
.prpExecute = rs1![Exec]
qOk = .CreateQuery
```

`clsCreateRecord` deletes and recreates the named QueryDef, then executes it only when `Exec = True`.

#### Practical interpretation

```text
isTable? = True
    -> this row identifies a disposable output/cache table

isTable? = False
    -> this row defines a QueryDef
```

Do not interpret `[isTable?]` as a restriction on what ADO can open during final transfer. It controls preparation and cleanup behavior.

### `SQL`

- **Purpose:** Stores the Access SQL assigned to a dynamically created QueryDef.
- **Typical uses:**
  - `SELECT ... INTO` make-table queries
  - Append queries
  - Update queries
  - Delete queries
  - Saved select queries when `Exec = False`

Use Access SQL syntax, not SQL Server syntax, unless the QueryDef is explicitly configured as a pass-through query elsewhere.

The SQL should end with a semicolon and should be tested directly in Access before being stored in the sequence table.

### `Spacing`

- **Purpose:** Associates a sorting/grouping source with a worksheet.
- **Typical value on a sorting-table row:** The worksheet's `Tab` value.
- **Example:** A row for `tblSorKPI_MonthlyReport` uses `Spacing = KPI_Monthly`.

The launcher builds a spacing-to-tab mapping by joining the spacing value to the worksheet tab value.

### `Spacing Order`

- **Purpose:** Controls the order of grouping sources associated with a worksheet.
- **Typical type:** Long Integer.
- **Use:** Required when a worksheet has nested or multiple grouping levels.
- **Guidance:** The primary sorting source should generally have the first spacing order.

### `Exec`

- **Purpose:** Tells `clsCreateRecord` whether to execute the newly created QueryDef.
- **Typical type:** Yes/No.

#### `Exec = True`

The QueryDef is created and executed with `dbFailOnError`. Use this for action SQL such as `SELECT ... INTO`.

#### `Exec = False`

The QueryDef is created and saved without execution. Use this for:

- Saved select queries
- Temporary worksheet/spacing views
- Queries that will be opened later as record sources

### `HasTotals`

- **Purpose:** Enables report-level grand-total behavior for the worksheet.
- **Typical type:** Yes/No.
- **Important:** Grand-total metadata must exist for the same report and tab when this feature is enabled.

This field does not replace ordinary F1 or F2 subtotals. It controls the additional report-level grand-total block.

### `Over_ride_head1`

- **Purpose:** Replaces the default number of primary-header rows for the worksheet.
- **Typical type:** Integer.
- **Safe behavior in the current Medicare renderer:** A null value uses the default allocation.
- **Important:** A custom value must accommodate every configured H1 position and merge.

### `Over_ride_foot1`

- **Purpose:** Replaces the default number of primary-footer rows for the worksheet.
- **Typical type:** Integer.
- **Safe behavior in the current Medicare renderer:** A null value uses the default allocation.
- **Important:** A custom value must accommodate F1 and any report-specific footer requirements.

### `Notes`

- **Purpose:** Human-readable implementation guidance.
- **Execution effect:** None confirmed in the current renderer.
- **Recommended content:**
  - Why a row exists
  - Required predecessor
  - Expected output
  - Sorting-table contract
  - Known special behavior

## 2. Execution Model

For the selected report, the launcher constructs an ordered recordset from `tblReportSequence`:

```sql
SELECT tblReportSequence.*
FROM tblReportSequence
WHERE tblReportSequence.Report = [selected report]
ORDER BY tblReportSequence.Sort;
```

The exporter then processes every row in order.

### 2.1 Table rows

When `[isTable?] = True`, the named table is deleted if it exists. This prepares the database for a later `SELECT ... INTO` action query.

### 2.2 QueryDef rows

When `[isTable?] = False`, the engine creates or replaces the named QueryDef from the stored SQL.

- `Exec = True` executes the QueryDef.
- `Exec = False` leaves the QueryDef available for later use.

### 2.3 Worksheet transfer

The tab recordset contains the final worksheet definitions. For each worksheet, the exporter:

1. Creates or selects the worksheet.
2. Reads `QueryName` from the tab row.
3. Opens that named source through ADO.
4. Copies the returned recordset beginning at cell `A2`.

### 2.4 Formatting handoff

After the raw records are copied, `clsExport` receives:

- The output filename
- The tab query
- The spacing query
- The spacing-to-tab mapping
- The metadata table name
- The generated sorting/grouping table name
- The render-visibility option

The formatter then calculates breaks, inserts rows, applies metadata, creates totals and charts, recalculates, and saves the workbook.

## 3. Standard Four-Row Worksheet Pattern

The public demos use a reliable four-row pattern for each worksheet.

### Row A: final output table marker

```text
QueryName:    FinalWorksheetTable
FinalOutput:  True
Tab:          WorksheetName
Tab Order:    1
isTable?:     True
Exec:         False
```

Purpose:

- Delete the old final-output table.
- Mark the rebuilt table as the worksheet source.

### Row B: final output action QueryDef

```text
QueryName:    qryBuildFinalWorksheetTable
isTable?:     False
SQL:          SELECT ... INTO FinalWorksheetTable ...
Exec:         True
```

Purpose:

- Create the QueryDef.
- Execute the QueryDef.
- Materialize the worksheet-source table.

### Row C: sorting-table marker

```text
QueryName:     SortingTable
isTable?:      True
Spacing:       WorksheetName
Spacing Order: 1
HasTotals:     True or False
```

Purpose:

- Delete the old sorting table.
- Associate the sorting source with the worksheet.

### Row D: sorting-table action QueryDef

```text
QueryName:    qryBuildSortingTable
isTable?:     False
SQL:          SELECT ... INTO SortingTable ...
Exec:         True
```

Purpose:

- Create and execute the QueryDef that materializes the sorting table.

## 4. Sorting-Table Requirement

The primary sorting table is a contract with `fn_FindBreaks`. Its first fields must provide:

1. An ordered grouping key or sequence value
2. A display description
3. A record count

A valid conceptual shape is:

```text
BandSequence    BandDescription    RecCount
1               Category A         25
2               Category B         18
3               Category C         31
```

The field names used by a specific version should match the implementation's expected headers. Do not rely only on visual similarity or data type.

## 5. KPI Worked Sequence

The KPI report contains two worksheet branches in one report definition. The supplied sequence contains eight ordered rows.

### 5.1 `KPI_Monthly` branch

#### Sort 1: final-output table marker

```text
Report:       Monthly_Report.XLSX
Sort:         1
QueryName:    KPI_Results_Monthly
FinalOutput:  True
Tab:          KPI_Monthly
Tab Order:    1
isTable?:     True
Exec:         False
HasTotals:    False
```

This row deletes the old `KPI_Results_Monthly` table and marks the rebuilt object as the first worksheet source.

#### Sort 2: build the final-output table

```text
QueryName:    qryKPI_Results_Monthly_2
isTable?:     False
Exec:         True
```

The SQL selects prepared Monthly KPI fields, calculates work-order and sample metrics, constructs `Year_Range`, and writes the result into `KPI_Results_Monthly`.

#### Sort 3: sorting-table marker

```text
QueryName:     tblSorKPI_MonthlyReport
isTable?:      True
Spacing:       KPI_Monthly
Spacing Order: 1
HasTotals:     True
```

This row deletes the old sorting table and associates the rebuilt sorting table with the `KPI_Monthly` worksheet.

#### Sort 4: build the sorting table

```text
QueryName:    qryInsertInto_SorKPI_MonthlyReport
isTable?:     False
Exec:         True
```

The action SQL writes three key values into `tblSorKPI_MonthlyReport`:

```text
SpecialSort
Month_Year
CountOfWorkOrdr
```

These represent the ordered group key, visible group description, and detail-record count.

### 5.2 `KPI_Special` branch

Sorts 5 through 8 repeat the same pattern for:

- Final output: `KPI_Results_Special`
- Worksheet: `KPI_Special`
- Sorting table: `tblSorKPI_SpecialReport`
- Build query: `qryInsertInto_SorKPI_SpeicalReport`

The two final-output rows use different `Tab Order` values, allowing one report definition to create and process both worksheets.

### 5.3 What the KPI sequence demonstrates

The KPI sequence proves that `tblReportSequence` can coordinate:

- Multiple final worksheet sources
- Multiple sorting tables
- Multiple action QueryDefs
- Worksheet-specific grouping associations
- Worksheet-specific totals
- One workbook generated from one selected report

## 6. Medicare Worked Sequence

The Medicare report uses the same four-row pattern.

### Sort 1: final-output table marker

```text
Report:       Medicare_Report.XLSX
Sort:         1
QueryName:    Medicare_ReportFinal
FinalOutput:  True
Tab:          Medicare_Report
Tab Order:    1
isTable?:     True
```

### Sort 2: build the final-output table

```text
QueryName:    qryMedicareReport_Final
isTable?:     False
Exec:         True
```

The action SQL selects the final report fields from `qryMedicare_ReportFinal`, orders the records by RBCS band, estimated Medicare payment, provider type, and place of service, and materializes `Medicare_ReportFinal`.

### Sort 3: sorting-table marker

```text
QueryName:     tblSort_MedicareFinal
isTable?:      True
Spacing:       Medicare_Report
Spacing Order: 1
HasTotals:     True
Over_ride_head1: 4
Over_ride_foot1: 3
```

This row associates the sorting table with the Medicare worksheet and applies custom primary header and footer allocations.

### Sort 4: build the sorting table

```text
QueryName:    qryInsertInto_tblSortMedicareFinal
isTable?:     False
Exec:         True
```

The SQL calculates a band sequence, retains `RBCSBandKey` as the band description/key, counts records per band, and writes the results into `tblSort_MedicareFinal`.

## 7. Direct Saved-Query Output

The code supports a final source that is a saved select QueryDef, even though the supplied demos materialize tables.

A direct-query configuration would conceptually use:

```text
QueryName:    qryFinalWorksheetSource
FinalOutput:  True
Tab:          WorksheetName
Tab Order:    1
isTable?:     False
SQL:          SELECT ... FROM ... ORDER BY ...;
Exec:         False
```

The sequence loop creates the saved QueryDef without executing it as an action query. Later, ADO opens `qryFinalWorksheetSource` and copies the returned rows into Excel.

### Conditions for success

- The query must return a normal recordset.
- All parameters must be resolved.
- The fields must match the metadata.
- The ordering must align with the sorting table.
- The sorting table must still satisfy the break-engine contract.

### Documentation status

This route is supported by the implementation but is not represented by the current public KPI or Medicare configurations. Materialized tables remain the demonstrated and recommended starting pattern.

## 8. Adding a New Worksheet to an Existing Report

Use this procedure after the report-ready source SQL has been tested.

### Step 1: define the final worksheet source

Decide whether to:

- Materialize a local table with an action QueryDef, or
- Create a saved select QueryDef with `Exec = False`

For a first implementation, use a materialized table.

### Step 2: define the sorting source

Create a query that produces:

1. Ordered group key
2. Display description
3. Record count

Verify the record counts sum to the number of rows in the final worksheet source.

### Step 3: add the final-output marker

Add a row with:

```text
FinalOutput = True
Tab = desired worksheet name
Tab Order = desired worksheet position
isTable? = True
```

Set `QueryName` to the final table name.

### Step 4: add the final-output build query

Add a later row with:

```text
QueryName = action QueryDef name
SQL = SELECT ... INTO final table ...
Exec = True
isTable? = False
```

### Step 5: add the sorting-table marker

Add a row with:

```text
QueryName = sorting table name
isTable? = True
Spacing = worksheet Tab value
Spacing Order = grouping level order
HasTotals = desired grand-total behavior
```

### Step 6: add the sorting-table build query

Add a later row with:

```text
QueryName = sorting-table action QueryDef name
SQL = SELECT ... INTO sorting table ...
Exec = True
isTable? = False
```

### Step 7: add metadata

Add `tblGoodFieldName_Report` rows for the same `Report` and `Tab` values.

If `HasTotals = True`, add matching `tblGoodFieldName_G_Totals` rows.

### Step 8: validate the sequence

Before rendering, verify:

- Table-marker rows precede their build-query rows.
- Final source and sorting source names match the SQL targets.
- `Tab` and `Spacing` values match exactly.
- `Tab Order` and `Spacing Order` are intentional.
- Overrides are null unless a custom allocation is required.

### Step 9: test with a small source

Use a small representative dataset, such as the 99-row Medicare test source, before running the full dataset.

## 9. Common Configuration Errors

### 9.1 Build query runs before cleanup marker

**Symptom:** `SELECT ... INTO` reports that the destination table already exists.

**Correction:** Place the `[isTable?] = True` marker earlier in `Sort` order.

### 9.2 `Exec` is False on an action query

**Symptom:** The QueryDef exists, but the destination table is never created.

**Correction:** Set `Exec = True` for make-table, append, update, and delete queries.

### 9.3 `Exec` is True on a saved select query

**Symptom:** DAO attempts to execute a record-returning query as an action.

**Correction:** Use `Exec = False` when the QueryDef is intended to be opened later as a select source.

### 9.4 `FinalOutput` source does not exist

**Symptom:** ADO cannot open the configured `QueryName` during worksheet transfer.

**Correction:** Verify that an earlier sequence row creates the source successfully.

### 9.5 `Spacing` does not match `Tab`

**Symptom:** The worksheet receives no grouping source or is not formatted correctly.

**Correction:** Use the worksheet's exact `Tab` value in the sorting-table row's `Spacing` field.

### 9.6 Sorting-table counts do not match the final source

**Symptom:** Physical report bands do not align with the copied records.

**Correction:** Reconcile sorting-table counts to the final ordered recordset.

### 9.7 Sorting-table fields are in the wrong order

**Symptom:** `fn_FindBreaks` fails or calculates incorrect coordinates.

**Correction:** Restore the required leading sorting-table contract.

### 9.8 Incorrect `Tab Order`

**Symptom:** Worksheets are created or processed in an unexpected order.

**Correction:** Assign unique worksheet-order values to final-output rows.

### 9.9 Unsafe overrides

**Symptom:** Header or footer content overlaps or disappears.

**Correction:** Leave overrides null to use defaults, or allocate enough rows for every configured position.

### 9.10 SQL dialect mismatch

**Symptom:** The dynamically created QueryDef fails to parse.

**Correction:** Use Access SQL syntax in `tblReportSequence`. Do not paste SQL Server-only syntax into a native Access QueryDef.

## 10. Validation Checklist

Before running a report, confirm:

- [ ] Every required row uses the same `Report` value.
- [ ] `Sort` values reflect the intended execution order.
- [ ] Cleanup rows precede make-table queries.
- [ ] Action queries use `Exec = True`.
- [ ] Saved select queries use `Exec = False`.
- [ ] Final-output names match created table or query names.
- [ ] `FinalOutput` rows contain valid `Tab` values.
- [ ] `Tab Order` values are unique and intentional.
- [ ] Sorting-table rows use `Spacing` values matching worksheet `Tab` values.
- [ ] `Spacing Order` values reflect grouping hierarchy.
- [ ] Sorting tables expose the required first fields.
- [ ] Sorting counts reconcile to final-source rows.
- [ ] `HasTotals` has matching grand-total metadata.
- [ ] Header and footer overrides are null or large enough.
- [ ] Stored SQL uses Access SQL syntax.
- [ ] Stored SQL has been tested independently in Access.

## 11. Minimal Example

A minimal materialized-table configuration can be represented as:

```text
Sort 1
    QueryName:    tblFinalOutput
    FinalOutput:  True
    Tab:          ReportSheet
    Tab Order:    1
    isTable?:     True

Sort 2
    QueryName:    qryBuildFinalOutput
    SQL:          SELECT ... INTO tblFinalOutput FROM ... ORDER BY ...;
    Exec:         True
    isTable?:     False

Sort 3
    QueryName:     tblSortFinalOutput
    Spacing:       ReportSheet
    Spacing Order: 1
    isTable?:      True
    HasTotals:     True

Sort 4
    QueryName:    qryBuildSortFinalOutput
    SQL:          SELECT GroupKey, GroupLabel, Count(*) AS RecCount
                  INTO tblSortFinalOutput
                  FROM tblFinalOutput
                  GROUP BY GroupKey, GroupLabel
                  ORDER BY GroupKey;
    Exec:         True
    isTable?:     False
```

This four-row pattern is the safest starting point for a new worksheet.

## 12. Related Documentation

- [Project README](../README.md)
- [Architectural White Paper](WhitePaper.md)
- [Metadata Configuration Guide](MetadataConfigurationGuide.md)
