# Troubleshooting Guide

[Back to project README](../README.md) | [Getting Started](GettingStarted.md) | [Metadata Configuration Guide](MetadataConfigurationGuide.md) | [`tblReportSequence` Reference](TblReportSequenceReference.md) | [Medicare Pipeline](WorkedExamples/MedicarePipeline.md) | [KPI Pipeline](WorkedExamples/KPIPipeline.md)

## Purpose

This guide helps diagnose failures in the Metadata-Driven Excel Reporting Engine. The engine combines Access SQL, DAO QueryDefs, ADO recordsets, Excel COM automation, calculated report bands, and executable metadata. A failure visible in Excel may originate much earlier in the pipeline.

The most important troubleshooting principle is:

> Do not assume that every runtime error indicates a defect in the VBA engine. The metadata, final worksheet source, sorting table, sequence order, and generated Excel formulas are all part of the program.

## 1. Troubleshoot in Pipeline Order

Always diagnose the earliest failing stage.

```text
1. Prepared source data
2. Final worksheet source
3. Sorting table
4. tblReportSequence
5. ADO transfer
6. tbl_temp_sorting_grouping
7. Metadata rendering
8. Excel formulas
9. Grand totals
10. Charts and saved workbook
```

If the final output table is wrong, editing Excel metadata will not fix it. If the sorting-table counts are wrong, debugging borders and formulas is premature.

## 2. Use a Small Test Dataset First

The Medicare demo includes a 99-row testing strategy. Use a similarly small, representative source for new reports.

A useful test source should include:

- Several report bands
- Different band sizes
- Zero and nonzero numeric values
- Nulls where nulls are valid
- Rows that exercise conditional formulas
- Enough records to create multiple inserted headers and footers

Recommended loop:

```text
Prepare test source
    -> build final output
    -> build sorting table
    -> render workbook
    -> inspect every band
    -> change one item
    -> render again
```

If the report fails with 99 rows, investigate configuration, field mapping, formulas, and band logic. If it works with 99 rows but fails with the full source, investigate volume, numeric conversion, Excel row positions, and performance.

## 3. Preserve the Last Working Configuration

Before changing metadata:

1. Copy the working report and tab rows.
2. Give the copy a new report or tab identifier.
3. Change one property group at a time.
4. Render after each meaningful change.
5. Keep the last working copy until the new layout is stable.

Begin with a working layout from `tblGoodFieldName_Report` rather than creating every field from blank records.

## 4. Errors Before Excel Opens

### 4.1 Report is missing from the launcher

Check:

- A matching report row exists in `tblReports`.
- The report name exactly matches `tblReportSequence.Report`.
- Leading or trailing spaces are not present.
- The launcher row source includes the new report.

### 4.2 Nothing happens after selecting the report

Check:

- The report has sequence rows.
- `Sort` values are populated and ordered.
- `QueryName` is populated on every active row.
- The launcher is passing the expected sequence SQL to `createSpreadsheetReport`.
- The shared ADO connection is open.

### 4.3 Output path is rejected

Check:

- A valid filename was selected.
- The destination folder exists.
- The workbook is not open or locked.
- The user has write permission to the destination.
- The selected filename has an Excel-compatible extension.

### 4.4 Existing workbook cannot be overwritten

Close the workbook in Excel and try again. Also check for a background Excel process that still has the file open.

## 5. `tblReportSequence` Errors

### 5.1 Destination table already exists

Likely causes:

- The table-marker row is missing.
- `[isTable?]` is not enabled on the marker row.
- The marker row appears after the make-table query.
- `QueryName` does not exactly match the destination table in the SQL.

Correct order:

```text
Table marker
    -> action QueryDef that recreates the table
```

### 5.2 QueryDef exists but the destination table is missing

Check `Exec` on the action-query row.

```text
Exec = True
```

A make-table, append, update, or delete QueryDef must be executed to change data.

### 5.3 QueryDef is recreated but not executed

This is expected when `Exec = False`. Use that setting for a saved select QueryDef, not for an action query that must materialize a table.

### 5.4 Wrong object is deleted

`[isTable?]` controls the object lifecycle branch:

```text
True
    -> delete an Access table with DoCmd.DeleteObject acTable

False
    -> create or replace a QueryDef through clsCreateRecord
```

Do not mark a QueryDef as a table merely because the QueryDef creates a table when executed.

### 5.5 Worksheet is missing

Check:

- `FinalOutput = True` on the final worksheet source row.
- `Tab` is populated.
- `Tab Order` is populated.
- The final source exists after the sequence executes.
- The tab query used by the launcher returns the row.

### 5.6 Wrong worksheet order

Check `Tab Order` on final-output rows. Do not rely on sequence `Sort` alone to control workbook tab order.

### 5.7 Sorting table is not associated with the worksheet

The sorting-table marker's `Spacing` value must exactly match the worksheet row's `Tab` value.

```text
Tab:      KPI_Monthly
Spacing:  KPI_Monthly
```

Check spelling, spaces, capitalization, and underscores.

### 5.8 Grand totals do not appear

Check:

- `HasTotals = True` on the sorting-table marker row.
- Matching rows exist in `tblGoodFieldName_G_Totals`.
- The metadata `Report` and `Tab` values match the active worksheet.
- The grand-total formulas are valid.

## 6. Final-Source and ADO Errors

### 6.1 ADO cannot find the final source

Check:

- The table or saved select query exists after sequence execution.
- The final-output `QueryName` is spelled correctly.
- The source does not require unresolved parameters.
- A form reference used by the query is available.
- The ADO connection points to the current database.

### 6.2 ADO opens the source but returns no rows

Run the object directly in Access. Verify:

- Filter values match actual source values.
- Date criteria are valid.
- Join fields use compatible data types.
- Required source tables have records.
- Application-specific preparation ran before the launcher.

In the KPI demo, the Monthly category is stored as `Monthy`. Changing only the query or only the data will produce no matches.

### 6.3 Worksheet columns do not match metadata

The raw ADO recordset is copied in field order. Compare:

```text
Final source field order
    vs.
Metadata Sort and Letter order
```

A field added in the middle of the final source can shift every later worksheet column.

### 6.4 Final source sorting is unstable

Add explicit sort fields and a complete `ORDER BY` clause. Grouping data must be contiguous before the renderer inserts report bands.

Avoid relying on display labels alone when numeric or normalized sort fields are needed.

## 7. Sorting-Table Errors

### 7.1 Sorting-table counts do not reconcile

Verify:

```text
Sum of sorting-table record counts
    = number of records in the final worksheet source
```

If the values differ, investigate:

- Different filters
- Null group keys
- Noncontiguous sort order
- Different grouping expressions
- Stale destination tables
- Duplicated rows

### 7.2 Bands begin or end in the wrong place

Likely causes:

- Sorting-table counts are wrong.
- Final-source records are not ordered by the grouping key.
- Sorting-table group order differs from final-source group order.
- The wrong sorting table is associated with the worksheet.

### 7.3 `fn_FindBreaks` fails immediately

Check the sorting-table contract. The first fields must provide:

1. A sortable sequence or key
2. A text group description
3. A record count

Do not insert unrelated fields before these expected values.

### 7.4 Group description is blank

Check the second sorting-table field. It should contain the text the break engine uses for the band description.

### 7.5 Record count is Null, text, or zero

The count field should be numeric. A zero-row group should normally be excluded from the sorting table rather than passed to the break engine.

### 7.6 KPI fails only on a very large worksheet

The older KPI `clsExport` uses `CLng()` when working with generated row positions. If a small test works and a large report fails during formatting, inspect numeric conversions and calculated worksheet coordinates.

Treat Medicare as the current renderer baseline when evaluating a correction.

## 8. `tbl_temp_sorting_grouping` Errors

This table is generated output. Do not normally edit it manually.

### 8.1 Calculated positions are Null or zero

Investigate upstream values:

- Sorting-table count
- Header allocation
- Footer allocation
- Original starting row
- Worksheet starting row
- Previous band coordinates

### 8.2 Calculated ranges overlap

Check:

- `Begin1 <= End1`
- Header rows end before detail rows begin
- Footer rows begin after detail rows end
- The next band's header begins after the preceding footer
- Override values are not too small

### 8.3 Generated table contains stale rows

Confirm the pipeline clears `tbl_temp_sorting_grouping` before recalculation and that the current report's sorting tables are written to `tblNameHolder`.

## 9. Metadata Errors

### 9.1 General rule for Null and zero

Null and zero are not interchangeable.

- Null may mean disabled, optional, or use the default.
- Zero may be treated as a real numeric instruction.
- Some enabled features require positive numeric values.

Do not replace a working value with Null or zero unless the code path is known to support it.

### 9.2 Runtime error after a metadata change

Compare the first failing version with the last working version. Check these fields first:

- `Sort`
- `Letter`
- `ColumnWidth`
- `Position_H1_A`
- `Position_F2_A`
- `Font_Size_*`
- `Row_Height_*`
- Formula and merge-range fields
- Alignment, color, format, and line-weight lookup values

### 9.3 Metadata row is ignored

Check:

- `Report` exactly matches the active report.
- `Tab` exactly matches the worksheet.
- The row is included by the renderer's metadata query.
- `Sort` is populated.
- The feature's enable flag is selected.
- The field is actually implemented in the current renderer.

### 9.4 Section-wide setting has no effect

Many shared H1, F2, fill, merge, and boundary settings are read from the first metadata record in `Sort` order.

Move the shared setting to the first sorted record and render again.

### 9.5 First sorted row is hidden

The engine determines shared-range boundaries from visible columns but may read shared properties after returning to the first metadata record. Keep the first sorted record visible whenever practical.

### 9.6 Later metadata row does not override H1 or F2 formatting

Do not assume section-wide fields work like per-column fields. Shared section behavior often uses a single authoritative row.

### 9.7 Metadata field exists but does nothing

The current reference renderer does not confirm implementation for fields including:

- `PasteSpecial`
- `PasswordProtect`
- `Password`
- `List_B`
- `ErrorTitle_B`
- `ErrorMessage_B`
- `Lock_B`

Treat these as reserved, legacy, or unimplemented until the code explicitly consumes them.

## 10. Header and Footer Allocation Errors

### 10.1 Blank overrides were changed to zero

Leave `Over_ride_head1` and `Over_ride_foot1` blank to use the engine defaults.

Do not enter zero to mean default.

### 10.2 Custom header formatting overlaps detail rows

When overriding primary allocations, use practical minimums of:

```text
Over_ride_head1 >= 4
Over_ride_foot1 >= 3
```

Values below those tested limits can reduce structural space until formatting ranges collide.

### 10.3 Header item is missing

Check:

- `HeaderCol_H1_A = True`
- `Position_H1_A` fits within the allocated rows
- `FieldName_H1_A` is populated
- `Add_Name_From_Data_H1_A` points to a valid Excel column when data is copied
- The source value exists at the row used by the header routine

### 10.4 F2 item is missing

Check:

- `HeaderCol_F2_A = True`
- `Position_F2_A` fits within the footer allocation
- `Add_Name_From_Data_F2_A` points to the correct helper or source column
- `GetFromSum_F2_A` uses the intended source row

### 10.5 Worksheet-level header does not appear

Review the constants in `fn_FindBreaks`:

```vb
Const original = 2
Const beginSheet = 10
```

The Medicare reference implementation uses `beginSheet = 10` to reserve enough space for the sheet header. Lower values may not provide enough room for the header configuration.

## 11. Formula Errors

### 11.1 Formula begins with `==`

Metadata formulas should normally omit the leading equals sign because the VBA supplies it.

Store:

```text
SUM(K@:K#)
```

not:

```text
=SUM(K@:K#)
```

### 11.2 Formula still contains `?`, `@`, or `#`

Check that the formula was entered in the correct metadata field for the zone being processed. Token meaning depends on the code path.

### 11.3 Formula points to the wrong rows

Check:

- `Begin1`
- `End1`
- Footer row coordinates
- Formula field used
- Absolute and relative references
- Whether inserted rows changed the expected context

### 11.4 Formula fills only one detail row

Verify the renderer assigns the formula to the first detail cell and fills through the intended B_D range.

### 11.5 Percentage formula returns `#DIV/0!`

Add an Excel guard such as `IFERROR` or an explicit denominator test.

Example:

```text
IFERROR(M?/L?,0)
```

### 11.6 Running count double-counts totals

Verify the `SUMIFS` criterion range contains the subtotal label and the criterion excludes `Totals:`.

Representative pattern:

```text
SUMIFS($K$1:INDEX(K:K,ROW()-1),$J$1:INDEX(J:J,ROW()-1),"<>Totals:")
```

Also verify that the criterion and sum columns match the actual worksheet configuration.

### 11.7 Running count resets unexpectedly

Check for relative references that moved during fill operations. Use absolute column anchors and confirm whether the formula is meant to accumulate across the whole worksheet or restart at each band.

## 12. Hidden Helper Column Errors

### 12.1 Visible footer shows blank helper result

Temporarily unhide the helper column and check:

- The formula exists in every required row.
- The formula references the intended columns.
- The source cells are numeric where required.
- The F2 source column points to the helper column.
- The selected source row contains the completed result.

### 12.2 Helper formula is overwritten

Check that the helper column does not conflict with a field copied from the ADO recordset and that no later metadata row writes to the same Excel column.

### 12.3 Hidden column does not hide

Check `Column_Hide` on the correct report, tab, and letter row. Do not use a zero width as a substitute unless the code explicitly supports that behavior.

## 13. Formatting Errors

### 13.1 Invalid alignment, color, style, or line weight

Use values already present in a working metadata layout. An intuitive synonym may not match the renderer's `Select Case` logic.

### 13.2 Column width or row height causes an error

Confirm the value is numeric and positive. Avoid Null or zero when the feature is enabled and the code expects a numeric assignment.

### 13.3 Merge operation fails

Check:

- The range is valid.
- The range does not cross an incompatible previously merged area.
- The row token is replaced correctly.
- The first and last visible columns create the expected shared range.

### 13.4 Borders appear on the wrong columns

Inspect the first and last visible metadata rows and the first sorted metadata record. Shared boundary settings are not necessarily read from every column row.

## 14. Grand-Total Errors

### 14.1 Grand-total block is missing

Check `HasTotals`, matching grand-total metadata, and report/tab identifiers.

### 14.2 Grand-total formulas reference the wrong bands

Inspect the final generated rows in `tbl_temp_sorting_grouping` and verify the grand-total formula tokens resolve after all standard bands are created.

### 14.3 Grand totals are doubled

Check whether the formula sums detail rows and subtotal rows together. Exclude inserted subtotal labels or sum only the intended detail columns.

## 15. Chart Errors

### 15.1 Chart sheet is missing

Check that the chart routine is called for the active report and tab after worksheet formatting.

### 15.2 Chart is empty

Verify:

- Source ranges exist.
- Formula cells have calculated values.
- Helper columns were not deleted.
- The chart routine runs after formulas and totals are created.

### 15.3 Chart contains stale values

Ensure calculation is restored and completed before saving. The Medicare reference renderer performs a final workbook calculation and full rebuild before save.

### 15.4 KPI chart fails only with a large source

First test the report worksheets without chart creation. If the worksheet itself fails, inspect row-position conversions. If the worksheets succeed but charting fails, inspect the completed chart source ranges.

## 16. Excel Automation Errors

### 16.1 Excel remains running after an error

Use Task Manager to close orphaned Excel processes before testing again. During development, ensure error handlers close recordsets, workbooks, and the Excel application.

### 16.2 Workbook is saved but invisible

Check the render and visibility flags. The engine can build with screen updating disabled.

### 16.3 Workbook opens with stale formulas

Confirm the final renderer:

1. Restores automatic calculation.
2. Calculates the workbook.
3. Performs a full calculation rebuild when required.
4. Saves after calculation completes.

### 16.4 Different code branches place raw data on different rows

Review both the overwrite and new-file branches in `modExportToExcel`. Confirm both branches use the same initial destination row expected by `fn_FindBreaks`.

## 17. Naming and Legacy-Object Errors

The demos contain several inherited or inconsistent names. Do not "correct" only one side of a working dependency.

Examples include:

- `qryKPI_Results`, which is a local table
- `Monthy`, used by the KPI Monthly filter
- `qryInsertInto_SorKPI_SpeicalReport`, with `Speical` in the name
- Similar-looking sorting-table names with and without underscores
- `RBCSBandLable` in older metadata versus `RBCSBandLabel` in the current SQL pipeline

Before renaming:

1. Search queries.
2. Search form code.
3. Search modules and classes.
4. Search `tblReportSequence`.
5. Search metadata tables.
6. Update every dependency together.

## 18. Debugging Checklist

### Access pipeline

- [ ] Source tables contain expected records.
- [ ] Saved select queries return expected rows.
- [ ] Action queries run manually without prompts.
- [ ] Destination tables are rebuilt successfully.
- [ ] Final field order matches Excel metadata.

### Sequence

- [ ] `Sort` values are sequential and correct.
- [ ] Table markers precede build queries.
- [ ] Action-query rows use `Exec = True`.
- [ ] Final-output rows use `FinalOutput = True`.
- [ ] `Tab` and `Spacing` match exactly.
- [ ] `Tab Order` and `Spacing Order` are correct.

### Sorting

- [ ] First three sorting fields satisfy the contract.
- [ ] Group keys are non-null.
- [ ] Counts are numeric and positive.
- [ ] Counts reconcile to the final source.
- [ ] Group order matches final-source order.

### Metadata

- [ ] Report and tab identifiers match.
- [ ] First sorted metadata row is authoritative for shared fields.
- [ ] Column letters are valid and unique where required.
- [ ] Enabled numeric settings are not Null or zero unexpectedly.
- [ ] Formula templates omit the leading equals sign.
- [ ] Source-column references exist.
- [ ] Reserved fields are not assumed to be active.

### Excel

- [ ] Raw records appear on the expected starting row.
- [ ] Bands do not overlap.
- [ ] Hidden helpers calculate correctly.
- [ ] Subtotals and grand totals reconcile.
- [ ] Charts use completed ranges.
- [ ] Reopened workbook values are current.

## 19. Recommended Diagnostic Outputs

During development, preserve these artifacts for each test run:

- Final worksheet source row count
- Sorting-table row count
- Sum of sorting-table counts
- Current `tblReportSequence` rows
- Current metadata rows for the report and tab
- Generated `tbl_temp_sorting_grouping`
- Immediate Window output from the sequence and export routines
- The small generated workbook

Comparing these artifacts between the last working run and the first failing run usually identifies the changed stage.

## 20. Related Documentation

- [Project README](../README.md)
- [Getting Started](GettingStarted.md)
- [Architectural White Paper](WhitePaper.md)
- [Metadata Configuration Guide](MetadataConfigurationGuide.md)
- [`tblReportSequence` Reference](TblReportSequenceReference.md)
- [Medicare Pipeline](WorkedExamples/MedicarePipeline.md)
- [KPI Pipeline](WorkedExamples/KPIPipeline.md)
