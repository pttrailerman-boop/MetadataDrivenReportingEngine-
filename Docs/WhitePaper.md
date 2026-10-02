# Metadata‑Driven Reporting Engine — Architectural White Paper
## Introduction
The MetadataDrivenReportingEngine is a synthetic demonstration of a metadata‑driven reporting architecture implemented using Microsoft Access (VBA/DAO/ADO) and SQL Server. It illustrates how reporting systems can be designed to minimize hard‑coded logic, maximize configurability, and support scalable workflows through metadata tables that define structure, formatting, grouping, and rendering behavior.

This white paper describes the full reporting pipeline, beginning with SQL orchestration in tblReportSequence, continuing through grouping and break‑table construction, and concluding with metadata‑driven Excel rendering.

## 2. Pipeline Overview
The reporting engine follows a deterministic, metadata‑driven pipeline:

```
Code
+-----------------------+
| tblReportSequence     |  <-- SQL pipeline + workflow orchestration
+-----------+-----------+
            |
            v
+-----------------------+
| Grouping Tables       |  <-- tblNameHolder + spacing tables
+-----------+-----------+
            |
            v
+-----------------------+
| Recursive Break Engine|  <-- fn_FindBreaks
+-----------+-----------+
            |
            v
+-----------------------+
| Break Table           |  <-- tbl_Temp_sorting_grouping
+-----------+-----------+
            |
            v
+-----------------------+
| Metadata Formatting   |  <-- Header/Body/Footer tables
+-----------+-----------+
            |
            v
+-----------------------+
| Excel Rendering       |  <-- fn_applyFormat
+-----------------------+
```
Each stage is driven by metadata rather than hard‑coded logic.

## 3. tblReportSequence — SQL Pipeline & Workflow Orchestrator
The reporting pipeline begins with tblReportSequence, which defines:

ReportName — logical identifier

TabOrder — ordering of report tabs

QueryName — source query or SQL

SQL_PreProcess — optional SQL executed before rendering

SpacingTable — grouping metadata

HasTotals — whether summary rows are required

Header/Footer overrides

Notes — human‑readable instructions

This table acts as the report manifest, describing how each report tab should be processed.

This table acts as the report manifest, describing how each report tab should be processed.

## 4. Grouping & Sorting Tables
Grouping metadata is stored in:

tblNameHolder — group names, descriptions, and order

spacing tables — define Begin/End positions for each group

These tables determine:

grouping hierarchy

sort order

spacing rules

header/footer placement

nested group behavior

## 5.Recursive Line Break Engine (fn_FindBreaks)
The recursive break engine is responsible for constructing the break table that drives row placement in Excel.

Responsibilities:
Read grouping metadata

Calculate Begin/End positions

Insert header/footer spacing

Handle nested groups

Build tbl_Temp_sorting_grouping

ASCII Diagram
```
Code
Group A
 ├── Header
 ├── Detail Rows
 ├── Subgroup A1
 │     ├── Header
 │     ├── Detail Rows
 │     └── Footer
 └── Footer
```

## 6. Row Insertion (fn_break)
The break table contains Insert_1 values that determine where rows should be inserted in Excel.

Responsibilities:
Insert rows at correct positions

Apply spacing rules

Maintain group hierarchy

## 7. Metadata Formatting and Configuration
 
Excel layout and formatting are controlled primarily through two metadata tables:
 
### tblGoodFieldName_Report
 
`tblGoodFieldName_Report` defines the worksheet-level presentation rules for each report and tab. Each record generally represents an Excel column and contains the metadata used to format and populate that column.
 
The table controls:
 
- Excel column letters, widths, visibility, and processing order
- Primary group headers
- Column headings
- Detail-row formatting
- Primary subtotal rows
- Secondary and nested footer panels
- Fonts, styles, sizes, colors, and underlining
- Horizontal and vertical alignment
- Number and date formats
- Cell background colors
- Row heights
- Borders and line weights
- Merged cells
- Static labels and values copied from report data
- Metadata-driven Excel formulas
- Hidden helper columns used for intermediate calculations
 
The table divides each report block into several logical zones:
 
- **H1**: Primary group header
- **B**: Body or column-heading row
- **B_D**: Detail-data region
- **F1**: Primary subtotal or summary row
- **F2**: Secondary or nested footer panel
 
Fields sharing these suffixes control the corresponding worksheet zone. For example, `Font_Color_H1_A` affects the primary header, while `Font_Color_F2_A` affects the secondary footer panel.
 
### tblGoodFieldName_G_Totals
 
`tblGoodFieldName_G_Totals` contains metadata used for grand-total presentation and calculations.
 
This table controls report-level summary behavior that occurs after individual groups and nested blocks have been processed. Depending on the report configuration, it may define:
 
- Grand-total labels
- Grand-total formulas
- Summary formatting
- Number formats
- Fonts and colors
- Borders and line weights
- Final report-level calculations
 
Grand totals are separate from ordinary F1 and F2 subtotals because they summarize the completed report rather than an individual grouping block.
 
### Formula Placeholders
 
Some metadata fields accept Excel formula expressions containing placeholders that the rendering engine replaces with calculated worksheet row numbers.
 
The current implementation uses placeholders such as:
 
- `@` for a calculated beginning row
- `#` for a calculated ending row
- `?` for a context-specific row supplied by the formatting routine
 
This allows formulas to adjust automatically when the engine inserts headers, footers, nested groups, and subtotal rows.
``` 
For example:

SUM(J@:J#)

If a detail block begins at row 18 and ends at row 35, the engine converts the expression to:

SUM(J18:J35)
```
For complete field definitions, accepted values, formula tokens, worksheet-zone diagrams, and examples, see the [Metadata Configuration Guide](MetadataConfigurationGuide.md).

## 8. Excel Rendering Engine (fn_applyFormat)
The rendering engine applies metadata to produce the final Excel workbook.

Responsibilities:
Render headers

Render detail rows

Render footers

Apply formatting

Insert spacing

Add grand totals

Finalize workbook

## 9. Example Walkthrough
A full example demonstrates:

SQL pre‑processing

grouping table

break table

metadata tables

final Excel output

(This section can be expanded later.)

## 10. Why Metadata‑Driven Reporting Scales
Metadata‑driven design allows:

adding new reports without code

changing formatting without code

adding new grouping logic without code

(This section can also be expanded later.)



