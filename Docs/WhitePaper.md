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

## 4. GGrouping & Sorting Tables
Grouping metadata is stored in:

tblNameHolder — group names, descriptions, and order

spacing tables — define Begin/End positions for each group

These tables determine:

grouping hierarchy

sort order

spacing rules

header/footer placement

nested group behavior

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

## 7. Metadata Formatting Tables
Formatting is controlled by three metadata tables:

Header (tblGoodFieldName_Header)
Column headings

Visibility

Line weights

Bold/italic

Alignment

Background color

Body (tblGoodFieldName_Report)
Detail row formatting

Numeric formats

Text formats

Conditional formatting

Column widths

Visibility

Footer / Summary (tblGoodFieldName_G_Totals)
Summary rows

Extended summaries

Grand totals

Line weights

Spacing rules

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



