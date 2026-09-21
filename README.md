# README.md — MetadataDrivenReportingEngine
A fully synthetic, metadata‑driven reporting engine demonstrating the architectural patterns I use in enterprise reporting systems. All tables, fields, logic, and workflows in this demo are fictional and do not correspond to any real system.

## 📌 Overview
The MetadataDrivenReportingEngine is a lightweight hybrid reporting framework built using Microsoft Access (VBA/DAO/ADO) and SQL Server. It demonstrates how operational teams can structure reporting workflows, apply metadata‑driven transformations, and generate KPI dashboards and Excel reports.

This demo intentionally uses sample data only.
No live SQL queries are executed.

The Report Manager form displays pre‑loaded sample data and a locked date range. The Get Data button simply shows an instructional message explaining that this is a demo. The actual reporting workflow is launched through the chart icon, which opens the KPI launcher form.

## Demo Files:

Download the demo files from the /Demo folder:

ReportKPI.accdb

DemoKPI_ExcelWriter.xlsx

## Windows Security Note:

Windows may block the Access file because it contains VBA code.
To unblock:
Right‑click the downloaded file
Select Properties
Check Unblock
Click OK
This is a standard Windows security feature.

## This project showcases:
metadata‑driven report definitions
hybrid Access + SQL workflow patterns
operational KPI extraction
Excel automation via VBA
separation of data retrieval and visualization

## 🧩 Architecture summary
### 1. Report Manager (frmReportManager)
This form is a static preview of a reporting UI. It displays:
A locked date range (non‑interactive)
A grid of pre‑loaded sample data
A Get Data button that shows the instructional message:

“Done retrieving data, please review your results. (This is sample data, no data actually retrieved. Use this button to run your custom query logic.)”

A chart icon that launches the reporting workflow
The Report Manager does not retrieve or process data.
The real reporting engine begins when the user clicks the chart icon, which opens frmReportLauncher.

### 2. Charting module (frmReportLauncher)
This is the functional core of the demo.
The launcher form provides:
A dropdown to select a report definition
A checkbox labeled Render Sheet?
A button to generate Excel output

Render Sheet? behavior:
Unchecked:
The engine builds the Excel workbook without rendering the sheet to the screen.
Typical runtime: around 2 minutes (varies by machine).

Checked:
The engine renders the sheet visibly while building it.
Typical runtime: around 5 minutes (varies by machine).
Useful for demonstration or debugging.

The charting module uses metadata stored inside ReportKPI.accdb to produce:

KPI charts
Excel dashboards
summary sheets

This is where the reporting engine actually runs.

### 3. Metadata tables (inside Access)
All metadata lives inside ReportKPI.accdb, including:

report definitions
column mappings
KPI formulas
chart configuration
sheet layout metadata

No external metadata files are required.

### 4. VBA modules (sanitized)
All modules are exported as .txt files for easy viewing.
These modules contain generic workflow logic and do not reference any real systems.

Included modules:

basBrowse.txt — file browsing utilities

clsCreateRecord.txt — record creation and querydef lifecycle

clsExport.txt — Excel export and formatting logic

Connections.txt — ADO/DAO connection helpers

modOpenFile.txt — file opening utilities

modMousePointer.txt — UI pointer state management

modExportToExcel.txt — Excel automation pipeline

modCreateQueryDef.txt — querydef creation and cleanup

mod_IsReportOpen.txt — report/form state checks

These modules demonstrate clean separation of responsibilities and a metadata‑driven workflow.


## 📁 Repository structure
```
MetadataDrivenReportingEngine/
│
├── README.md
│
├── Screenshots/
│   ├── ReportManager.png
│   ├── KPIChartingForm.png
│
├── VBA/
│   ├── basBrowse.txt
│   ├── clsCreateRecord.txt
│   ├── clsExport.txt
│   ├── Connections.txt
│   ├── modOpenFile.txt
│   ├── modMousePointer.txt
│   ├── modExportToExcel.txt
│   ├── modCreateQueryDef.txt
│   ├── mod_IsReportOpen.txt
│
└── Demo/
    ├── DemoKPI_ExcelWriter.xlsx
    └── ReportKPI.accdb
```

### 🚀 How to run the demo
Download the repository file ReportKPI.accdb
Open ReportKPI.accdb in Microsoft Access.
Open frmReportManager.
Review the pre‑loaded sample data and locked date range.
Click Get Data to see the instructional message.
Click the chart icon to open frmReportLauncher.
Select a report from the dropdown.
(Optional) Check Render Sheet? if you want to watch the sheet being drawn.
Click Create Spreadsheets to generate Excel output.
The resulting Excel file will resemble DemoKPI_ExcelWriter.xlsx.

This demo is designed to show the workflow and architecture, not to retrieve real data.

### 👤 About the author
Alex Fritts is a workflow engineer specializing in:

Access/VBA/COM automation
SQL Server hybrid architectures
metadata‑driven reporting systems
operational KPI pipelines
Excel automation and reporting engines

He builds tools that help analysts, operators, and technical teams work faster and more reliably.

### 🏭 Summary
This demo showcases a metadata‑driven reporting engine similar to the architecture used in enterprise environments. It demonstrates how Access, SQL Server, and Excel can work together to produce parameterized KPI dashboards, structured reporting workflows, and automated spreadsheet generation. All data, tables, and logic are fully synthetic and safe for public review.

### 🎯 Purpose
This repository is intended as a professional demonstration of:

SQL + Access hybrid engineering
metadata‑driven reporting
workflow automation
enterprise reporting patterns

It is not a production system and contains no proprietary data.
