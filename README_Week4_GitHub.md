# Week 4 – Excel Automation and VBA Macro Development

This repository contains a beginner-friendly example project for automating repetitive sales-data preparation and analysis in Microsoft Excel.

## Files
- `Week4_SalesAutomation.bas` — VBA module with three macros:
  - `CleanSalesData`: trims text spaces and applies basic number/date formatting.
  - `BuildRegionSummary`: totals revenue by region and creates/refreshes a summary sheet.
  - `RefreshAnalysis`: refreshes workbook data/pivots, then runs the preparation and summary routines.
- `YuvaIntern_Week_4_Automation_Macro_Development_Report.docx` — detailed report with objective, pseudocode, design, testing, benefits, risks, debugging, and scalability.

## Assumptions
The source sheet is named `Sales Data`. Row 1 contains headers including `Region` and `Revenue` (or `Sales`). The macros search for the Region and Revenue headers for the regional summary. The cleaning macro formats columns named `Date` and `Revenue` if present. Always check the column names and use a backup copy first.

## Importing the module
1. Open a copy of the workbook in desktop Excel.
2. Save as **Excel Macro-Enabled Workbook (`.xlsm`)**.
3. Press `Alt + F11` to open the Visual Basic Editor.
4. Choose **File → Import File...** and select `Week4_SalesAutomation.bas`.
5. Return to Excel, press `Alt + F8`, choose a macro, and click **Run**.
6. Enable macros only for files you trust. Review the output before using it for decisions.

## Important
This is an educational example. It has not been run against every possible workbook layout. Test on a copy, verify the column headers, and compare summary totals with a manual calculation before relying on results.
