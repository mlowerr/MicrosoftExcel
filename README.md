# MicrosoftExcel Workbook Utilities

This repository contains example Excel formulas and VBA macros for managing data inside workbooks. Each file captures a focused snippet you can adapt for your own spreadsheets, including practical context on what you get and why you would use it.

## Files

- **AGENTS.md** — Project context and maintenance instructions for future agents.
- **FilterFunction-LimitAndReorderReturnedColumns.txt** — Demonstrates an Excel `FILTER` formula that narrows rows by parent work item IDs, work item types, and date thresholds while using `INDEX` to choose and reorder returned columns for presentation.
- **FilterFunction-ReplaceZerosWithBlanksUsingLET.txt** — Demonstrates using `LET` to store a `FILTER` result and post-process the spilled range so returned `0` values are displayed as blanks (`""`) when `0` is not a valid data value.
- **Macro-BlockingRefreshAndUpdateOfPivotTableAfter** — VBA macro that refreshes a query table in blocking mode, stamps a refresh timestamp, refreshes a pivot table cache, and returns to the summary sheet.
- **Macro-BulkAssignNotes** — VBA macro that reads values from the third row and bulk assigns them as cell notes in the second row across all populated columns of the active worksheet.
- **macro-resolve-dynamic-range-spacing-and-remove-excess-rows.txt** — VBA macros that insert blank rows to resolve `#SPILL!` errors in column A, then remove extraneous consecutive blank rows to clean the dataset.
