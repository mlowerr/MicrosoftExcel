# AGENTS.md

## Project overview
- This repository is a small collection of Microsoft Excel formulas and VBA macro snippets.
- The files are single-purpose examples; there is no build system or automated test suite.
- The project lives entirely in the repository root with plain-text files.

## Files and responsibilities
- `FilterFunction-LimitAndReorderReturnedColumns.txt`: Excel `FILTER`/`INDEX` formula example that limits rows and reorders returned columns.
- `FilterFunction-ReplaceZerosWithBlanksUsingLET.txt`: Excel `LET` + `FILTER` example that replaces returned `0` values with blanks for display-only outputs where `0` is not a valid real value.
- `Macro-BlockingRefreshAndUpdateOfPivotTableAfter`: VBA macro to refresh a query table, timestamp it, refresh a pivot cache, and return to a summary sheet.
- `Macro-BulkAssignNotes`: VBA macro to bulk-assign cell notes from row 3 into row 2.
- `macro-resolve-dynamic-range-spacing-and-remove-excess-rows.txt`: VBA macros to insert blank rows for `#SPILL!` fixes, then remove excess blank rows.
- `README.md`: Top-level description and file index.

## Working conventions
- New formula tip files should use clear `FilterFunction-...` names when they are FILTER-specific examples.
- Keep updates focused on the specific Excel formula or VBA macro referenced by the filename.
- Preserve the existing plain-text format of the macro/formula files.
- Prefer a short instructional structure in snippet files: title, usage context, "What you get", "Why you would use this", then the formula/macro.
- Update the file list and descriptions in `README.md` whenever files change.

## Required maintenance steps
- After completing each work request, review and update this `AGENTS.md` file to capture any newly learned project context.
- Every time you run, review and update `README.md` to keep repository information current.
