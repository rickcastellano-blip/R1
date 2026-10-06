# Notes for Claude sessions on this repo

Read README.md first: it maps every button to its macro, Excel module, file and update macro.

## How work flows

- The user runs these macros in Excel on Windows (`XMaster.xlsm`); Claude cannot run Excel
  or compile VBA. Check logic with a Python port against the user's sample files, and say
  plainly what was and wasn't tested.
- Push changes to `claude/tender-ramanujan-047vzz` (the only branch, and the one the Update
  buttons read). The user pulls a change by clicking that module's **Update** button.
- Changes to `Updater.bas` need a manual paste into the **Updaters** module; say so.
- The user reports results with screenshots. Ask for the Immediate window (Ctrl+G) output if
  temporary `Debug.Print` logging is needed, and remove the logging afterwards.

## Code rules

- **Every module is self-contained.** No module calls into another; helpers are `Private`
  and copied where needed (e.g. `NiceScale`, `IsNum`, `CellText` exist in both graph
  modules). Only the button macros are `Public`.
- **VBA is case-insensitive.** No two names in a procedure may differ only by case
  (`eh`/`eH` broke a compile), and no name may be a keyword in disguise (`oR` = `Or`).
  Before pushing, scan each procedure for case-insensitive duplicate declarations.
- `VBA And` does not short-circuit: never compare a cell that may hold an error value
  (`#N/A`) directly; use `IsNum` / `CellText` / `IsBlankValue`.
- No success message boxes for the graphs, updaters or NT492 (errors only). RCPT keeps its
  per-import summary.
- Excel module names equal file names (BarGraph, LineGraph, RCPTImport, RCPTRecovered,
  NTBuild492); keep the button macro names stable so existing button assignments work.
- Line Graph settings are found by header text on the Buttons sheet, never by fixed cell.
- Numbers written into worksheet formulas go through `Num()` (locale-safe decimal point).

## Domain notes

- TTi logger CSVs: `#TimeStamp,Volts,TimeStamp,Amps,...` (also a 9-column layout and Test
  Bridge "recovered" exports, where headers may label every channel Volts). Rows with blank
  readings are skipped. Switch-on/off transients (a reading between levels) must not be taken
  as initial or final currents.
- RCPT (ASTM C1202): 60 V, 6 h, charge in coulombs; a file may contain several runs.
- NT BUILD 492: Table 1 picks U and t from the 30 V current; Dnssm uses the inverse error
  function via `NORM.S.INV`.
- Graph conventions follow the user's MATLAB `xgrapher.m` / `resize.m` (tags, column layout,
  palette, marker order, 550x400 figure proportions).
