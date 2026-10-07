# XMaster macros

VBA modules for the `XMaster.xlsm` workbook: concrete test imports (RCPT, NT BUILD 492)
and MATLAB-style line and bar charts. Each button on the workbook's **Buttons** sheet runs
one macro, and the **Update** button next to it pulls that macro's module from this repo.

## Buttons, modules and files

| Button | Macro | Excel module | File | Update macro |
|---|---|---|---|---|
| Bar Graph | `GenerateStrengthBarGraph` | BarGraph | `BarGraph.bas` | `UpdateBarGraph` |
| Line Graph | `GenerateLineGraph` | LineGraph | `LineGraph.bas` | `UpdateLineGraph` |
| RCPT | `AddRCPTRunFromCSV` | RCPTImport | `RCPTImport.bas` | `UpdateRCPT` |
| RCPT Recovered | `AddRun_RecoveredTTi` | RCPTRecovered | `RCPTRecovered.bas` | `UpdateRCPTRecovered` |
| NT492 | `AnalyzeNTBuild492` | NTBuild492 | `NTBuild492.bas` | `UpdateNTBuild492` |
| NT492 Recovered | `AnalyzeNTBuild492Recovered` | NTBuild492Recovered | `NTBuild492Recovered.bas` | `UpdateNTBuild492Recovered` |
| (the Update buttons) | – | Updaters | `Updater.bas` | – |
| Do not click | `DoNotClick` | Module7 | not in the repo | – |

An Update button replaces exactly one module. The modules are self-contained except the
two recovered importers: **RCPT Recovered** and **NT492 Recovered** only read their file,
then use the analysis in **RCPTImport** and **NTBuild492**. So analysis changes arrive with
the RCPT / NT492 Update buttons and apply to both buttons of a pair; a recovered Update
button only changes how its files are read.

## Updating

- The update macros download from the branch `claude/tender-ramanujan-047vzz` (the repo's
  default and only branch) through the GitHub API, and replace the module's code.
- Excel needs **File > Options > Trust Center > Trust Center Settings > Macro Settings >
  "Trust access to the VBA project object model"**.
- The updater never overwrites its own module. When `Updater.bas` changes, paste it into
  the **Updaters** module by hand.
- The Excel module names must match the table (the updater finds modules by name and
  creates a missing one).
- An update is silent when it works; a message appears only when something fails.
- After an update: Debug > Compile VBAProject, then save.

## Line Graph

Select a block containing one or more `graph` tags (also `graph linlog`, `graph loglin`,
`graph loglog`); each tag makes its own chart.

- The tag cell is the X column; Y is the next column; the series name is in the column left
  of X, on the series' first row. Axis titles are the row below the tag (linked to the cells).
- A row where X **and** Y are both blank starts a new series; two such rows end the graph.
  A row with one value or `#N/A` stays in the series with no point.
- Point labels: the first text column right of Y (or right of numeric ±Y / ±X error-bar
  columns). An empty column stops the search.
- Settings are read from the **Buttons** sheet, from the cell below each header (the block
  can be moved anywhere):

  | Header | Value |
  |---|---|
  | Line | 1 = connecting lines, 0 = markers only |
  | Color scheme | standard, stoplight, green→red |
  | Marker Size | points (2–72); blank = default |
  | Line Width | points (0.25–10), line and marker outline; blank = default |

- The legend goes in a corner where it covers no data (1 or 2 columns, raising the axes a
  step if needed), or outside the plot on the right when nothing fits.

## RCPT

Imports a TTi CPX400DP logger CSV into the **RCPT** sheet, one row (and current chart) per
6 h run; a file holding several runs gives several rows. Columns are read from the
`#TimeStamp` header (5- and 9-column layouts). **RCPT Recovered** reads Test Bridge
recovered logs (Volts in field 2, Amps in field 4) and uses the same analysis.

## NT BUILD 492

Imports a TTi logger CSV into the **NT492** sheet, one row per specimen. Columns A:M match
the *Data Summary* sheet of the NT492 Measurement workbook so rows copy straight across;
N:P add an RCPT-equivalent charge, its class and bulk resistivity. **NT492 Recovered**
reads Test Bridge recovered logs and uses the same analysis.
