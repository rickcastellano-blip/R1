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
- Column-header tags (in the row below the tag, right of Y): `errorbar` = ±Y errors,
  `errorbarx` = ±X errors, `note` = point labels. If any is present only those columns are
  used; otherwise the columns right of Y are read as below.
- Columns right of Y: numeric ±Y, then ±X error bars (drawn in the series colour), and
  point labels in the first text column. Only columns up to the first column that is
  blank over all the graph's data rows are used, so a table beyond a blank column is
  never read as error bars or labels.
- Settings are fixed cells on the **Buttons** sheet (if they move, the constants at the top
  of LineGraph must change):

  | Cell | Setting |
  |---|---|
  | F7 | Line: 1 = connecting lines, 0 = markers only |
  | G7 | Color scheme: standard, stoplight, green→red, darkening pairs (as the Bar Graph's), pairs (light/full of each standard colour: light blue, blue, light red, red, gray, black, ...) |
  | H7 | Marker Size: points (2–72); blank = default |
  | I7 | Line Width: points (0.25–10), line and marker outline; blank = default |
  | J7 | Label Position: right (default), above, below, left, in (centred on the point), or upper right |
  | K7 | Label font size: points (4–72); blank = 10 |
  | L7 | Width: figure width in inches (1–20); blank = 5. Only the plot widens: height, text, lines and the left/right margins stay fixed |

- The legend goes in a corner where it covers no data (1 or 2 columns, raising the axes a
  step if needed), or outside the plot on the right when nothing fits.
- Point labels are kept inside the plot and off the legend (axes widened a step at a time).
- Both graphs: no fill behind the figure, white behind the axes.
- Both graphs save a PNG of every chart (`<sheet> <chart name>.png`) to the fixed folder
  `FIG_DIR` at the top of each module (the team's OneDrive `General\Fig Pix`), drawn at 4x size
  (`FIG_SCALE`, about 384 dpi).

## Bar Graph

Select a block with one or more `bargraph` tags (each optionally with an `xlabels` tag below it,
in the same column); each tag makes its own chart, stacked down the right of the selection. Each block of values is a
series (legend entry); each row within a block is a group on the x axis. A blank cell ends a
block; `#N/A` stays in it as a group with no bar. The series name is in the column left of
the tag, on the block's first row. The column right of the values holds ± errors; bar labels
(data tags) are text in that column (then no error bars), or whatever is in the next
column (text or numbers, shown as displayed), drawn above the error bars. Header tags `errorbar` and `note` in the row below the tag
(right of the values) pick those columns instead, as in the Line Graph. **Buttons!H3** sets their direction: `horiz` (default) or `vert`; **Buttons!I3** their font size in points (blank = 10). **Buttons!L3** = `off` draws no legend. **Buttons!K3** says where a legend that fits no corner goes: `right` (default; the chart
widens) or `above` (the chart grows taller and the plot moves down). **Buttons!J3** sets the figure width in inches (blank = 5); only the
plot widens (height, text, lines and the left/right margins stay fixed). An empty column stops the search. The colour scheme is
**Buttons!G3**: standard (the xgrapher.m colour list, as in the Line Graph), or darkening
pairs (light blue, light red, blue, red, dark blue, dark red, from the standard blue and red;
just blue and red when there is only one pair), or darkening triplets (the same with gray as a
third colour: light blue, light red, light gray, blue, red, gray, dark blue, dark red, black;
just blue, red, gray for up to three series). Or pairs: each standard colour light then full (light
blue, blue, light red, red, gray, black, ...). Or triplets: light, full, dark of each
standard colour (light blue, blue, dark blue, light red, red, dark red, light gray, gray, black, ...). The legend is
placed like the Line Graph's: a clear corner (1 or 2 columns, raising the axis a step if
needed), or outside the plot on the right.

## RCPT

Imports a TTi CPX400DP logger CSV into the **RCPT** sheet, one row (and current chart) per
6 h run; a file holding several runs gives several rows. Columns are read from the
`#TimeStamp` header (5- and 9-column layouts). A log with both supply channels gives a run
for each channel that reached test voltage, named from the file name split at commas
(`11368J, 11366J` → channel 1 = 11368J, channel 2 = 11366J), else the name plus ` ch2`. **RCPT Recovered** reads Test Bridge
recovered logs (Volts in field 2, Amps in field 4) and uses the same analysis.

## NT BUILD 492

Imports a TTi logger CSV into the **NT492** sheet, one row per specimen. Columns A:M match
the *Data Summary* sheet of the NT492 Measurement workbook so rows copy straight across;
N:P add an RCPT-equivalent charge, its class and bulk resistivity. **NT492 Recovered**
reads Test Bridge recovered logs and uses the same analysis.
