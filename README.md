# stats_plots
LibreOffice Calc extension called stats_plots to create histograms and boxplots
The extension will contain two functions, one to create a histogram called histplot,
and the second call boxplot to create boxplot. The histplot function, in addition to the column, takes a second variable to set the bin width.

## Project layout

- `src/Module1.bas` — canonical LibreOffice Basic source (the macros).
- `extension/` — static `.oxt` layout (`description.xml`, `META-INF/manifest.xml`, library index).
- `build.ps1` — generates the packaged module from `src/Module1.bas` and zips `stats_plots.oxt`.

## Building

```powershell
./build.ps1
```

This produces `stats_plots.oxt` in the project root.

## Installing

In LibreOffice: **Tools ▸ Extension Manager… ▸ Add…**, select `stats_plots.oxt`, then restart LibreOffice.

## Using histplot

1. Select a column of numeric values in Calc.
2. Run the `histplot` macro (**Tools ▸ Macros ▸ Run Macro…**, library `stats_plots`).
3. Enter the bin width when prompted.

The macro writes the bin/count table to a sheet named `histplot` and embeds the
histogram (a vertical column chart) on that sheet.

## Using boxplot

1. Select a column of numeric values in Calc.
2. Run the `boxplot` macro (**Tools ▸ Macros ▸ Run Macro…**, library `stats_plots`).

The macro writes the five-number summary (minimum, Q1, median, Q3, maximum) to a
sheet named `boxplot` and embeds a box-and-whisker chart on that sheet. The box
spans Q1–Q3 and the whiskers span min–max.

> Since the chart engine has no native box-plot type, the box is built from a
> stacked column chart: a transparent base lifts the box to Q1, and two shaded
> segments (Q1–median and median–Q3) make the median read as a line across the
> box. The whiskers to Min and Max are drawn as error bars. Quartiles use the
> linear-interpolation method (Calc's `QUARTILE.INC`).
