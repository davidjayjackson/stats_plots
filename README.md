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

> `boxplot` is not implemented yet.
