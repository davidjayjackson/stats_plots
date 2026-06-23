# stats_plots
LibreOffice Calc extension called stats_plots to create histograms and boxplots
The extension will contain two functions, one to create a histogram called histplot,
and the second call boxplot to create boxplot. The histplot function, in addition to the column, takes a second variable to set the bin width.

See [CHANGELOG.md](CHANGELOG.md) for release history.

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

**From a release (recommended)**

1. Download `stats_plots.oxt` from the [latest release](https://github.com/davidjayjackson/stats_plots/releases/latest).
2. In LibreOffice: **Tools ▸ Extension Manager… ▸ Add…**, select the downloaded `stats_plots.oxt`.
3. Restart LibreOffice.

**From source**

Run `./build.ps1` (see [Building](#building)) to produce `stats_plots.oxt`, then add it via
**Tools ▸ Extension Manager… ▸ Add…** and restart LibreOffice.

To upgrade, install the new `.oxt` the same way — the Extension Manager replaces the
previous version. The macros appear under library `stats_plots` in **Tools ▸ Macros**.

## Menu

Once installed, both commands appear at the bottom of the **Data** menu, just
below the built-in **Statistics** submenu: **histplot** and **boxplot**. You can
also run them from **Tools ▸ Macros ▸ Run Macro…** (library `stats_plots`).

## Using histplot

1. Select a column of numeric values in Calc.
2. Run **histplot** from the **Data** menu (or via
   **Tools ▸ Macros ▸ Run Macro…**, library `stats_plots`).
3. Enter the bin width when prompted.

The macro writes the bin/count table to a sheet named `histplot` and embeds the
histogram (a vertical column chart) on that sheet.

## Using boxplot

1. Select a column of numeric values in Calc.
2. Run **boxplot** from the **Data** menu (or via
   **Tools ▸ Macros ▸ Run Macro…**, library `stats_plots`).

The macro writes a summary table to a sheet named `boxplot` and embeds a
box-and-whisker chart on that sheet. The box spans Q1–Q3 with a median line; the
whiskers extend to the most extreme values within the Tukey fences, and any
outliers are drawn as separate point markers.

**Outlier detection** uses the 1.5 × IQR rule:

- `IQR = Q3 − Q1`
- Lower fence = `Q1 − 1.5 × IQR`, upper fence = `Q3 + 1.5 × IQR`
- Values outside the fences are outliers; the whiskers stop at the most extreme
  values that are *inside* the fences.

The summary table lists the five-number summary plus IQR, both fences, the
outlier count, and the outlier values themselves (column `I`).

> Since the chart engine has no native box-plot type, the box is built from a
> stacked column chart: a transparent base lifts the box to Q1, and two shaded
> segments (Q1–median and median–Q3) make the median read as a line across the
> box. Whiskers are drawn as error bars, and each outlier is a single-point
> symbol series overlaid on the box. Quartiles use the linear-interpolation
> method (Calc's `QUARTILE.INC`).
