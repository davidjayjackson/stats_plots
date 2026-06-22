Option Explicit

'***************************************************************************
' stats_plots
'
'   histplot  - histogram (vertical column chart) from a column of numbers.
'   boxplot   - box-and-whisker plot from a column of numbers.
'
' Both macros operate on the current Calc selection. Select a column of
' numbers, then run the macro (Tools > Macros, or a toolbar button).
'***************************************************************************

Const HIST_SHEET As String = "histplot"
Const HIST_CHART As String = "histplot_chart"
Const BOX_SHEET  As String = "boxplot"
Const BOX_CHART  As String = "boxplot_chart"


'==========================================================================
' histplot
'==========================================================================
Sub histplot()
    Dim vData() As Double
    Dim nCount As Long
    Dim binWidth As Double
    Dim sInput As String

    nCount = GetNumericSelection(vData)
    If nCount < 0 Then Exit Sub                      ' invalid context (already reported)
    If nCount = 0 Then
        MsgBox "The selection contains no numeric values.", 16, "stats_plots"
        Exit Sub
    End If

    sInput = Trim(InputBox("Bin width:", "histplot - bin width"))
    If sInput = "" Then Exit Sub                     ' user cancelled
    binWidth = Val(sInput)
    If binWidth <= 0 Then
        MsgBox "Bin width must be a number greater than zero.", 16, "stats_plots"
        Exit Sub
    End If

    BuildHistogram(ThisComponent, vData, nCount, binWidth)
End Sub


Sub BuildHistogram(oDoc As Object, vData() As Double, nCount As Long, binWidth As Double)
    Dim dMin As Double, dMax As Double
    Dim lo As Double
    Dim numBins As Long
    Dim counts() As Long
    Dim i As Long, idx As Long
    Dim oSheet As Object

    ' --- range of the data ---
    dMin = vData(0) : dMax = vData(0)
    For i = 1 To nCount - 1
        If vData(i) < dMin Then dMin = vData(i)
        If vData(i) > dMax Then dMax = vData(i)
    Next i

    ' --- bins aligned to multiples of binWidth ---
    lo = Int(dMin / binWidth) * binWidth
    numBins = Int((dMax - lo) / binWidth) + 1
    If numBins < 1 Then numBins = 1
    ReDim counts(0 To numBins - 1)

    For i = 0 To nCount - 1
        idx = Int((vData(i) - lo) / binWidth)
        If idx < 0 Then idx = 0
        If idx > numBins - 1 Then idx = numBins - 1
        counts(idx) = counts(idx) + 1
    Next i

    ' --- write the bin/count table to a clean sheet ---
    oSheet = GetCleanSheet(oDoc, HIST_SHEET)
    oSheet.getCellByPosition(0, 0).setString("Bin")
    oSheet.getCellByPosition(1, 0).setString("Count")
    For i = 0 To numBins - 1
        oSheet.getCellByPosition(0, i + 1).setString( _
            BinLabel(lo + i * binWidth, lo + (i + 1) * binWidth))
        oSheet.getCellByPosition(1, i + 1).setValue(counts(i))
    Next i

    CreateColumnChart(oSheet, numBins)
End Sub


Function BinLabel(loEdge As Double, hiEdge As Double) As String
    BinLabel = "[" & Format(loEdge, "0.###") & ", " & Format(hiEdge, "0.###") & ")"
End Function


Sub CreateColumnChart(oSheet As Object, numBins As Long)
    Dim oCharts As Object
    Dim oRect As New com.sun.star.awt.Rectangle
    Dim oRanges(0) As New com.sun.star.table.CellRangeAddress
    Dim oChart As Object, oDiagram As Object

    oRect.X = 8000 : oRect.Y = 500
    oRect.Width = 14000 : oRect.Height = 9000

    oRanges(0).Sheet = oSheet.RangeAddress.Sheet
    oRanges(0).StartColumn = 0
    oRanges(0).StartRow = 0
    oRanges(0).EndColumn = 1
    oRanges(0).EndRow = numBins                      ' header row + numBins data rows

    oCharts = oSheet.Charts
    If oCharts.hasByName(HIST_CHART) Then oCharts.removeByName(HIST_CHART)
    ' bColumnHeaders = True (row 0 holds headers), bRowHeaders = True (col 0 holds labels)
    oCharts.addNewByName(HIST_CHART, oRect, oRanges, True, True)

    oChart = oCharts.getByName(HIST_CHART).getEmbeddedObject()
    oDiagram = oChart.createInstance("com.sun.star.chart.BarDiagram")
    oChart.setDiagram(oDiagram)
    oDiagram.Vertical = False                         ' False = vertical columns
    oChart.HasMainTitle = True
    oChart.Title.String = "Histogram"
End Sub


'==========================================================================
' boxplot
'==========================================================================
Sub boxplot()
    Dim vData() As Double
    Dim nCount As Long

    nCount = GetNumericSelection(vData)
    If nCount < 0 Then Exit Sub                      ' invalid context (already reported)
    If nCount = 0 Then
        MsgBox "The selection contains no numeric values.", 16, "stats_plots"
        Exit Sub
    End If

    BuildBoxplot(ThisComponent, vData, nCount)
End Sub


Sub BuildBoxplot(oDoc As Object, vData() As Double, nCount As Long)
    Dim sorted() As Double
    Dim i As Long
    Dim dMin As Double, q1 As Double, med As Double, q3 As Double, dMax As Double
    Dim iqr As Double, loFence As Double, hiFence As Double
    Dim loWhisker As Double, hiWhisker As Double
    Dim outliers() As Double, nOut As Long
    Dim oSheet As Object

    ' --- sort a copy and compute the five-number summary ---
    ReDim sorted(0 To nCount - 1)
    For i = 0 To nCount - 1 : sorted(i) = vData(i) : Next i
    SortDoubles(sorted, 0, nCount - 1)

    dMin = sorted(0)
    dMax = sorted(nCount - 1)
    q1  = Quantile(sorted, nCount, 0.25)
    med = Quantile(sorted, nCount, 0.5)
    q3  = Quantile(sorted, nCount, 0.75)

    ' --- Tukey outlier detection (1.5 * IQR rule) ---
    iqr = q3 - q1
    loFence = q1 - 1.5 * iqr
    hiFence = q3 + 1.5 * iqr

    ' Whiskers reach the most extreme values that lie within the fences.
    loWhisker = dMin
    For i = 0 To nCount - 1
        If sorted(i) >= loFence Then loWhisker = sorted(i) : Exit For
    Next i
    hiWhisker = dMax
    For i = nCount - 1 To 0 Step -1
        If sorted(i) <= hiFence Then hiWhisker = sorted(i) : Exit For
    Next i

    ' Values outside the fences are outliers.
    nOut = 0
    ReDim outliers(0 To nCount - 1)
    For i = 0 To nCount - 1
        If sorted(i) < loFence Or sorted(i) > hiFence Then
            outliers(nOut) = sorted(i)
            nOut = nOut + 1
        End If
    Next i

    oSheet = GetCleanSheet(oDoc, BOX_SHEET)

    ' --- summary table (columns A:B, for reference) ---
    oSheet.getCellByPosition(0, 0).setString("Statistic")
    oSheet.getCellByPosition(1, 0).setString("Value")
    WriteStat(oSheet, 1, "Minimum",     dMin)
    WriteStat(oSheet, 2, "Q1",          q1)
    WriteStat(oSheet, 3, "Median",      med)
    WriteStat(oSheet, 4, "Q3",          q3)
    WriteStat(oSheet, 5, "Maximum",     dMax)
    WriteStat(oSheet, 6, "IQR",         iqr)
    WriteStat(oSheet, 7, "Lower fence", loFence)
    WriteStat(oSheet, 8, "Upper fence", hiFence)
    WriteStat(oSheet, 9, "Outliers",    nOut)

    ' --- outlier values (column I), also used as the chart's point source ---
    oSheet.getCellByPosition(8, 0).setString("Outliers")
    For i = 0 To nOut - 1
        oSheet.getCellByPosition(8, i + 1).setValue(outliers(i))
    Next i

    ' --- stacked-column source data (columns D:G) ---
    ' The box is drawn as a stacked column of three segments:
    '   Base       = Q1          (transparent -> lifts the box to start at Q1)
    '   Q1-Median  = Median - Q1 (lower half of the box)
    '   Median-Q3  = Q3 - Median (upper half of the box)
    ' The border between the two visible segments is the median line. Whiskers
    ' to the fence values are added afterwards as error bars (see CreateBoxChart).
    oSheet.getCellByPosition(3, 0).setString("")          ' category header
    oSheet.getCellByPosition(4, 0).setString("Base")
    oSheet.getCellByPosition(5, 0).setString("Q1-Median")
    oSheet.getCellByPosition(6, 0).setString("Median-Q3")
    oSheet.getCellByPosition(3, 1).setString("Data")
    oSheet.getCellByPosition(4, 1).setValue(q1)
    oSheet.getCellByPosition(5, 1).setValue(med - q1)
    oSheet.getCellByPosition(6, 1).setValue(q3 - med)

    ' Axis spans the full data range (Min..Max) so any outliers stay in view.
    CreateBoxChart(oSheet, q1, q3, loWhisker, hiWhisker, dMin, dMax, nOut)
End Sub


Sub WriteStat(oSheet As Object, nRow As Long, sName As String, dVal As Double)
    oSheet.getCellByPosition(0, nRow).setString(sName)
    oSheet.getCellByPosition(1, nRow).setValue(dVal)
End Sub


Sub CreateBoxChart(oSheet As Object, q1 As Double, q3 As Double, _
                   loWhisker As Double, hiWhisker As Double, _
                   axisMin As Double, axisMax As Double, nOut As Long)
    Dim oCharts As Object
    Dim oRect As New com.sun.star.awt.Rectangle
    Dim oRanges(0) As New com.sun.star.table.CellRangeAddress
    Dim oChart As Object, oDiagram As Object
    Dim oProp As Object
    Dim oYAxis As Object, pad As Double
    Dim oDia2 As Object, oCoo As Variant, oTypes As Variant, oSeries As Variant
    Dim oErr As Object
    Dim oCooSys As Object, oProvider As Object, oLineType As Object
    Dim oVals As Object, oLab As Object, oDS As Object
    Dim seqArr(0) As Object
    Dim k As Long

    oRect.X = 8000 : oRect.Y = 500
    oRect.Width = 9000 : oRect.Height = 12000

    oRanges(0).Sheet = oSheet.RangeAddress.Sheet
    oRanges(0).StartColumn = 3                        ' D
    oRanges(0).StartRow = 0
    oRanges(0).EndColumn = 6                          ' G
    oRanges(0).EndRow = 1

    oCharts = oSheet.Charts
    If oCharts.hasByName(BOX_CHART) Then oCharts.removeByName(BOX_CHART)
    oCharts.addNewByName(BOX_CHART, oRect, oRanges, True, True)

    oChart = oCharts.getByName(BOX_CHART).getEmbeddedObject()
    oDiagram = oChart.createInstance("com.sun.star.chart.BarDiagram")
    oChart.setDiagram(oDiagram)
    oDiagram.Vertical = False                         ' False = vertical columns
    oDiagram.Stacked = True                           ' stack the three segments
    oChart.HasMainTitle = True
    oChart.Title.String = "Boxplot"
    oChart.HasLegend = False

    ' --- widen the value axis so whiskers and outliers stay visible ---
    ' Auto-scaling only considers the stacked column tops (Q1..Q3) and would
    ' otherwise clip the whiskers and any outlier points.
    pad = (axisMax - axisMin) * 0.1
    If pad <= 0 Then pad = 1
    oYAxis = oDiagram.getYAxis()
    oYAxis.AutoMin = False
    oYAxis.Min = axisMin - pad
    oYAxis.AutoMax = False
    oYAxis.Max = axisMax + pad

    ' --- style the three series (0 = base, 1 = Q1..median, 2 = median..Q3) ---
    ' The two box halves use slightly different shades so the boundary between
    ' them - the median - reads as a distinct line across the box.
    oProp = oDiagram.getDataRowProperties(0)          ' base: invisible
    oProp.FillStyle = com.sun.star.drawing.FillStyle.NONE
    oProp.LineStyle = com.sun.star.drawing.LineStyle.NONE

    oProp = oDiagram.getDataRowProperties(1)          ' lower box half (Q1..median)
    oProp.FillStyle = com.sun.star.drawing.FillStyle.SOLID
    oProp.FillColor = RGB(120, 160, 210)
    oProp.LineStyle = com.sun.star.drawing.LineStyle.SOLID
    oProp.LineColor = RGB(0, 0, 0)

    oProp = oDiagram.getDataRowProperties(2)          ' upper box half (median..Q3)
    oProp.FillStyle = com.sun.star.drawing.FillStyle.SOLID
    oProp.FillColor = RGB(185, 210, 235)
    oProp.LineStyle = com.sun.star.drawing.LineStyle.SOLID
    oProp.LineColor = RGB(0, 0, 0)

    ' --- whiskers as error bars (via the chart2 model) ---
    ' The chart2 ErrorBar service must be created from the global service
    ' manager; the embedded chart's own createInstance only serves the legacy
    ' com.sun.star.chart factory and returns Null for chart2 services.
    oDia2 = oChart.getFirstDiagram()
    oCoo = oDia2.getCoordinateSystems()
    oTypes = oCoo(0).getChartTypes()
    oSeries = oTypes(0).getDataSeries()

    ' lower whisker: negative error on the base series (its top sits at Q1)
    oErr = createUnoService("com.sun.star.chart2.ErrorBar")
    oErr.ErrorBarStyle = com.sun.star.chart.ErrorBarStyle.ABSOLUTE
    oErr.ShowPositiveError = False
    oErr.ShowNegativeError = True
    oErr.PositiveError = 0
    oErr.NegativeError = q1 - loWhisker
    oSeries(0).ErrorBarY = oErr

    ' upper whisker: positive error on the top series (its top sits at Q3)
    oErr = createUnoService("com.sun.star.chart2.ErrorBar")
    oErr.ErrorBarStyle = com.sun.star.chart.ErrorBarStyle.ABSOLUTE
    oErr.ShowPositiveError = True
    oErr.ShowNegativeError = False
    oErr.PositiveError = hiWhisker - q3
    oErr.NegativeError = 0
    oSeries(2).ErrorBarY = oErr

    ' --- outliers as individual point markers (chart2 combination) ---
    ' Each outlier (column I of the sheet) becomes its own single-point symbol
    ' series. They all share the one category, so they stack vertically above or
    ' below the box at the same horizontal position.
    If nOut > 0 Then
        Dim oSymbol As New com.sun.star.chart2.Symbol
        Dim oSz As New com.sun.star.awt.Size

        oCooSys = oCoo(0)
        oProvider = oChart.getDataProvider()
        oLineType = createUnoService("com.sun.star.chart2.LineChartType")
        oCooSys.addChartType(oLineType)

        oSymbol.Style = com.sun.star.chart2.SymbolStyle.STANDARD
        oSymbol.StandardSymbol = 0
        oSz.Width = 300 : oSz.Height = 300
        oSymbol.Size = oSz

        For k = 0 To nOut - 1
            oVals = oProvider.createDataSequenceByRangeRepresentation( _
                "$" & oSheet.Name & ".$I$" & (k + 2))
            oVals.Role = "values-y"
            oLab = createUnoService("com.sun.star.chart2.data.LabeledDataSequence")
            oLab.setValues(oVals)
            oDS = createUnoService("com.sun.star.chart2.DataSeries")
            seqArr(0) = oLab
            oDS.setData(seqArr())
            oDS.setPropertyValue("LineStyle", com.sun.star.drawing.LineStyle.NONE)
            oDS.setPropertyValue("Color", RGB(200, 30, 30))
            oDS.setPropertyValue("Symbol", oSymbol)
            oLineType.addDataSeries(oDS)
        Next k
    End If
End Sub


'==========================================================================
' Shared helpers
'==========================================================================

' Numeric values from the current selection, returned in vOut().
' Returns the count, or -1 if the context/selection is invalid (in which
' case a message has already been shown to the user).
Function GetNumericSelection(ByRef vOut() As Double) As Long
    Dim oDoc As Object, oSel As Object

    oDoc = ThisComponent
    If IsNull(oDoc) Then
        MsgBox "No document is open.", 16, "stats_plots"
        GetNumericSelection = -1 : Exit Function
    End If
    If Not oDoc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then
        MsgBox "stats_plots must be run from a Calc spreadsheet.", 16, "stats_plots"
        GetNumericSelection = -1 : Exit Function
    End If

    oSel = oDoc.CurrentSelection
    If IsNull(oSel) Or Not oSel.supportsService("com.sun.star.sheet.SheetCellRange") Then
        MsgBox "Please select a single column of numeric data first.", 16, "stats_plots"
        GetNumericSelection = -1 : Exit Function
    End If

    GetNumericSelection = ReadColumn(oSel, vOut)
End Function


' Read every numeric value from a cell range into vOut().
' Returns the count of values found.
Function ReadColumn(oRange As Object, ByRef vOut() As Double) As Long
    Dim vRows As Variant
    Dim r As Long, c As Long                         ' Long: a full column has more than 32767 rows
    Dim vVal As Variant
    Dim n As Long, total As Long

    vRows = oRange.getDataArray()

    ' First pass: count numeric cells so vOut can be sized once.
    total = 0
    For r = LBound(vRows) To UBound(vRows)
        For c = LBound(vRows(r)) To UBound(vRows(r))
            If VarType(vRows(r)(c)) = 5 Then total = total + 1   ' 5 = Double (numeric cell)
        Next c
    Next r

    If total = 0 Then
        ReDim vOut(0 To 0)
        ReadColumn = 0
        Exit Function
    End If

    ' Second pass: copy the numeric values.
    ReDim vOut(0 To total - 1)
    n = 0
    For r = LBound(vRows) To UBound(vRows)
        For c = LBound(vRows(r)) To UBound(vRows(r))
            vVal = vRows(r)(c)
            If VarType(vVal) = 5 Then
                vOut(n) = CDbl(vVal)
                n = n + 1
            End If
        Next c
    Next r
    ReadColumn = n
End Function


Function GetCleanSheet(oDoc As Object, sName As String) As Object
    Dim oSheets As Object
    oSheets = oDoc.Sheets
    If oSheets.hasByName(sName) Then oSheets.removeByName(sName)
    oSheets.insertNewByName(sName, oSheets.Count)
    GetCleanSheet = oSheets.getByName(sName)
End Function


' Linear-interpolation quantile (type 7 / Calc QUARTILE.INC) on a sorted array.
Function Quantile(sorted() As Double, n As Long, p As Double) As Double
    Dim h As Double, k As Long, frac As Double
    If n = 1 Then
        Quantile = sorted(0) : Exit Function
    End If
    h = (n - 1) * p
    k = Int(h)
    frac = h - k
    If k >= n - 1 Then
        Quantile = sorted(n - 1)
    Else
        Quantile = sorted(k) + frac * (sorted(k + 1) - sorted(k))
    End If
End Function


' In-place quicksort of arr(lo..hi).
Sub SortDoubles(arr() As Double, lo As Long, hi As Long)
    Dim i As Long, j As Long
    Dim pivot As Double, tmp As Double

    i = lo : j = hi
    pivot = arr((lo + hi) \ 2)
    Do While i <= j
        Do While arr(i) < pivot : i = i + 1 : Loop
        Do While arr(j) > pivot : j = j - 1 : Loop
        If i <= j Then
            tmp = arr(i) : arr(i) = arr(j) : arr(j) = tmp
            i = i + 1 : j = j - 1
        End If
    Loop
    If lo < j Then SortDoubles(arr, lo, j)
    If i < hi Then SortDoubles(arr, i, hi)
End Sub
