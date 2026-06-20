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
    oDiagram.Vertical = True                          ' vertical columns
    oChart.setDiagram(oDiagram)
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

    oSheet = GetCleanSheet(oDoc, BOX_SHEET)

    ' --- five-number summary table (columns A:B, for reference) ---
    oSheet.getCellByPosition(0, 0).setString("Statistic")
    oSheet.getCellByPosition(1, 0).setString("Value")
    WriteStat(oSheet, 1, "Minimum", dMin)
    WriteStat(oSheet, 2, "Q1",      q1)
    WriteStat(oSheet, 3, "Median",  med)
    WriteStat(oSheet, 4, "Q3",      q3)
    WriteStat(oSheet, 5, "Maximum", dMax)

    ' --- candlestick source data (columns D:H) ---
    ' A candlestick (Open-Low-High-Close) draws a box from Open to Close with
    ' a whisker line from Low to High. Map the box to Q1..Q3 and whiskers to
    ' Min..Max. (The median line is not rendered by the legacy stock chart.)
    oSheet.getCellByPosition(3, 0).setString("")     ' category header
    oSheet.getCellByPosition(4, 0).setString("Open")
    oSheet.getCellByPosition(5, 0).setString("Low")
    oSheet.getCellByPosition(6, 0).setString("High")
    oSheet.getCellByPosition(7, 0).setString("Close")
    oSheet.getCellByPosition(3, 1).setString("Data")
    oSheet.getCellByPosition(4, 1).setValue(q1)      ' Open  = Q1
    oSheet.getCellByPosition(5, 1).setValue(dMin)    ' Low   = minimum
    oSheet.getCellByPosition(6, 1).setValue(dMax)    ' High  = maximum
    oSheet.getCellByPosition(7, 1).setValue(q3)      ' Close = Q3

    CreateStockChart(oSheet)
End Sub


Sub WriteStat(oSheet As Object, nRow As Long, sName As String, dVal As Double)
    oSheet.getCellByPosition(0, nRow).setString(sName)
    oSheet.getCellByPosition(1, nRow).setValue(dVal)
End Sub


Sub CreateStockChart(oSheet As Object)
    Dim oCharts As Object
    Dim oRect As New com.sun.star.awt.Rectangle
    Dim oRanges(0) As New com.sun.star.table.CellRangeAddress
    Dim oChart As Object, oDiagram As Object

    oRect.X = 8000 : oRect.Y = 500
    oRect.Width = 10000 : oRect.Height = 11000

    oRanges(0).Sheet = oSheet.RangeAddress.Sheet
    oRanges(0).StartColumn = 3                        ' D
    oRanges(0).StartRow = 0
    oRanges(0).EndColumn = 7                          ' H
    oRanges(0).EndRow = 1

    oCharts = oSheet.Charts
    If oCharts.hasByName(BOX_CHART) Then oCharts.removeByName(BOX_CHART)
    oCharts.addNewByName(BOX_CHART, oRect, oRanges, True, True)

    oChart = oCharts.getByName(BOX_CHART).getEmbeddedObject()
    oDiagram = oChart.createInstance("com.sun.star.chart.StockDiagram")
    oDiagram.Volume = False
    oDiagram.UpDown = True                            ' candlestick (box) rendering
    oChart.setDiagram(oDiagram)
    oChart.HasMainTitle = True
    oChart.Title.String = "Boxplot"
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
    Dim r As Integer, c As Integer
    Dim vVal As Variant
    Dim n As Long

    vRows = oRange.getDataArray()
    ReDim vOut(0 To 0)
    n = 0
    For r = LBound(vRows) To UBound(vRows)
        For c = LBound(vRows(r)) To UBound(vRows(r))
            vVal = vRows(r)(c)
            If VarType(vVal) = 5 Then                ' 5 = Double (numeric cell)
                ReDim Preserve vOut(0 To n)
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
