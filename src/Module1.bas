Option Explicit

'***************************************************************************
' stats_plots :: histplot
'
' Creates a histogram (vertical column chart) from a single column of
' numeric data.
'
' Interactive usage:
'   1. Select a column of numbers in Calc.
'   2. Run the histplot macro (Tools > Macros, or a toolbar button).
'   3. Enter the bin width when prompted.
'
' The macro writes the bin/count table to a sheet named "histplot" and
' embeds the chart on that sheet.
'***************************************************************************

Const DATA_SHEET As String = "histplot"
Const CHART_NAME As String = "histplot_chart"


Sub histplot()
    Dim oDoc As Object
    Dim oSel As Object
    Dim vData() As Double
    Dim nCount As Long
    Dim binWidth As Double
    Dim sInput As String

    oDoc = ThisComponent
    If IsNull(oDoc) Then
        MsgBox "No document is open.", 16, "stats_plots"
        Exit Sub
    End If
    If Not oDoc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then
        MsgBox "histplot must be run from a Calc spreadsheet.", 16, "stats_plots"
        Exit Sub
    End If

    oSel = oDoc.CurrentSelection
    If IsNull(oSel) Or Not oSel.supportsService("com.sun.star.sheet.SheetCellRange") Then
        MsgBox "Please select a single column of numeric data first.", 16, "stats_plots"
        Exit Sub
    End If

    nCount = ReadColumn(oSel, vData)
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

    BuildHistogram(oDoc, vData, nCount, binWidth)
End Sub


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
    oSheet = GetCleanSheet(oDoc, DATA_SHEET)
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


Function GetCleanSheet(oDoc As Object, sName As String) As Object
    Dim oSheets As Object
    oSheets = oDoc.Sheets
    If oSheets.hasByName(sName) Then oSheets.removeByName(sName)
    oSheets.insertNewByName(sName, oSheets.Count)
    GetCleanSheet = oSheets.getByName(sName)
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
    If oCharts.hasByName(CHART_NAME) Then oCharts.removeByName(CHART_NAME)
    ' bColumnHeaders = True (row 0 holds headers), bRowHeaders = True (col 0 holds labels)
    oCharts.addNewByName(CHART_NAME, oRect, oRanges, True, True)

    oChart = oCharts.getByName(CHART_NAME).getEmbeddedObject()
    oDiagram = oChart.createInstance("com.sun.star.chart.BarDiagram")
    oDiagram.Vertical = True                          ' vertical columns
    oChart.setDiagram(oDiagram)
    oChart.HasMainTitle = True
    oChart.Title.String = "Histogram"
End Sub
