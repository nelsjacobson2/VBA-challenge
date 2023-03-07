
Sub RunStockData()
    Dim ws As Worksheet
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        Call StockData
    Next ws
End Sub

Sub StockData()
' Declare variables
Dim TickerSymbol As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalVolume As Double
Dim OpenValue As Double
Dim CloseValue As Double
Dim LookupValue As String
Dim TableArray As Range
Dim FirstInstance As Long
Dim LastInstance As Long
Dim ws As Worksheet
Dim last_Row As Long
Dim rng As Range
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim LargestVolume As Double
Dim minTicker As String
Dim maxTicker As String
Dim maxVolTicker As String
    
Set TableArray = Range("A:G")

' Set initial values for variables
TickerSymbol = ""
YearlyChange = 0
PercentChange = 0
TotalVolume = 0

' Loop through each row of data
For i = 2 To Range("A" & Rows.Count).End(xlUp).Row

    ' Check if we've moved on to a new ticker symbol
    If Range("A" & i).Value <> TickerSymbol Then
        If TickerSymbol <> "" Then
            ' Output results for previous ticker symbol
            Range("I" & Rows.Count).End(xlUp).Offset(1, 0).Value = TickerSymbol
            Range("J" & Rows.Count).End(xlUp).Offset(1, 0).Value = YearlyChange
            Range("K" & Rows.Count).End(xlUp).Offset(1, 0).Value = PercentChange
            Range("L" & Rows.Count).End(xlUp).Offset(1, 0).Value = TotalVolume
        End If
        
        ' Set new ticker symbol
        TickerSymbol = Range("A" & i).Value
        
        ' Reset variables for new ticker symbol
        YearlyChange = 0
        PercentChange = 0
        TotalVolume = 0
        
        ' Find the first instance of TickerSymbol
        FirstInstance = WorksheetFunction.Match(TickerSymbol, TableArray.Columns(1), 0)
        
        ' Get the open value for the first instance of the TickerSymbol
        OpenValue = TableArray.Cells(FirstInstance, 3).Value
        
        ' Find the last instance of TickerSymbol
        LastInstance = WorksheetFunction.Match(TickerSymbol, TableArray.Columns(1), xlPrevious)
        
        ' Get the close value for the last instance of the TickerSymbol
        CloseValue = TableArray.Cells(LastInstance, 5).Value
        
        ' Calculate yearly change and percent change
        YearlyChange = CloseValue - OpenValue
        If OpenValue <> 0 Then
            PercentChange = YearlyChange / OpenValue
        End If
        
    Else
        ' Add to TotalVolume for current ticker symbol
        TotalVolume = TotalVolume + Range("G" & i).Value
    End If
    
Next i

'Create row headers for output
Range("I1").Value = "Ticker Symbol"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Volume"

'AutoFit Columns
Cells.EntireColumn.AutoFit

' Output results for last ticker symbol
Range("I" & Rows.Count).End(xlUp).Offset(1, 0).Value = TickerSymbol
Range("J" & Rows.Count).End(xlUp).Offset(1, 0).Value = YearlyChange
Range("K" & Rows.Count).End(xlUp).Offset(1, 0).Value = PercentChange
Range("L" & Rows.Count).End(xlUp).Offset(1, 0).Value = TotalVolume

'Conditional formatting
For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
        Set rng = ws.Range("K2:K" & lastRow)
        
        ' Remove any existing conditional formatting for column K
        rng.FormatConditions.Delete
        
        ' Add new conditional formatting for column K
        With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.Color = vbRed
        End With
        With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            .Interior.Color = vbGreen
        End With
    Next ws
    
For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
        Set rng = ws.Range("J2:J" & lastRow)
        
        ' Remove any existing conditional formatting for column K
        rng.FormatConditions.Delete
        
        ' Add new conditional formatting for column K
        With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.Color = vbRed
        End With
        With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            .Interior.Color = vbGreen
        End With
    Next ws
    
'Greatest Percent Increase
GreatestIncrease = Application.WorksheetFunction.Max(Range("K:K"))

' Find the minimum value in column K
    GreatestDecrease = Application.WorksheetFunction.Min(Range("K:K"))
    
' Find the ticker symbol for the minimum value
    minTicker = Range("I" & Application.WorksheetFunction.Match(GreatestDecrease, Range("K2:K" & Cells(Rows.Count, "K").End(xlUp).Row), 0) + 1).Value
    
' Find the maximum value in column K
    GreatestIncrease = Application.WorksheetFunction.Max(Range("K:K"))
    
' Find the ticker symbol for the maximum value
    maxTicker = Range("I" & Application.WorksheetFunction.Match(GreatestIncrease, Range("K2:K" & Cells(Rows.Count, "K").End(xlUp).Row), 0) + 1).Value
    
' Find the ticker symbol with the greatest total volume
    LargestVolume = Application.WorksheetFunction.Max(Range("L:L"))
    maxVolTicker = Range("I" & Application.WorksheetFunction.Match(LargestVolume, Range("L2:L" & Cells(Rows.Count, "L").End(xlUp).Row), 0) + 1).Value

'Label Bonus Cells
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("P2").Value = minTicker
Range("Q2").Value = GreatestDecrease
Range("P3").Value = maxTicker
Range("Q3").Value = GreatestIncrease
Range("P4").Value = maxVolTicker
Range("Q4").Value = LargestVolume

End Sub

