Attribute VB_Name = "Module1"
Sub AllSheets()
    
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call ParentMacros
    Next
    Application.ScreenUpdating = True

End Sub

Public Sub ParentMacros()

Call SummaryTable
Call TickerNameAndStockVolume
Call OpenPrice
Call YearlyChange
Call PercentChange


End Sub

Sub SummaryTable()

'Format table
Range("J1") = "Ticker"
Range("J1").Select
Selection.Interior.ColorIndex = 15

Range("K1") = "Opening Price"
Range("K1").Select
Selection.Interior.ColorIndex = 15

Range("L1") = "Closing Price"
Range("L1").Select
Selection.Interior.ColorIndex = 15

Range("M1") = "Yearly Change"
Range("M1").Select
Selection.Interior.ColorIndex = 15

Range("N1") = "Percent Change"
Range("N1").Select
Selection.Interior.ColorIndex = 15

Range("O1") = "Total Stock Volume"
Range("O1").Select
Selection.Interior.ColorIndex = 15

End Sub


Sub TickerNameAndStockVolume()

'Assign variables
Dim TickerName As String
Dim TickerTotal As Variant
Dim SummaryTableRow As Integer
SummaryTableRow = 2
Dim ClosePrice As Double

Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row


'List all tickers, capture closing price and compute for total stock volume
For i = 2 To LastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        TickerName = Cells(i, 1).Value
        ClosePrice = Cells(i, 6).Value
        Range("J" & SummaryTableRow).Value = TickerName
        Range("L" & SummaryTableRow).Value = ClosePrice
        Range("O" & SummaryTableRow).Value = TickerTotal
        SummaryTableRow = SummaryTableRow + 1
        TickerTotal = 0
    Else
        TickerTotal = TickerTotal + Cells(i, 7).Value
    End If
    
    Next i
End Sub

Sub OpenPrice()

'Assign variables
Dim SummaryTableRow As Integer
SummaryTableRow = 3
Dim OpenPrice As Double

Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row


'Capture opening price
[K2] = Cells(2, 3).Value
For i = 2 To LastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        OpenPrice = Cells(i + 1, 3).Value
        Range("K" & SummaryTableRow).Value = OpenPrice
        SummaryTableRow = SummaryTableRow + 1
    End If
Next i


End Sub

Sub YearlyChange()

Dim LastRow As Long
LastRow = Cells(Rows.Count, 10).End(xlUp).Row
Dim SummaryTableRow As Integer
SummaryTableRow = 2

'Compute for yearly change using formula: ClosePrice - OpenPrice

For i = 2 To LastRow
   Cells(i, 13).Value = (Range("L" & SummaryTableRow).Value) - (Range("K" & SummaryTableRow).Value)
   SummaryTableRow = SummaryTableRow + 1
Next i


End Sub

Sub PercentChange()

'Compute for percent change using formula: (ClosePrice-OpenPrice)/Total number of days

Dim LastRow As Long
LastRow = Cells(Rows.Count, 10).End(xlUp).Row

Dim SummaryTableRow As Variant
SummaryTableRow = 2

For i = 2 To LastRow
    If Cells(i, 11).Value <> 0 And Cells(i, 13).Value <> 0 Then
            Cells(i, 14).Value = (Range("M" & SummaryTableRow).Value) / (Range("K" & SummaryTableRow).Value)
            Range("N" & SummaryTableRow).NumberFormat = "0.00%"
            SummaryTableRow = SummaryTableRow + 1
    ElseIf Cells(i, 11).Value = 0 Or Cells(i, 13).Value = 0 Then
            Cells(i, 14).Value = 0
            Range("N" & SummaryTableRow).NumberFormat = "0.00%"
            SummaryTableRow = SummaryTableRow + 1
    Else
            Cells(i, 14).Value = 0
            Range("N" & SummaryTableRow).NumberFormat = "0.00%"
            SummaryTableRow = SummaryTableRow + 1
    End If
Next i


End Sub







'---------------------------------------------------------------------
'Steps To Do:
'1. List down all the tickers
'2. Add all the tickers
'3. Compute for yearly change using formula: EndDate-StartDate
    'Use a counter for this
'4. Compute for % change using formula: (EndDate-StartDate)/Total number of Days



