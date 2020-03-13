Attribute VB_Name = "Module2"

Public Sub GreatestTable()

Dim Max As Double

Range("R2") = "Greatest % Increase"
Range("R3") = "Greatest % Decrease"
Range("R4") = "Greatest Total Volume"
Range("S1") = "Ticker"
Range("S1").Select
Selection.Interior.ColorIndex = 15
Range("T1") = "Value"
Range("T1").Select
Selection.Interior.ColorIndex = 15


Dim ws As Long
Dim i As Long
Dim LastRow As Long
LastRow = Cells(Rows.Count, 10).End(xlUp).Row
Dim MaxValue As Variant
Dim MaxTicker As String
Dim MinValue As Variant
Dim MinTicker As String
Dim MaxStockValue As Variant
Dim MaxStockTicker As String

For ws = 1 To 3
    MaxValue = 0
    MinValue = 0
    MaxStock = 0
    For i = 2 To LastRow
        If Worksheets(ws).Cells(i, 14).Value >= MaxValue Then
            MaxValue = Worksheets(ws).Cells(i, 14).Value
            MaxTicker = Worksheets(ws).Cells(i, 10).Value
        ElseIf Worksheets(ws).Cells(i, 14).Value <= MinValue Then
            MinValue = Worksheets(ws).Cells(i, 14).Value
            MinTicker = Worksheets(ws).Cells(i, 10).Value
        ElseIf Worksheets(ws).Cells(i, 15).Value >= MaxStockValue Then
             MaxStockValue = Worksheets(ws).Cells(i, 15).Value
             MaxStockTicker = Worksheets(ws).Cells(i, 10).Value
        End If
    Next i
Next ws

Range("T2") = MaxValue
Range("T2").NumberFormat = "0.00%"
Range("S2") = MaxTicker
Range("T3") = MinValue
Range("T3").NumberFormat = "0.00%"
Range("S3") = MinTicker
Range("T4") = MaxStockValue
Range("S4") = MaxStockTicker

 
 End Sub
 

