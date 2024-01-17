Sub stockfunction()

For Each ws In ActiveWorkbook.Worksheets

Dim ticker As String
Dim YearChange As Double
Dim openprice As Double
Dim closingprice As Double
Dim PercentChange As Double
Dim TSV As LongLong
TSV = 0
Dim SumTable As Integer
SumTable = 2



lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest TSV"

For i = 2 To lastrow
    
If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
openprice = ws.Cells(i, 3).Value
End If
      
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker = ws.Cells(i, 1).Value
TSV = TSV + ws.Cells(i, 7).Value
ws.Range("I" & SumTable).Value = ticker
ws.Range("L" & SumTable).Value = TSV
SumTable = SumTable + 1
TSV = 0
closingprice = ws.Cells(i, 6).Value
YearChange = closingprice - openprice
ws.Range("J" & SumTable - 1).Value = YearChange
If YearChange > 0 Then
ws.Range("J" & SumTable - 1).Interior.ColorIndex = 4
Else
ws.Range("J" & SumTable - 1).Interior.ColorIndex = 3
End If
PercentChange = (closingprice / openprice) - 1
ws.Range("K" & SumTable - 1).Value = PercentChange
ws.Range("K" & SumTable - 1).NumberFormat = "0.00%"
If PercentChange > 0 Then
ws.Range("K" & SumTable - 1).Interior.ColorIndex = 4
Else
ws.Range("K" & SumTable - 1).Interior.ColorIndex = 3
End If

Else
TSV = TSV + ws.Cells(i, 7).Value
End If

Next i

Dim max As Double
Dim min As Double
Dim maxtsv As LongLong
Dim maxTicker As String
Dim minTicker As String
Dim maxtsvticker As String

max = WorksheetFunction.max(ws.Range("K2:K" & lastrow))
maxTicker = WorksheetFunction.Index(ws.Range("I:I"), Application.Match(max, ws.Range("K2:K" & lastrow), 0) + 1)
ws.Cells(2, 17).Value = max
ws.Cells(2, 16).Value = maxTicker
ws.Cells(2, 17).NumberFormat = "0.00%"

min = WorksheetFunction.min(ws.Range("K2:K" & lastrow))
minTicker = WorksheetFunction.Index(ws.Range("I:I"), Application.Match(min, ws.Range("K2:K" & lastrow), 0) + 1)
ws.Cells(3, 17).Value = min
ws.Cells(3, 16).Value = minTicker
ws.Cells(3, 17).NumberFormat = "0.00%"

maxtsv = WorksheetFunction.max(ws.Range("L2:L" & lastrow))
maxtsvticker = WorksheetFunction.Index(ws.Range("I:I"), Application.Match(maxtsv, ws.Range("L2:L" & lastrow), 0) + 1)
ws.Cells(4, 17).Value = maxtsv
ws.Cells(4, 16).Value = maxtsvticker

Next ws

End Sub