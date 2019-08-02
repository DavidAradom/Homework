Attribute VB_Name = "Module1"
Sub StockData()
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row

Range("j1").Value = "Ticker"
Range("k1").Value = "Yearly Change"
Range("l1").Value = "Percent Change"
Range("m1").Value = "Total Volume"

Dim ticker_symbol As String
Dim openprice As Double
Dim closedprice As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim tickervolume As Double


tickervolume = 0


Dim row As Double

row = 2


Dim j As Integer
j = 1
Dim i As Long


openprice = Cells(2, j + 3).Value

For i = 2 To lastrow

If Cells(i + 1, j).Value <> Cells(i, j).Value Then

ticker_symbol = Cells(i, 1).Value
Cells(row, j + 9).Value = ticker_symbol

closedprice = Cells(i, 6).Value

yearlychange = closedprice - openprice
Cells(row, j + 10).Value = yearlychange

If (openprice = 0 And closedprice = 0) Then
percentchange = 0

ElseIf (openprice = 0 And closedprice <> 0) Then
percentchange = 1

Else: percentchange = yearlychange / openprice
Cells(row, j + 11).Value = percentchange
Cells(row, j + 11).NumberFormat = "0.00%"

End If

tickervolume = tickervolume + Cells(i, 7).Value
Cells(row, j + 12).Value = tickervolume

row = row + 1

openprice = Cells(i + 1, j + 3)
tickervolume = 0

Else
tickervolume = tickervolume + Cells(i, 7).Value
End If

Next i

ycbottomrow = ws.Cells(Rows.Count, j + 10).End(xlUp).row

For c = 2 To ycbottomrow
If (Cells(c, j + 10).Value > 0 Or Cells(c, j + 10).Value = 0) Then
Cells(c, j + 10).Interior.ColorIndex = 10
ElseIf Cells(c, j + 10).Value < 0 Then
Cells(c, j + 10).Interior.ColorIndex = 3

End If

Next c



Next ws

End Sub
