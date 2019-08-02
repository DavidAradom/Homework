Attribute VB_Name = "Module1"
Sub StockData()
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

Dim ticker_symbol As String

Dim tickervolume As Double
tickervolume = 0

Dim summary As Integer
summary = 2

Range("I1").Value = "Ticker"
Range("J1").Value = "Total Volume"
lastrow = Cells(Rows.Count, 1).End(xlUp).row
  
For i = 2 To lastrow
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
ticker_symbol = Cells(i, 1).Value
tickervolume = tickervolume + Cells(i, 7).Value

Range("I" & summary).Value = ticker_symbol
Range("J" & summary).Value = tickervolume

summary = summary + 1
      
tickervolume = 0

Else

tickervolume = tickervolume + Cells(i, 7).Value

End If

Next i

Next ws

End Sub
