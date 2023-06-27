VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stocks()

'1
Dim ws As Worksheet

'2
For Each ws In Worksheets

'3
ws.Range("A1").CurrentRegion.Sort key1:=ws.Range("A1"), order1:=xlAscending, Header:=xlYes

'4
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'5
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'6
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'7
Dim ticker As String

Dim tickervolume As Double
tickervolume = 0

'8
Dim summary_ticker_row As Double
summary_ticker_row = 2

'9
Dim open_price As Double

'10
Dim close_price As Double

'11
Dim yearly_change As Double

'12
Dim percent_change As Double

'13
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'14
For i = 2 To lastrow
  
'15
ticker = ws.Cells(i, 1).Value

'16
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'17
tickervolume = tickervolume + ws.Cells(i, 7).Value

'18
ws.Range("i" & summary_ticker_row).Value = ticker

'19
close_price = ws.Cells(i, 6).Value

'20
open_price = ws.Cells(i, 3).Value

'21
yearly_change = (close_price) - (open_price)

'22
ws.Range("J" & summary_ticker_row).Value = yearly_change

'23
If yearly_change > 0 Then

'24
ws.Range("J" & summary_ticker_row).Interior.ColorIndex = 4

Else

'25
ws.Range("J" & summary_ticker_row).Interior.ColorIndex = 3

End If

'26
ws.Range("L" & summary_ticker_row).Value = tickervolume

'27
If open_price = 0 Then

percent_change = 0

Else

percent_change = ((yearly_change) / (open_price)) * 100

End If

'28
ws.Range("K" & summary_ticker_row).Value = percent_change

'29
summary_ticker_row = summary_ticker_row + 1

'30
tickervolume = 0

'31
open_price = ws.Cells(i + 1, 3)

Else

'32
tickervolume = tickervolume + ws.Cells(i, 7).Value

End If

Next i

Next ws
    
End Sub
