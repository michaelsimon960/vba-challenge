Attribute VB_Name = "Module1"
Sub stock_calc()

Dim ticker As String
Dim number_tickers As Integer
Dim lastRowState As Long
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim greatest_percent_increase As Double
Dim greatest_percent_increase_ticker As String
Dim greatest_percent_decrease As Double
Dim greatest_percent_decrease_ticker As String
Dim greatest_stock_volume As Double
Dim greatest_stock_volume_ticker As String



For Each ws In Worksheets

ws.Activate

lastRowState = ws.Cells(Rows.Count, "A").End(x1Up).Row

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

number_tickers = 0
ticker = ""
yearly_change = 0
opening_price = 0
percent_change = 0
total_stock_volume = 0

For i = 2 To lastRowState

ticker = Cells(i, 1).Value








End Sub
