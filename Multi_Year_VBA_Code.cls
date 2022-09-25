VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub sheet_loop1()

Dim wb As Workbook
Set wb = ActiveWorkbook

Dim headings_list() As Variant

headings_list() = Array("ticker", "date", "open", "High", "Low", "close", "Volume", " ", "Ticker", "YearlyChange", "Percentchange", "StockVolume ", " ", " ", " ", "Ticker", "value ")


Dim ws As Worksheet


For Each ws In wb.Sheets

     With ws
     .Rows(1).Value = ""
     For i = LBound(headings_list()) To UBound(headings_list())

 .Cells(1, 1 + i).Value = headings_list(i)
 

Next i

.Rows(1).Font.Bold = True
.Rows(1).VerticalAlignment = xlCenter
End With
Next ws



For Each ws In wb.Sheets


Dim summary_table_row As Long

summary_table_row = 2


Dim lastrow As Long
Dim ticker_name As String

Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim total_volume As Double
 
 total_volume = 0

Dim percent_change As Double
 
 percent_change = 0

Dim min_increase As Double
Dim max_increase As Double
Dim max_ticker_name As String
Dim min_ticker_name1 As String

Dim max_volume As Double
Dim max_volume_ticker_name As String

max_volume = 0
min_volume = 0
max_increase = 0


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
open_price = ws.Cells(2, 3).Value

For i = 2 To lastrow
'this is for calculating total stock volume


If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
'this is for calculating total for each ticker.
 total_volume = total_volume + ws.Cells(i, 7).Value
 
' this has to be printed for first ticker
 ticker_name = ws.Cells(i, 1).Value
 
 Else

total_volume = total_volume + ws.Cells(i, 7).Value
ws.Range("I" & summary_table_row).Value = ticker_name

End If

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
  
 ticker_name = ws.Cells(i, 1).Value
 close_price = ws.Cells(i, 6).Value
 yearly_change = close_price - open_price
 
If open_price <> 0 Then
  percent_change = (yearly_change / open_price) * 100
 
End If
 
  
    ws.Range("J" & summary_table_row).Value = yearly_change
 
 'for yearly chage calculation
 
  If (yearly_change > 0) Then
  
  ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
  
  ElseIf (yearly_change <= 0) Then
  
  ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
  
  End If
  
ws.Range("K" & summary_table_row).Value = (CStr(percent_change) & " % ")

ws.Range("L" & summary_table_row).Value = total_volume

summary_table_row = summary_table_row + 1

open_price = ws.Cells(i + 1, 3).Value

'calculating percent change


If (percent_change > max_increase) Then
 max_increase = percent_change
 max_ticker_name = ticker_name

 ElseIf (percent_change < min_increase) Then
 min_increase = percent_change
  ticker_name = ws.Cells(i, 1).Value

 min_ticker_name1 = ticker_name

End If

If (total_volume > max_volume) Then

max_volume = total_volume
max_volume_ticker_name = ticker_name

End If

' reset values

total_volume = 0
percent_change = 0



End If


Next i



ws.Range("Q2").Value = (CStr(max_increase) & " % ")

ws.Range("Q3").Value = (CStr(min_increase) & " % ")
ws.Range("p2").Value = max_ticker_name
ws.Range("p3").Value = min_ticker_name1

ws.Range("p4").Value = max_volume_ticker_name

ws.Range("Q4").Value = max_volume
ws.Range("O2").Value = " Greatest % increase "
ws.Range("O3").Value = " Greatest % decrease "
ws.Range("O4").Value = "Greatest total volume "


Next ws

max_ticker_name = ""
min_ticker_name1 = ""
End Sub

