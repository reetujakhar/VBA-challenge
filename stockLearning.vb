Sub StocksLearning()

 ' LOOP THROUGH ALL SHEETS
For Each wSheet In Worksheets
wSheet.Activate
       Cells(1, 9).Value = "Ticker"
       Cells(1, 10).Value = "Yearly Change"
       Cells(1, 11).Value = "Percent Change"
       Cells(1, 12).Value = "Total Stock Volume"
'declare variables

Dim ticker As String
Dim yearly_change As Double
Dim openprice As Double
Dim closeprice As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim last_row As Long

'initialize variables
total_stock_volume = 0
openprice = Cells(2, 3).Value
output_row = 2

' Determine the Last Row
     last_row = wSheet.Cells(Rows.Count, 1).End(xlUp).Row
     
      For i = 2 To last_row
      
       If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
       total_stock_volume = total_stock_volume + Cells(i, 7).Value
       closeprice = Cells(i, 6).Value
       yearly_change = closeprice - openprice
       percent_change = yearly_change / openprice
       Cells(output_row, 9).Value = Cells(i, 1).Value
       Cells(output_row, 10).Value = yearly_change
       Cells(output_row, 11).Value = percent_change
       Cells(output_row, 12).Value = total_stock_volume
               
               If yearly_change > 0 Then
                   Range("J" & output_row).Interior.Color = vbGreen
               ElseIf yearly_change < 0 Then
                   Range("J" & output_row).Interior.Color = vbRed
               Else
                   Range("J" & output_row).Interior.Color = vbWhite
               End If
' Resetting variable after change in ticker value

               total_stock_volume = 0
               openprice = Cells(i + 1, 3).Value
               output_row = output_row + 1
       
       Else
 'Counting total stock volume on similar ticker cell value
 
       total_stock_volume = total_stock_volume + Cells(i, 7).Value
              End If
              Next i
                  wSheet.Columns("J").NumberFormat = "$0.00"
    wSheet.Columns("K").NumberFormat = "0.00%"
    wSheet.Columns.AutoFit
              Next wSheet
              
              
     
End Sub

