Sub StocksLearningAnalysis()

 ' LOOP THROUGH ALL SHEETS
For Each wSheet In Worksheets
wSheet.Activate

       Cells(1, 16).Value = "Ticker"
       Cells(1, 17).Value = "Value"
       
              Cells(2, 15).Value = "Greatest % Increase"
       Cells(3, 15).Value = "Greatest % Decrease"
              Cells(4, 15).Value = "Greatest Total Volume"

       
'declare variables

Dim tickerTotalVolume As String
Dim tickerGreatIncrease As String
Dim tickerGreatDecrease As String
Dim greatTotalVolume As Double
Dim greatDecrease As Double
Dim greatIncrease As Double
Dim last_row As Long

'initialize variables
greatTotalVolume = 0
greatDecrease = 0
greatIncrease = 0

' Determine the Last Row
     last_row = wSheet.Cells(Rows.Count, 9).End(xlUp).Row
     
      For i = 2 To last_row
      If i <> last_row Then
            If Cells(i, 12).Value > greatTotalVolume Then
            tickerTotalVolume = Cells(i, 9).Value
      greatTotalVolume = Cells(i, 12).Value
       End If
             If Cells(i, 11).Value > greatIncrease Then
               tickerGreatIncrease = Cells(i, 9).Value
      greatIncrease = Cells(i, 11).Value
       End If
             If Cells(i, 11).Value < greatDecrease Then
               tickerGreatDecrease = Cells(i, 9).Value
      greatDecrease = Cells(i, 11).Value
       End If
       
      
      'Last row processing
            
      Else
      
            If Cells(i, 12).Value > greatTotalVolume Then
            tickerTotalVolume = Cells(i, 9).Value
      greatTotalVolume = Cells(i, 12).Value
       End If
             If Cells(i, 11).Value > greatIncrease Then
               tickerGreatIncrease = Cells(i, 9).Value
      greatIncrease = Cells(i, 11).Value
       End If
             If Cells(i, 11).Value < greatDecrease Then
               tickerGreatDecrease = Cells(i, 9).Value
      greatDecrease = Cells(i, 11).Value
       End If
       
       greatIncrease = 100 * greatIncrease
       greatDecrease = 100 * greatDecrease
               Cells(2, "P").Value = tickerGreatIncrease
               Cells(2, "Q").Value = "%" & greatIncrease
               Cells(3, "P").Value = tickerGreatDecrease
               Cells(3, "Q").Value = "%" & greatDecrease
               Cells(4, "P").Value = tickerTotalVolume
               Cells(4, "Q").Value = greatTotalVolume
      
      End If
      
      

              Next i


    wSheet.Columns.AutoFit
              Next wSheet
              
              
     
End Sub


