Sub Labels()
    
    For Each ws In Worksheets
    
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
    
    Next ws
    
End Sub


Public Sub StockInfo()

For Each ws In Worksheets

'Set Variables
Dim Ticker As String

Dim Year_Open As Double
Dim Year_Close As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double


'Ticker Location
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2



'assign variable
Total_Stock_Volume = 0

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row



        For i = 2 To LastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                'where Ticker name appears
                ws.Range("J" & Summary_Table_Row).Value = Ticker
                
                'Where Volume appears
                ws.Range("M" & Summary_Table_Row).Value = Total_Stock_Volume
                
                
                
                'Calculations for yearly_Change and Percent_Change
                Year_Open = ws.Range("C" & Summary_Table_Row).Value
                Year_Close = ws.Range("F" & Summary_Table_Row).Value
                Yearly_Change = Year_Open - Year_Close
                
                    If Year_Open = 0 Then
                        Percent_Change = 0
                    Else
                        Percent_Change = Yearly_Change / Year_Open * 100
                    End If
                    
                 'where totals of yearly_change and Percent_change live
                 ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
                
                 
                 ws.Range("L" & Summary_Table_Row).Value = Percent_Change
                 ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
             
                        'Format-pos+gree,neg=red
                        If ws.Range("K" & Summary_Table_Row).Value > 0 Then
                           ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                        Else
                            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                        End If
                    
             
                    Summary_Table_Row = Summary_Table_Row + 1
             
                    Total_Stock_Volume = 0
             
                Else
                
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                End If
                
    
        Next i
    
  Next ws
  
End Sub

Public Sub Percent()
For Each ws In Worksheets

Dim Increase As Double
Dim Decrease As Double
Dim Ticker As String
Dim Greatest As Double

ws.Cells(2, 15).Value = "Greastest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"


Increase = ws.Cells(2, 12).Value
Decrease = ws.Cells(2, 12).Value
Greatest = ws.Cells(2, 13).Value

LastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row


For i = 2 To LastRow
 
 
   If ws.Cells(i, 13).Value > Greatest Then
   Greatest = ws.Cells(i, 13).Value
   ws.Cells(4, 17).Value = ws.Cells(i, 10).Value
   
   Else
    Greatest = Greatest
    
      
    End If
    
    If ws.Cells(i, 12).Value > Increase Then
    Increase = ws.Cells(i, 12).Value
    ws.Cells(2, 17).Value = ws.Cells(i, 10).Value
    
    Else
    
    Increase = Increase
    
    End If
    
    If ws.Cells(i, 12).Value < Decrease Then
    Decrease = ws.Cells(i, 12).Value
    ws.Cells(3, 17).Value = ws.Cells(i, 10).Value
    
    Else
    Decrease = Decrese
    
    End If
    
    ws.Cells(2, 17).Value = Format(Increase, "Percent")
    ws.Cells(3, 17).Value = Format(Decrease, "Percent")
    ws.Cells(4, 17).Value = Format(Greatest, "Scientific")
        
 Next i


Next ws


End Sub