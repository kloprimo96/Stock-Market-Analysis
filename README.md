# Stock-Market-Analysis VBA_Challenge Code

  Sub StockMarketAnalysis():
  
    'Dimensions for Analysis'
  
   Dim ws As Worksheet
  
    Dim colummIndex As Integer
    Dim rowIndex As Long
   Dim yearlychange As Double
   Dim percentChange As Double
   Dim total As Double
   Dim start As Long
   Dim days As Integer
   Dim change As Double
   Dim dailyChange As Single
   Dim averageChange As Double
   Dim rowCount As Long
  
  For Each ws In Worksheets
      columNIndex = 0
      total = 0
      change = 0
      start = 2
      dailyChange = 0

    'Locations'
    
      ws.Range("I1").Value = "Ticker"
      ws.Range("J1").Value = "Yearly Change"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Stock Volume"
      ws.Range("O2").Value = "Greatest % Increase"
      ws.Range("O3").Value = "Greatest % Decrease"
      ws.Range("O4").Value = "Greatest Total Volumne"
      ws.Range("P1").Value = "Ticker"
      ws.Range("Q1").Value = "Value"
  
      rowCount = ws.Cells(rows.Count, "A").End(xlUp).row
  
      For rowIndex = 2 To rowCount
  
          If ws.Cells(rowIndex + 1, 1).Value <> ws.Cells(rowIndex, 1).Value Then
    
             total = total + ws.Cells(rowIndex, 7).Value
        
             If total = 0 Then
             
                ws.Range("I" & 2 + columNIndex).Value = Cells(rowIndex, 1).Value
                ws.Range("J" & 2 + columNIndex).Value = 0
                ws.Range("K" & 2 + columNIndex).Value = "%" & 0
                ws.Range("L" & 2 + columNIndex).Value = 0
            
            
             Else
                If ws.Cells(start, 3) = 0 Then
                    For find_value = start To rowIndex
                        If ws.Cells(find_value, 3).Value <> 0 Then
                           start = find_value
                           Exit For
                        End If
                    Next find_value
                End If
            
            change = (ws.Cells(rowIndex, 6) - ws.Cells(start, 3))
            percentChange = change / ws.Cells(start, 3)
            
            start = rowIndex + 1
            
            ws.Range("I" & 2 + columNIndex) = ws.Cells(rowIndex, 1).Value
            ws.Range("J" & 2 + columNIndex) = change
            ws.Range("J" & 2 + columNIndex).NumberFormat = "0.00"
            ws.Range("K" & 2 + columNIndex).Value = percentChange
            ws.Range("K" & 2 + columNIndex).NumberFormat = "0.00%"
            ws.Range("L" & 2 + columNIndex).Value = total
            
            Select Case change
                Case Is > 0
                    ws.Range("J" & 2 + columNIndex).Interior.ColorIndex = 4
                Case Is < 0
                    ws.Range("J" & 2 + columNIndex).Interior.ColorIndex = 3
                Case Else
                    ws.Range("J" & 2 + columNIndex).Interior.ColorIndex = 0
            End Select
            
            
          End If
          
        total = 0
        change = 0
        columNIndex = columNIndex + 1
        days = 0
        dailyChange = 0
          
     Else
         total = total + ws.Cells(rowIndex, 7).Value
         
          
     End If
      
                
      Next rowIndex
    
    
  ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
  ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
  ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
    
  increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
  decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:k" & rowCount)), ws.Range("K2:K" & rowCount), 0)
  volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
    
  ws.Range("P2") = ws.Cells(increase_number + 1, 9)
  ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
  ws.Range("P4") = ws.Cells(volume_number + 1, 9)
    
  Next ws
  
End Sub


