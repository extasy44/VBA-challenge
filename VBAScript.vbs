Sub StockMarketAnalysis3()
  For Each ws In Worksheets
    Dim YearlyChangeOpen As Double
    Dim YearlyChangeClose As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    
    Dim SummaryRow As Double
    Dim Row As Double
    Dim RowCount As Double
 
    Dim GreatestIncrease, GreatestDecrease, GreatestTotalVolume As Double
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestTotalVolume = 0
        
     'Set heading and labels   
     ws.Cells(1, 9).Value = "Ticker"
     ws.Cells(1, 10).Value = "Yearly Change"
     ws.Cells(1, 11).Value = "Percent Change"
     ws.Cells(1, 12).Value = "Total Stock Volume"
     
     ws.Cells(1, 16).Value = "Ticker"
     ws.Cells(1, 17).Value = "Value"
     ws.Cells(2, 15).Value = "Greatest & Increase"
     ws.Cells(3, 15).Value = "Greatest & Decrease"
     ws.Cells(4, 15).Value = "Greatest Total Volume"

    'initialize first values
    SummaryRow = 2
    ws.Cells(SummaryRow, 9).Value = ws.Cells(SummaryRow, 1).Value
    RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
    YearlyChangeOpen = ws.Cells(2, 3).Value
    
    For Row = 2 To RowCount
        
        If ws.Cells(Row, 1).Value = ws.Cells(SummaryRow, 9) Then
            TotalVolume = TotalVolume + ws.Cells(Row, 7).Value
         Else
            YearlyChangeClose = ws.Cells(Row - 1, 6)
            YearlyChange = YearlyChangeClose - YearlyChangeOpen
            If YearlyChange = 0 Then
                PercentChange = 0
            ElseIf YearlyChangeOpen = 0 And YearlyChangeClose <> 0 Then
                PercentChange = 1
            Else
                PercentChange = YearlyChange / YearlyChangeOpen
            End If
            ws.Cells(SummaryRow, 10).Value = YearlyChange
            ws.Cells(SummaryRow, 11).Value = PercentChange
            ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
            If PercentChange < 0 Then
                ws.Cells(SummaryRow, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(SummaryRow, 11).Interior.ColorIndex = 4
            End If
            
            If PercentChange > GreatestIncrease Then
                GreatestIncrease = PercentChange
                ws.Cells(2, 16).Value = ws.Cells(Row - 1, 1).Value
                ws.Cells(2, 17).Value = GreatestIncrease
                
            ElseIf PercentChange < GreatestDecrease Then
                GreatestDecrease = PercentChange
                ws.Cells(3, 16).Value = ws.Cells(Row - 1, 1).Value
                ws.Cells(3, 17).Value = GreatestDecrease
            End If
            
            If TotalVolume > GreatestTotalVolume Then
                GreatestTotalVolume = TotalVolume
                ws.Cells(4, 16).Value = ws.Cells(Row - 1, 1).Value
                ws.Cells(4, 17).Value = GreatestTotalVolume
            End If
            
            'reset value for next iteration
            YearlyChangeOpen = ws.Cells(Row, 6)
            ws.Cells(SummaryRow, 12).Value = TotalVolume
            TotalVolume = ws.Cells(Row, 7).Value
            SummaryRow = SummaryRow + 1
            ws.Cells(SummaryRow, 9).Value = ws.Cells(Row, 1).Value
         End If
         
     Next Row
     
     
     'assign values for the last Summary row
     YearlyChange = YearlyChangeClose - YearlyChangeOpen
     ws.Cells(SummaryRow, 10).Value = YearlyChange
      If YearlyChange = 0 Then
          PercentChange = 0
      ElseIf YearlyChangeOpen = 0 And YearlyChangeClose <> 0 Then
          PercentChange = 1
      Else
          PercentChange = YearlyChange / YearlyChangeOpen
      End If
      ws.Cells(SummaryRow, 11).Value = PercentChange
      ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
      If PercentChange < 0 Then
          ws.Cells(SummaryRow, 11).Interior.ColorIndex = 3
      Else
          ws.Cells(SummaryRow, 11).Interior.ColorIndex = 4
      End If
      ws.Cells(SummaryRow, 12).Value = TotalVolume
      
      If PercentChange < 0 Then
          ws.Cells(SummaryRow, 11).Interior.ColorIndex = 3
      Else
          ws.Cells(SummaryRow, 11).Interior.ColorIndex = 4
        End If
              
        If PercentChange > GreatestIncrease Then
          GreatestIncrease = PercentChange
          ws.Cells(2, 16).Value = ws.Cells(Row - 1, 1).Value
          ws.Cells(2, 17).Value = GreatestIncrease
                  
        ElseIf PercentChange < GreatestDecrease Then
          GreatestDecrease = PercentChange
          ws.Cells(3, 16).Value = ws.Cells(Row - 1, 1).Value
          ws.Cells(3, 17).Value = GreatestDecrease
        End If
              
        If TotalVolume > GreatestTotalVolume Then
          GreatestTotalVolume = TotalVolume
          ws.Cells(4, 16).Value = ws.Cells(Row - 1, 1).Value
          ws.Cells(4, 17).Value = GreatestTotalVolume
        End If
              
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
            
    Next ws

End Sub

