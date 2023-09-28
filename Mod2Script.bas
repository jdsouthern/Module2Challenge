Attribute VB_Name = "Module1"
Sub stockSummary():
    
  For Each ws In Worksheets

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim tickerCol As String
    Dim totalVol As Double
    Dim summaryRows As Integer
    Dim openPrice As Single
    Dim closePrice As Single
    Dim yearlyChange As Double
    Dim percentChange As Single
    
    ws.Range("I1").Value = ("Ticker")
    ws.Range("J1").Value = ("Yearly Change")
    ws.Range("K1").Value = ("Percent Change")
    ws.Range("L1").Value = ("Total Stock Volume")
    
    summaryRows = 2 'initialize summary table rows
    openPrice = ws.Cells(2, 3).Value 'initialize open price
    
    For Row = 2 To lastRow
      
        If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then 'look if next ticker is different than current
            ws.Cells(summaryRows, 9).Value = ws.Cells(Row, 1).Value 'set ticker in table
            ws.Cells(summaryRows, 12).Value = totalVol 'set yearly volume in table
            closePrice = ws.Cells(Row, 6).Value 'set close price of current ticker block
            yearlyChange = closePrice - openPrice
            yearlyChange = WorksheetFunction.Round(yearlyChange, 2) 'Round off yearly change value
            ws.Cells(summaryRows, 10).Value = yearlyChange 'set yearly change in table
            If ws.Cells(summaryRows, 10).Value >= 0 Then
                ws.Cells(summaryRows, 10).Interior.ColorIndex = 4
                Else
                ws.Cells(summaryRows, 10).Interior.ColorIndex = 3
            End If
            percentChange = yearlyChange / openPrice
            ws.Cells(summaryRows, 11).Value = percentChange 'set percent change in table
            ws.Cells(summaryRows, 11).NumberFormat = "0.00%"
            ws.Columns("I:L").AutoFit
            totalVol = 0 'reset total vol
            summaryRows = summaryRows + 1 'move to next table row
            openPrice = ws.Cells(Row + 1, 3).Value 'set open price of new ticker block
            
        Else
            totalVol = totalVol + ws.Cells(Row, 7).Value
        End If
        
    Next Row
    
    sumLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    Dim maxPercentIncrease As Single
    Dim maxPercentDecrease As Single
    Dim maxVolume As Double
    
    ws.Range("O2").Value = ("Greatest % Increase")
    ws.Range("O3").Value = ("Greatest % Decrease")
    ws.Range("O4").Value = ("Greatest Total Volume")
    ws.Columns("O").AutoFit
    
    ws.Range("P1").Value = ("Ticker")
    ws.Range("Q1").Value = ("Value")
   
    For Row = 2 To sumLastRow
    
        If ws.Cells(Row + 1, 11).Value > maxPercentIncrease Then
            maxPercentIncrease = ws.Cells(Row + 1, 11).Value
        End If
            ws.Cells(2, 17).Value = maxPercentIncrease
            ws.Cells(2, 17).NumberFormat = "0.00%"
        
        If ws.Cells(Row, 11).Value = maxPercentIncrease Then
            ws.Cells(2, 16) = ws.Cells(Row, 9).Value
        End If
        '''''''''''''''''''
        If ws.Cells(Row + 1, 11).Value < maxPercentDecrease Then
            maxPercentDecrease = ws.Cells(Row + 1, 11).Value
        End If
            ws.Cells(3, 17).Value = maxPercentDecrease
            ws.Cells(3, 17).NumberFormat = "0.00%"
        
        If ws.Cells(Row, 11).Value = maxPercentDecrease Then
            ws.Cells(3, 16) = ws.Cells(Row, 9).Value
        End If
        '''''''''''''''''''
        If ws.Cells(Row + 1, 12).Value > maxVolume Then
            maxVolume = ws.Cells(Row + 1, 12).Value
        End If
            ws.Cells(4, 17).Value = maxVolume
        
        If ws.Cells(Row, 12).Value = maxVolume Then
            ws.Cells(4, 16) = ws.Cells(Row, 9).Value
        End If
        
    Next Row
    
    maxPercentIncrease = 0
    maxPercentDecrease = 0
    maxVolume = 0
    
  Next ws
            
End Sub
