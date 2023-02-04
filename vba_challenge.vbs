Sub mandhir_bajaj()

For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    
    ws.Cells(1, 10).Value = "Yearly Change"
    
    ws.Cells(1, 11).Value = "Percent Change"
    
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Cells(1, 15).Value = ws.Cells(1, 9).Value
    
    ws.Cells(1, 16).Value = "Value"
    
    ws.Cells(2, 14).Value = "Greatest % Increase"
    
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    Dim i, j, row1, row2, ticker  As Integer
    
    Dim day_start, day_end As Long
    
    Dim ticker_name As String
    
    ticker = 2
    
    day_start = 2
    
    row1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To row1
        
    vol = vol + ws.Cells(i, 7).Value
    
    ws.Cells(i, 11).NumberFormat = "0.00%"
        
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
        ws.Cells(ticker, 9).Value = ws.Cells(i, 1).Value
        
        
        
        day_end = i
        
        row2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
        ws.Cells(ticker, 10).Value = ws.Cells(day_end, 6).Value - ws.Cells(day_start, 3)
            
        If ws.Cells(ticker, 10).Value >= 0 Then
            
            ws.Cells(ticker, 10).Interior.ColorIndex = 4
                
        Else
            ws.Cells(ticker, 10).Interior.ColorIndex = 3
            
        End If
            
        ws.Cells(ticker, 11).Value = ws.Cells(ticker, 10).Value / ws.Cells(day_start, 3).Value
            
        ws.Cells(ticker, 12).Value = vol
            
        ticker = ticker + 1
            
        vol = 0
        
        day_start = day_end + 1
            
    End If
                
    Next i
    
    greatest_increase = 0
    
    greateat_decrease = 0
    
    greatest_total_volume = 0
    
    For j = 2 To row2
    
        If ws.Cells(j, 11).Value >= greatest_increase Then
        
        greatest_increase = ws.Cells(j, 11).Value
        
        ws.Cells(2, 15).Value = ws.Cells(j, 9).Value
        
        ws.Cells(2, 16).Value = greatest_increase
        
        End If
        
        If ws.Cells(j, 11).Value <= greatest_decrease Then
        
        greatest_decrease = ws.Cells(j, 11).Value
        
        ws.Cells(3, 15).Value = ws.Cells(j, 9).Value
        
        ws.Cells(3, 16).Value = greatest_decrease
        
        End If
        
        If ws.Cells(j, 12).Value >= greatest_total_volume Then
        
        greatest_total_volume = ws.Cells(j, 12).Value
        
        ws.Cells(4, 15).Value = ws.Cells(j, 9).Value
        
        ws.Cells(4, 16).Value = greatest_total_volume
        
        End If
    
    Next j

Next ws

End Sub

