Sub Multiple_year()
    For Each ws In Worksheets
    Dim Ticker_name As String
    Dim N_Ticker As String
    Dim Open_value As Double
    Dim Close_vlaue As Double
    Dim Yearly_change As Double
    Dim Summary_stock_row As Integer
    Dim Percent_change As Double
    Dim Total_Stock_Volume As Double
    Dim Max As Double
    Dim Min As Double
    Dim Max_volume As Double
   
    
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    

    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Summary_stock_row = 2
    Total_Stock_Volume = 0
    Open_value = Cells(2, 3).Value
    For i = 2 To Last_Row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker_name = ws.Cells(i, 1).Value
            ws.Range("I" & Summary_stock_row).Value = Ticker_name
            Close_value = ws.Cells(i, 6).Value
            Yearly_change = Close_value - Open_value
            
            If Open_value <> 0 Then
                Percent_change = Yearly_change / Open_value
            Else
                Percent_change = 0
            End If
            
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            ws.Range("J" & Summary_stock_row).Value = Yearly_change
            ws.Range("K" & Summary_stock_row).Value = Format(Percent_change, "0.00%")
            ws.Range("L" & Summary_stock_row).Value = Total_Stock_Volume
            Summary_stock_row = Summary_stock_row + 1
            Open_value = ws.Cells(i + 1, 3).Value
            Total_Stock_Volume = 0
        Else
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
        End If
        Next i
        
            For i = 2 To Last_Row
                If ws.Cells(i, 11).Value < 0 Then
                    ws.Cells(i, 11).Interior.ColorIndex = 3
                ElseIf ws.Cells(i, 11).Value > 0 Then
                    ws.Cells(i, 11).Interior.ColorIndex = 4
                End If
            Next i
    
        Max = -1000
        For i = 2 To Last_Row
            If Max < ws.Cells(i, 11).Value Then
                Max = ws.Cells(i, 11).Value
                N_Ticker = ws.Cells(i, 9).Value
            End If
         Next i
         ws.Cells(2, 17).Value = Format(Max, "0.00%")
         ws.Cells(2, 16).Value = N_Ticker
         
        
        Min = 5000
        For i = 2 To Last_Row
            If Min > ws.Cells(i, 11).Value Then
                Min = ws.Cells(i, 11).Value
                N_Ticker = ws.Cells(i, 9).Value
            End If
        Next i
        ws.Cells(3, 17).Value = Format(Min, "0.00%")
        ws.Cells(3, 16).Value = N_Ticker
        
        Max_volume = 0
        For i = 2 To Last_Row
            If Max_volume < ws.Cells(i, 12).Value Then
                Max_volume = ws.Cells(i, 12).Value
                N_Ticker = ws.Cells(i, 9).Value
            End If
         Next i
         ws.Cells(4, 17).Value = Max_volume
         ws.Cells(4, 16).Value = N_Ticker
         
        
    Next ws

End Sub

