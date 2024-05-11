Sub stockData():

    Dim ticker As String
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim stock_volume As LongLong
    stock_volume = 0
    Dim open_price As Double
    
    Dim summaryTableRow As Integer
    summaryTableRow = 2
    
    Dim start_percentage As Double
    start_percentage = 0
    
    For Each ws In Worksheets
    
            'Set open price for the current worksheet
            open_price = ws.Range("C2").Value
            'Reset summary table row number
            summaryTableRow = 2
            
            'Assign summary table column headers
            ws.Range("I1") = "Ticker"
            ws.Range("J1") = "Yearly Change"
            ws.Range("K1") = "Percent Change"
            ws.Range("L1") = "Total Stock Volume"
            
            ws.Range("P1") = "Ticker"
            ws.Range("Q1") = "Value"
            ws.Range("O2") = "Greatest % Increase"
            ws.Range("O3") = "Greatest % Decrease"
            ws.Range("O4") = "Greatest Total Volume"
    
    'Get the last row count
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
            'Loop through all stocks for one year
            For k = 2 To (LastRow)
                If ws.Cells(k + 1, 1).Value <> ws.Cells(k, 1).Value Then
        
                    'Pull in values of interest
                    ticker = ws.Cells(k, 1).Value
                    yearly_change = ws.Cells(k, 6).Value - open_price
                    percentage_change = (ws.Cells(k, 6).Value - open_price) / open_price
                    stock_volume = stock_volume + ws.Cells(k, 7).Value
            
                    'Output ticker symbol
                    ws.Range("I" & summaryTableRow).Value = ticker
            
                    'Output yearly change
                    ws.Range("J" & summaryTableRow).Value = yearly_change
                    'Change the color depending on the percentage value
                    If yearly_change > 0 Then
                        ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
                    End If
                    
                    'Output percentage change (opening price to closing price)
                    ws.Range("K" & summaryTableRow).Value = percentage_change
                    ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"
             
                    'Output total stock volume of the stock
                    ws.Range("L" & summaryTableRow).Value = stock_volume
    
                    open_price = ws.Cells(k + 1, 3).Value
                    summaryTableRow = summaryTableRow + 1
                    stock_volume = 0
        
                Else
                    stock_volume = stock_volume + ws.Cells(k, 7).Value
                End If
                
            Next k
        
            
        'Return the percentage min and max
        Dim percentage_max As Double
        Dim percentage_min As Double
    
        percent_range = "K1:K" & LastRow
        volume_range = "L1:L" & LastRow
    
        percentage_max = WorksheetFunction.Max(ws.Range(percent_range))
        percentage_min = WorksheetFunction.Min(ws.Range(percent_range))

        ws.Range("Q2") = percentage_max
        ws.Range("Q3") = percentage_min
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
        greatest_percent_inc = WorksheetFunction.Match(ws.Range("Q2"), ws.Range("K:K"), 0)
        ws.Range("P2").Value = ws.Range("I" & greatest_percent_inc).Value
        greatest_percent_dec = WorksheetFunction.Match(ws.Range("Q3"), ws.Range("K:K"), 0)
        ws.Range("P3").Value = ws.Range("I" & greatest_percent_dec).Value
        
        
        
        'Return the greatest total volume
        Dim greatest_stock_volume As LongLong
    
        greatest_stock_volume = WorksheetFunction.Max(ws.Range(volume_range))
        ws.Range("Q4") = greatest_stock_volume
        
        greatest_vol_inc = WorksheetFunction.Match(ws.Range("Q4"), ws.Range("L:L"), 0)
        ws.Range("P4").Value = ws.Range("I" & greatest_vol_inc).Value
    Next ws
    
End Sub



