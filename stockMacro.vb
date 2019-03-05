Sub Stocks()
    'Loop through each worksheet
    For Each ws In Worksheets
        
        'Create variable to hold ticker symbol
        Dim ticker As String
        
        'Create variable to calculate the total per ticker
        Dim ticker_volume As Double
        ticker_volume = 0
        
        'Keep track of the location of each ticker in Ticker column
        Dim ticker_row As Double
        ticker_row = 2
        
        'Keep track of opening stock price
        Dim start_price As Double
        
        'Keep track of ending stock price
        Dim end_price As Double
        
        'Keep track of greatest % increase
        Dim greatest_increase As Double
        
        'Keep track of greatest % decrease
        Dim greatest_decrease As Double
        
        'Keep track of greatest total volume
        Dim greatest_volume As Double
         
        'Add "Ticker" header to Column I
        ws.Cells(1, 9).Value = "Ticker"
        
        'Add "Yearly Change header to Column J
        ws.Cells(1, 10).Value = "Yearly Change"
        
        'Add "Percent Change" header to Column K
        ws.Cells(1, 11).Value = "Percent Change"
        
        'Add "Total Stock Volume" header to Column L
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Set up Greatest Increase/Decrease/Total Volume Table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Set the row number for initial greatest_increase, greatest_decrease_ and greatest_volume
        greatest_increase = 2
        greatest_decrease = 2
        greatest_volume = 2
        
            
        'Loop through all data rows
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            'Check to see if we're in the first instance of a ticker
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set the start price
                start_price = ws.Cells(i, 3).Value
                                              
            'Check if we're still within the same stock ticker
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Set the ticker
                ticker = ws.Cells(i, 1).Value
            
                'Add to the ticker volume total
                ticker_volume = ticker_volume + ws.Cells(i, 7).Value
            
                'Print the ticker in the summary table
                ws.Cells(ticker_row, 9).Value = ticker
            
                'Print the total volume amount in the summary table
                ws.Cells(ticker_row, 12).Value = ticker_volume
                
                'Set the ticker end price
                end_price = ws.Cells(i, 6).Value
                
                'Caluclate yearly change and input in summary table
                ws.Cells(ticker_row, 10).Value = end_price - start_price
                
                'Add conditional formatting. If value greater than 0 fill with green. If less than 0 fill with red.
                If ws.Cells(ticker_row, 10).Value > 0 Then
                    ws.Cells(ticker_row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(ticker_row, 10).Interior.ColorIndex = 3
                End If
                
                'Calulate percent change and input in summary table
                If start_price = 0 Then
                    ws.Cells(ticker_row, 11).Value = 0
                    ws.Cells(ticker_row, 11).NumberFormat = "0.00%"
                Else
                    ws.Cells(ticker_row, 11).Value = (end_price - start_price) / start_price
                    ws.Cells(ticker_row, 11).NumberFormat = "0.00%"
                End If
            
                'Add one to the summary table row
                ticker_row = ticker_row + 1
            
                'Reset the ticker_volume
                ticker_volume = 0
                                
                
            'If the ticker immediately following a row is the same ticker...
            Else
            
                'Add to the ticker_volume
                ticker_volume = ticker_volume + ws.Cells(i, 7).Value
                
            End If
                          
        Next i
        
        
        'Find the Greatest % Inrease, Greatest % Decrease
        
        For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
            If ws.Cells(i, 11).Value > ws.Cells(greatest_increase, 11).Value Then
                greatest_increase = i
            ElseIf ws.Cells(i, 11).Value < ws.Cells(greatest_decrease, 11).Value Then
                greatest_decrease = i
            End If
               
            'Find the Greatest Total Volume
            If ws.Cells(i, 12).Value > ws.Cells(greatest_volume, 12).Value Then
                greatest_volume = i
            End If
        
        Next i
                    
        'Input values into greatest % increase/decrease and greatest total volume
        ws.Cells(2, 16).Value = ws.Cells(greatest_increase, 9).Value
        ws.Cells(2, 17).Value = ws.Cells(greatest_increase, 11).Value
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = ws.Cells(greatest_decrease, 9).Value
        ws.Cells(3, 17).Value = ws.Cells(greatest_decrease, 11).Value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = ws.Cells(greatest_volume, 9).Value
        ws.Cells(4, 17).Value = ws.Cells(greatest_volume, 12).Value
        
    Next ws
    
End Sub



