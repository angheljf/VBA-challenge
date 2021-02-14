Sub stock_1()
    Dim total_vol As Double
    Dim ticker As String
    Dim ticker_counter, ticker_open_close_counter As Double
    Dim yearly_open, yearly_close As Double
    
    total_volume = 0
    ticker_counter = 2              ' Keep track of row
    ticker_open_close_counter = 2   ' Keep track of row for openning and closing values
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row

        total_volume = total_volume + Cells(i, 7).Value
        ticker = Cells(i, 1).Value
        yearly_open = Cells(ticker_open_close_counter, 3)
        
        'Summarize if ticker is different
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            yearly_close = Cells(i, 6)
            Cells(ticker_counter, 9).Value = ticker
            Cells(ticker_counter, 10).Value = yearly_close - yearly_open
            'set value to null to avoid divinding by zero
            If yearly_open = 0 Then
                Cells(ticker_counter, 11).Value = Null
            Else
                Cells(ticker_counter, 11).Value = (yearly_close - yearly_open) / yearly_open
            End If
            Cells(ticker_counter, 12).Value = total_volume
            
            ' Color the cells accordinly
            If Cells(ticker_counter, 10).Value > 0 Then
                Cells(ticker_counter, 10).Interior.ColorIndex = 4
            Else
                Cells(ticker_counter, 10).Interior.ColorIndex = 3
            End If
            'Format as percentages
            Cells(ticker_counter, 11).NumberFormat = "0.00%"
            
            
            total_volume = 0
            ticker_counter = ticker_counter + 1
            ticker_open_close_counter = i + 1
        End If
        
    Next i
    'AutoFit columns to make them look nice
    Columns("J").AutoFit
    Columns("K").AutoFit
    Columns("L").AutoFit

End Sub

