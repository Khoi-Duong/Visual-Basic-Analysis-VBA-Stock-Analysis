Sub Stock_Information()
    Columns("J").ColumnWidth = 13
    Columns("K").ColumnWidth = 14.5
    Columns("L").ColumnWidth = 18
    
    Dim ticker, ticker_symbol As Range
    Dim ticker_lastrow, ticker_symbol_lastrow, stock_volume, minimum_date, maximum_date, open_begyear, close_endyear, i, k As Long
    
    Set ticker = Range("A1")
    Set ticker_symbol = Range("I1")

    'Filter Unique Tickers
    Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ticker_symbol, Unique:=True
    Cells(1, 9).Value = "Ticker"
    
    ticker_lastrow = ticker.End(xlDown).Row
    ticker_symbol_lastrow = ticker_symbol.End(xlDown).Row
    
    
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Range("K:K").NumberFormat = "0.00%"
    Cells(1, 12).Value = "Total Stock Volume"
    Range("L:L").NumberFormat = "0"

    'Calculating Yearly Change and Percent Change For Each Unique Ticker
    For i = 2 To ticker_symbol_lastrow
        minimum_date = 0
        maximum_date = 0
        open_begyear = 0
        close_endyear = 0
        For k = 2 To ticker_lastrow
            If Cells(i, 9).Value = Cells(k, 1).Value Then
                If minimum_date = 0 Then
                    minimum_date = Cells(k, 2).Value
                    open_begyear = Cells(k, 3).Value
                ElseIf maximum_date = 0 Then
                    maximum_date = Cells(k, 2).Value
                ElseIf maximum_date < Cells(k, 2).Value Then
                    maximum_date = Cells(k, 2).Value
                    close_endyear = Cells(k, 6).Value
                End If
                'Calculating Stock Volume For Each Unique Ticker
                stock_volume = Cells(i, 12).Value + Cells(k, 7).Value
                Cells(i, 12).Value = stock_volume
            End If
        Next k
        
        Cells(i, 10).Value = close_endyear - open_begyear
        
        'Conditional Formating For Yearly Change Column
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.Color = RGB(0, 255, 0)
        ElseIf Cells(i, 10).Value < 0 Then
            Cells(i, 10).Interior.Color = RGB(255, 0, 0)
        End If
        
        Cells(i, 11).Value = ((close_endyear / open_begyear) - 1)
        
    Next i

    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("O:O").Columns.AutoFit
    
    Dim percentchange, stockvolume As Range
    Dim percentchange_lastrow, stockvolume_lastrow, greatest_increase, greatest_decrease, greatest_volume, o, s As Long
    
    Set percentchange = Range("K1")
    Set stockvolume = Range("L1")
    
    percentchange_lastrow = percentchange.End(xlDown).Row
    stockvolume_lastrow = stockvolume.End(xlDown).Row
    
    greatest_increase = 0
    greatest_decrease = 0

    'Finding Ticker With The Greatest % Increase/Decrease
    For o = 2 To percentchange_lastrow
        If Cells(o, 11).Value > greatest_increase Then
            greatest_increase = Cells(o, 11).Value
            Cells(2, 16).Value = Cells(o, 9).Value
        ElseIf Cells(o, 11).Value < greatest_decrease Then
            greatest_decrease = Cells(o, 11).Value
            Cells(3, 16).Value = Cells(o, 9).Value
        End If
    Next o
            
    Cells(2, 17).Value = greatest_increase
    Cells(3, 17).Value = greatest_decrease
    
    greatest_volume = 0
    'Finding Ticker With The Greatest Volume
    For s = 2 To stockvolume_lastrow
        If Cells(s, 12).Value > greatest_volume Then
            greatest_volume = Cells(s, 12).Value
            Cells(4, 16).Value = Cells(s, 9).Value
        End If
    Next s
    Cells(4, 17).Value = greatest_volume

End Sub

