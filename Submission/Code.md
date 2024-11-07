Sub StocksLoop()

        'Loop Ticker Symbol
        
        'Loop Quarterly Change from opening price to closing price
        
        'The total stock volume
        
        'variables
        Dim ticker As String
        Dim next_ticker As String
        Dim open_price As Double
        Dim closing_price As Double
        Dim volume As LongLong
        Dim volume_total As LongLong
        Dim quarterly_change As Double
        Dim percent_change As Double
        Dim rowCount As Long
        Dim i As LongLong
        Dim j As LongLong
        Dim k As LongLong
        Dim leaderboard_row As Integer
        
        'Set title row
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Quarterly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
                  
       'Reset per ticker
       volume_total = 0
       leaderboard_row = 2
       open_price = Cells(2, 3)
       
       ' get the row number of the last row with data
       rowCount = Cells(Rows.Count, "A").End(xlUp).Row
       Range("K12").Value = rowCount
    
       For i = 2 To rowCount
            ticker = Cells(i, 1).Value
            volume = Cells(i, 7)
            next_ticker = Cells(i + 1, 1).Value
            
            
            'if statement
            If (ticker <> next_ticker) Then
                'add total
                volume_total = volume_total + volume
                closing_price = Cells(i, 6)
                quarterly_change = closing_price - open_price
                percent_change = (quarterly_change / open_price)
                               

                            
                'write to leaderboard
                Cells(leaderboard_row, 12).Value = volume_total
                Cells(leaderboard_row, 9).Value = ticker
                Cells(leaderboard_row, 10).Value = quarterly_change
                Cells(leaderboard_row, 11).Value = percent_change
                Cells(leaderboard_row, 11).NumberFormat = "0.00%"
                
                'Conditional formatting
                If quarterly_change > 0 Then
                    Cells(leaderboard_row, 10).Interior.ColorIndex = 4 'green
                ElseIf quarterly_change < 0 Then
                    Cells(leaderboard_row, 10).Interior.ColorIndex = 3 'red
                End If
                
                'reset total
                volume_total = 0
                leaderboard_row = leaderboard_row + 1
                open_price = Cells(i + 1, 3)
            

            
            Else
                'add total
                volume_total = volume_total + volume
                quarterly_change = closing_price - open_price
                
        
                
            End If
        Next i
        
        Dim max As Double
        Dim min As Double
        Dim max_value As LongLong
        Dim max_row As Double
        Dim ticker_max As String
        Dim ticker_min As String
        Dim ticker_max_value As String
        
        max = 0
        min = 1
        max_value = 0
        
        For k = 2 To 1501
            If Cells(k, 11) > max Then
                ticker_max = Cells(k, 9).Value
                max = Cells(k, 11).Value
            End If
            If Cells(k, 11) < min Then
                ticker_min = Cells(k, 9).Value
                min = Cells(k, 11).Value
            End If
            If Cells(k, 12) > max_value Then
                ticker_max_value = Cells(k, 9).Value
                max_value = Cells(k, 12).Value
            End If
        Next k
        Cells(2, 17).Value = max
        Cells(2, 16).Value = ticker_max
        Cells(3, 17).Value = min
        Cells(3, 16).Value = ticker_min
        Cells(4, 17).Value = max_value
        Cells(4, 16).Value = ticker_max_value
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 17).NumberFormat = "0.00%"
End Sub
