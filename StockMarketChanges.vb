Sub StockMarketChanges()
    ' -------------------------
    ' CREATE VARIABLES AND SET UP
    ' -------------------------
    For Each ws In Worksheets
    
        Dim WorksheetName As String
        ' ticker three letter id
        Dim ticker_id As String
        ' total stock volume for the year
        Dim stock_total As Double
        'price at open of the year
        Dim year_open As Double
        ' variable for total yearly change
        Dim yearly_change As Double
        'variable for percent change
        Dim percent_change As Double
        ' variable for summary table row
        Dim summary_table_row As Integer
    
        ' create column headers for first summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("I1:L1").Font.Bold = True
        
        'set first blank spot in table
        summary_table_row = 2
        'set stock total to 0
        stock_total = 0
        'set start row to 2, for the opening price of the current year
        year_open = 2
        
        'set variable for worksheet (autofit columns at end)
        WorksheetName = ws.Name
        
        ' find the last row of the data
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' --------------------------------------
        ' FIND YEARLY CHANGE, % CHANGE, AND TOTAL STOCK FOR EACH TICKER ID
        ' --------------------------------------
        
        ' loop through rows
        For row_ = 2 To last_row
            ' find where the current ticker code changes to the next
            If ws.Cells(row_ + 1, 1).Value <> ws.Cells(row_, 1).Value Then
                
                ' set the new ticker
                ticker_id = ws.Cells(row_, 1).Value
                
                'print ticker in table
                ws.Range("I" & summary_table_row).Value = ticker_id
                
                ' add stocks to stock total
                stock_total = stock_total + ws.Cells(row_, 7).Value
                
                'print stock total in table
                ws.Range("L" & summary_table_row).Value = stock_total
                
                ' find yearly change and print in table
                ws.Range("J" & summary_table_row).Value = ws.Cells(row_, 6).Value - ws.Cells(year_open, 3).Value
                
                    'conditional formatting
                    If ws.Range("J" & summary_table_row).Value > 0 Then
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                    End If
    
                ' find percent change and print in table
                ' create conditional to avoid division by 0
                If ws.Cells(year_open, 3).Value <> 0 Then
                    percent_change = (ws.Cells(row_, 6).Value - ws.Cells(year_open, 3).Value) / (ws.Cells(year_open, 3).Value)
                
                    'format percent change as a percentage
                    ws.Range("K" & summary_table_row).Value = Format(percent_change, "Percent")
                        'conditional formatting
                        If ws.Range("K" & summary_table_row).Value > 0 Then
                            ws.Range("K" & summary_table_row).Interior.ColorIndex = 4
                        Else
                            ws.Range("K" & summary_table_row).Interior.ColorIndex = 3
                        End If
                
                Else
                    ws.Range("K" & summary_table_row).Value = Format(0, "percent")
                End If
                
                'move to next blank row
                summary_table_row = summary_table_row + 1
                
                'reset stock total
                stock_total = 0
                
                'set new opening price
                year_open = row_ + 1
                
            'if current ticker is the same as previous, add to total
            Else
                stock_total = stock_total + ws.Cells(row_, 7).Value
            End If
        Next row_
        
        ' ------------------------------------------------
        ' FIND GREATEST % INCREASE, DECREASE, & STOCK VOLUME
        ' ------------------------------------------------
        
        ' find the last row of the created summary table
        last_summary_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' set up column and row headers for second summary table
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'set up first rows of the table for comparison
        greatest_incr = ws.Cells(2, 11).Value
        greatest_decr = ws.Cells(2, 11).Value
        greatest_vol = ws.Cells(2, 12).Value
        incr_ticker = ws.Cells(2, 9).Value
        decr_ticker = ws.Cells(2, 9).Value
        vol_ticker = ws.Cells(2, 9).Value
        
            'create for loop for summary
            For summary_row = 2 To last_summary_row
            
                ' find greatest percent increase
                If ws.Cells(summary_row, 11).Value > greatest_incr Then
                    greatest_incr = ws.Cells(summary_row, 11).Value
                    incr_ticker = ws.Cells(summary_row, 9).Value
                Else
                    greatest_incr = greatest_incr
                    incr_ticker = incr_ticker
                End If
                
                'find greatest percent decrease
                If ws.Cells(summary_row, 11).Value < greatest_decr Then
                    greatest_decr = ws.Cells(summary_row, 11).Value
                    decr_ticker = ws.Cells(summary_row, 9).Value
                Else
                    greatest_decr = greatest_decr
                    decr_ticker = decr_ticker
                End If
                
                 ' find greatest volume
                If ws.Cells(summary_row, 12).Value > greatest_vol Then
                    greatest_vol = ws.Cells(summary_row, 12).Value
                    vol_ticker = ws.Cells(summary_row, 9).Value
                Else
                    greatest_vol = greatest_vol
                    vol_ticker = vol_ticker
                End If
                
            'transfer summary results to table
            ws.Cells(2, 16).Value = incr_ticker
            ws.Cells(3, 16).Value = decr_ticker
            ws.Cells(4, 16).Value = vol_ticker
            ws.Cells(2, 17).Value = Format(greatest_incr, "Percent")
            ws.Cells(3, 17).Value = Format(greatest_decr, "Percent")
            ws.Cells(4, 17).Value = Format(greatest_vol, "Scientific")
            
            Next summary_row
            
        'autofit the columns
        Worksheets(WorksheetName).Range("A:Z").Columns.AutoFit
    Next ws
            
End Sub