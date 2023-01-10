Sub Multiple_Year_Stock()

'Worksheet loop
For Each ws In Worksheets

    'Create a variable to hold file name, last row, and year
    Dim WorksheetName As String

    'Determine the last row
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Grabbed the Worksheetame
    WorksheetName = ws.Name
    
    'Split the WorksheetName
    Letter = Split(WorksheetName, "_")

      'Set an initial variable for holding the ticker name
        Dim Ticker As String

        'Set an initial variable for holding the volume per ticker
        Total_Stock_Volume = 0
        
        'Keep track of the location for each ticker in the table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Loop through all tickers
        For i = 2 To LastRow
        
        'Check if we are still within the same ticker
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            OpenValue = ws.Cells(i, 3).Value
            
            End If

            'Check if we are still within the same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                'Set the ticker name
                Ticker = ws.Cells(i, 1).Value
        
         'Calculate Total
             Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
                'Print the ticker name in the summary table
                ws.Cells(1, 9) = "Ticker"
    
                'Set a value for close
                CloseValue = ws.Cells(i, 6).Value
        
                'Calculate yearly change and put it in the table
                Yearly_Change = CloseValue - OpenValue
                ws.Cells(1, 10) = "Yearly Change"
                ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
        
                    'If statement to change cell color
                    If ws.Cells(Summary_Table_Row, 10).Value >= 0 Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                    Else
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                    End If
        
                'Calculate the yearly percent change and put it in the table
                Yearly_Percent_Change = Yearly_Change / OpenValue
                ws.Cells(1, 11) = "Yearly Percent Change"
                ws.Cells(Summary_Table_Row, 11).Value = FormatPercent(Yearly_Percent_Change)
        
                'Print total stock volume in table
                ws.Cells(1, 12) = "Total Stock Volume"
                ws.Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume
        
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                Total_Stock_Volume = 0
                
                
            Else
                
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                
                'Print ticker and value in table
                ws.Cells(1, 16) = "Ticker"
                ws.Cells(1, 17) = "Value"
                
                'Print % increase, decrease, and total volume in table
                ws.Cells(2, 15) = "Greatest % Increase"
                ws.Cells(3, 15) = "Greatest % Decrease"
                ws.Cells(4, 15) = "Greatest Total Volume"
    
                'Find the max for yearly percent change
                Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & LastRow)) * 100
                increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
                Range("P2") = Cells(increase_number + 1, 9)
                
                'Find the min for yearly percent change
                Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & LastRow)) * 100
                increase_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
                Range("P3") = Cells(increase_number + 1, 9)
                
                'Find the greatest total volume
                Range("Q4") = WorksheetFunction.Max(Range("L2:L" & LastRow))
                increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & LastRow)), Range("L2:L" & LastRow), 0)
                Range("P4") = Cells(increase_number + 1, 9)
                
            End If
    
    Next i
                
Next ws

End Sub
