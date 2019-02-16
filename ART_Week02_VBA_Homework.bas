Attribute VB_Name = "Module1"
Sub StockTotals()

    For Each ws In Worksheets
     
        'Set column headers for new summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
        
        'Set a variable for holding the ticker name
        Dim Ticker_Name As String
        
        'Set a variable for holding the total per ticket
        Dim Ticker_Total As Double
        Ticker_Total = 0
        
        'Keep track of the location for each ticker label in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Determine last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through the values in Column 1
        For i = 2 To LastRow
            
            'Check if the ticker is repeated
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Set the ticker name
                Ticker_Name = ws.Cells(i, 1).Value
                
                'Add to the Ticker total
                Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
                
                'Print the ticker name in the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                'Print the ticker volume to the summary table
                ws.Range("J" & Summary_Table_Row).Value = Ticker_Total
                
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset the Ticker Total
                Ticker_Total = 0
                
            Else
            
                Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
                
            End If
            
        Next i
    
    Next ws

End Sub
