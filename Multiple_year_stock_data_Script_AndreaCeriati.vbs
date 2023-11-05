
'Create a script that loops through all the stocks for one year and outputs the following information:

'The ticker symbol

'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

'The total stock volume of the stock.

'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

'Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

'NOTE: Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

'----------------------------------------------------------------------------------------------------------------------------------------

Sub stock()

    'Loop through all sheets
    For Each ws In Worksheets
    
        'Define the variables
        Dim Stock_Ticker As String
        Dim Opening_Price As Double
        Dim Closing_Price As Double
        Dim Volume_Total As LongLong
        Dim Summary_Table_Row As Integer
        Dim Start As Double
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Max_Volume As Double
        
        'Define the inintial values for the variables below
        Summary_Table_Row = 2
        Start = 2
        Volume_Total = 0
        
        'Create all the headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change ($)"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Find the last row in the sheets
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
            
        'Loop through all the stock tickers
        For i = 2 To lastRow
                
            'Check if we are still in the same stock ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'Set the ticker name, opening price, closing price, total volume
                Stock_Ticker = ws.Cells(i, 1).Value
                Opening_Price = ws.Cells(Start, 3).Value
                Closing_Price = ws.Cells(i, 6).Value
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
            
                'Print the stock ticker, yearly change from the opening price at the beginning of a given year to the closing price at the end
                'of that year, the total volume, and the percentage change from the opening price at the beginning of a given year to the
                'closing price at the end of that year in the summary table
                ws.Cells(Summary_Table_Row, 9) = Stock_Ticker
                ws.Cells(Summary_Table_Row, 10) = Closing_Price - Opening_Price
                ws.Cells(Summary_Table_Row, 12) = Volume_Total
                ws.Cells(Summary_Table_Row, 11) = (Closing_Price - Opening_Price) / Opening_Price
                
                'Add one to the summary table row, set the new value for start, and re-set total volume
                Summary_Table_Row = Summary_Table_Row + 1
                Start = i + 1
                Volume_Total = 0
                
            Else
                
                'Add to the total volume
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                
            End If
        
        Next i
            
        'Re-set summary table row value
        Summary_Table_Row = 2
            
        'Find the last row in the summary table
        lastRow2 = ws.Cells(Rows.Count, "I").End(xlUp).Row
            
        'Set greatest increase, greatest decrease, and max volume
        Greatest_Increase = WorksheetFunction.Max(ws.Range("K2:K" & lastRow2))
        Greatest_Decrease = WorksheetFunction.Min(ws.Range("K2:K" & lastRow2))
        Max_Volume = WorksheetFunction.Max(ws.Range("L2:L" & lastRow2))
            
        'Loop through all the stock tickers
        For i = 2 To lastRow2
            
            'Check the ticker with the greatest increase
            If ws.Cells(i, 11).Value = Greatest_Increase Then
                
                'Print ticker and value in new summary table
                ws.Cells(Summary_Table_Row, 16) = ws.Cells(i, 9).Value
                ws.Cells(Summary_Table_Row, 17) = Greatest_Increase
            
            'Check the ticker with the greatest decrease
            ElseIf ws.Cells(i, 11).Value = Greatest_Decrease Then
                
                'Print ticker and value in new summary table
                ws.Cells(Summary_Table_Row + 1, 16) = ws.Cells(i, 9).Value
                ws.Cells(Summary_Table_Row + 1, 17) = Greatest_Decrease
                
            'Check the ticker with the max volume
            ElseIf ws.Cells(i, 12).Value = Max_Volume Then
                
                'Print ticker and value in new summary table
                ws.Cells(Summary_Table_Row + 2, 16) = ws.Cells(i, 9).Value
                ws.Cells(Summary_Table_Row + 2, 17) = Max_Volume
            
            End If
            
        Next i
    
        'Loop through the value in the summary table
        For i = 2 To lastRow2
        
            'Check if the yearly change is greater then or equal to 0
            If ws.Cells(i, 10).Value >= 0 Then
                
                'Highlight the cell in green
                ws.Cells(i, 10).Interior.ColorIndex = 4
    
            'Check if the yearly change is smaller then 0
            ElseIf ws.Cells(i, 10).Value < 0 Then
            
                'Highlight the cell in red
                ws.Cells(i, 10).Interior.ColorIndex = 3
    
            End If
    
        Next i
        
        'Format Percentage Change column in percentage
        ws.Range("K2:K" & lastRow2).NumberFormat = "0.00%"
        
        'Format greatest increase and decrease values in percentage
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        'Adjust the width of the columns
        ws.Columns("A:Q").AutoFit
        
    Next ws
    
End Sub
