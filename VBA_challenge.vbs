Sub VBA_challenge()

'Loop through all the stocks for one year and outputs the following information:
'       The ticker symbol
'       Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'       The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'       The total stock volume of the stock.

For Each ws In Worksheets
    Dim worksheet_name As String

'set an initial variable for holding the current row
    Dim i As Long
''set an initial variable for holding the starting row of ticker
    Dim j As Long
'set an initial variable for holding the ticker name
    Dim ticker_name As Long
'set last row in ticker, column A
    Dim ticker_last_row As Long
'set last row in ticker of summary table, column I
    Dim tickersummary_last_row As Long
'set an initial variable for holding the percentage change
    Dim percent_change As Double
'set an initial variable for holding the greatest increase
    Dim greatest_increase As Double
'set an initial variable for holding the greatest decrease
    Dim greatest_decrease As Double
'set an initial variable for holding the greatest volume
    Dim greatest_volume As Double
'set an initial variable for holding the running total    
    Dim total as double

' set initial values for the running total
    total = 0

'obtain worksheet_name
    worksheet_name = ws.Name
 
 ' Label summary table column names
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"

'Set Ticker Counter to first row
    ticker_name = 2
'Set starting row of ticker to 2
    j = 2
    
' Loop through all ticker stocks by rows
'   For i = 0 To 753001 or,
    ticker_last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To ticker_last_row
       
    'Check if next ticker name is same, if it is not...
        If ws.Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        total = total + ws.cells(i, 7).value

        'Print the ticker name, column I (column 9)
            ws.Cells(ticker_name, 9).Value = ws.Cells(i, 1).Value
        'Calculate and print yearly change, column J (column 10)
            ws.Cells(ticker_name, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                    
            'Conditional formatting, column J (column 10)
                If ws.Cells(ticker_name, 10).Value < 0 Then
                'Set cell background color to red (for negative change)
                    ws.Cells(ticker_name, 10).Interior.ColorIndex = 3
                Else
                    'Set cell background color to green (for positive change)
                        ws.Cells(ticker_name, 10).Interior.ColorIndex = 4
                End If
        
        'Calculate and print percent change, column K (column 11)
            If ws.Cells(j, 3).Value <> 0 Then
                percent_change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                
                'Set value as percentage format
                    ws.Cells(ticker_name, 11).Value = Format(percent_change, "Percent")
            Else
                ws.Cells(ticker_name, 11).Value = Format(0, "Percent")
            End If
        
        'Calculate and print total volume, column L (column 12)
            ws.Cells(ticker_name, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
        'Increase ticker_name by 1
            ticker_name = ticker_name + 1
        'Set new start row of the ticker block
            j = i + 1
        'set total in range("L2",j).value
        ws.range("L" &  2 + j).value = total
    ' If the ticker name is the same...
        Else 
            total = total + ws.cells(i, 7).value

        End If
    Next i
    
 'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
        
     'Label the for second summary table
        greatest_increase = ws.Range("K2").Value
        greatest_decrease = ws.Range("K2").Value
        greatest_volume = ws.Range("L2").Value
    
    'Loop second summary
        tickersummary_last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
        For i = 2 To tickersummary_last_row

    'For greatest increase, check if next value is larger. If yes, override and print
        If ws.Cells(i, 11).Value > greatest_increase Then
            greatest_increase = ws.Cells(i, 11).Value
            ws.Range("O2").Value = ws.Cells(i, 9).Value
        
        Else
         greatest_increase = greatest_increase

        End If
    
    'For greatest decrease, check if next value is larger. If yes, override and print
        If ws.Cells(i, 11).Value < greatest_decrease Then
            greatest_decrease = ws.Cells(i, 11).Value
            ws.Range("O3").Value = ws.Cells(i, 9).Value
        
        Else
            greatest_decrease = greatest_decrease
        
        End If
    
    'For greatest volume, check if next value is larger. If yes, override and print
        If ws.Cells(i, 12).Value > greatest_volume Then
            greatest_volume = ws.Cells(i, 12).Value
            ws.Range("O4").Value = ws.Cells(i, 9).Value
        
        Else
            greatest_volume = greatest_volume
        
        End If

     'Print second summary results
            ws.Range("P2").Value = Format(greatest_increase, "Percent")
            ws.Range("P3").Value = Format(greatest_decrease, "Percent")
            ws.Range("P4").Value = Format(greatest_volume, "Scientific")
    Next i

    'Adjust column widths automatically
        Worksheets(worksheet_name).Columns("A:Z").AutoFit
    
    Next ws

End Sub
