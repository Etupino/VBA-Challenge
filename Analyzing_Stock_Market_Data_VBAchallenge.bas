Attribute VB_Name = "Analyzing_Stock_Market"
Sub Analyzing_Stock_Market()


'loop to cycle through the worksheets in the workbook
    'Set a variable to cycle through the worksheets
    Dim ws As Worksheet

    'Start loop
    For Each ws In Worksheets

        'Create column labels for the Stock Index table
        ws.Range("I1").Value = "Ticker Symbol"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
		
		'Create column/rows for the Stock_movers table        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        
        Dim ticker_symbol As String
        Dim total_vol As Double
        total_vol = 0


        Dim rowcount As Single
        rowcount = 2

        'Set variable to hold year opening price
        Dim opening_price As Double
        opening_price = 0

        'Set variable to hold year closing price
        Dim Closing_price As Double
        Closing_price = 0
        
        'Set variable to hold the yearly price change
        Dim yearly_change As Double
        yearly_change = 0

        'Set variable to hold the yearly percent_change
        Dim percent_change As Double
        percent_change = 0

        'Set variable for total rows to loop through
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop to search through ticker symbols
        For i = 2 To lastrow
            
            'Conditional to grab opening price
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                opening_price = ws.Cells(i, 3).Value

            End If

            'Sum up the volume for each row to determine the total stock volume for the year
            total_vol = total_vol + ws.Cells(i, 7)

            'determine if the ticker symbol is sync in the Stock Index Table
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                'Copy ticker symbol to Index table
                ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value

                'Copy total stock volume to the Stock Index table
                ws.Cells(rowcount, 12).Value = total_vol

                'Grab year closing price
                Closing_price = ws.Cells(i, 6).Value

                'price change for the year and move it to the Stock Index table.
                yearly_change = (closing_price - opening_price)
                ws.Cells(rowcount, 10).Value = yearly_change
                ws.Cells(rowcount, 10).NumberFormat = "0.00"

                'format to highlight positive or negative change.
                If yearly_change >= 0 Then
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 43
                Else
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 30
                End If

                'Calculate the percent change for the year and move it to the summary table format as a percentage
                'Conditional for calculating percent change
                If opening_price = 0 And closing_price = 0 Then
                    'Starting at zero and ending at zero will be a zero increase.  Cannot use a formula because
                    'it would be dividing by zero.
                    percent_change = 0
                    ws.Cells(rowcount, 11).Value = percent_change
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                ElseIf opening_price = 0 Then
                    'If a stock starts at zero and increases, it grows by infinite percent.
                    'therefore, assign actual price increase by dollar amount to "New Stock" as percent change                  
                    Dim percent_change_NA As String
                    percent_change_NA = "New Stock"
                    ws.Cells(rowcount, 11).Value = percent_change
                Else
                    percent_change = yearly_change / opening_price
                    ws.Cells(rowcount, 11).Value = percent_change
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                End If

                'Add 1 to rowcount to move it to the next empty row in the Stock Index table
                rowcount = rowcount + 1

                'Reset total stock volume, opening price, closing price, yearly change, percent change
                total_vol = 0
                opening_price = 0
                closing_price = 0
                yearly_change = 0
                percent_change = 0
                
            End If
        Next i        

        'Assign lastrow to count the number of rows in the Stock Index table
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

        'Set variables to hold Top gainers, losers, and stock with the Highest volume
        Dim highest_stock As String
        Dim highest_value As Double

        'Set top gainers equal to the first stock
        Highest_value = ws.Cells(2, 11).Value

        Dim lowest_stock As String
        Dim lowest_value As Double

        'Set top losers equal to the 1st stock
        lowest_value = ws.Cells(2, 11).Value

        Dim highest_vol_stock As String
        Dim highest_vol_value As Double

        'Set highest volume equal to the 1st stock
        highest_vol_value = ws.Cells(2, 12).Value

        'Loop to search through Index table
        For j = 2 To lastrow

            'Conditional to determine Top Gainers
            If ws.Cells(j, 11).Value > highest_value Then
                highest_value = ws.Cells(j, 11).Value
                highest_stock = ws.Cells(j, 9).Value
            End If

            'Conditional to determine Top losers
            If ws.Cells(j, 11).Value < lowest_value Then
                lowest_value = ws.Cells(j, 11).Value
                lowest_stock = ws.Cells(j, 9).Value
            End If

            'determine stock with the highest volume traded
            If ws.Cells(j, 12).Value > highest_vol_value Then
               highest_vol_value = ws.Cells(j, 12).Value
                highest_vol_stock = ws.Cells(j, 9).Value
            End If

        Next j

        'Copy top gainers, losers, and stock with the highest volume items to the stock_movers table
        ws.Range("P2").Value = highest_stock
        ws.Range("Q2").Value = highest_value
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P3").Value = lowest_stock
        ws.Range("Q3").Value = lowest_value
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P4").Value = highest_vol_stock
        ws.Range("Q4").Value = highest_vol_value

        'Autofit table columns
        ws.Columns("I:L").EntireColumn.AutoFit
        ws.Columns("O:Q").EntireColumn.AutoFit

    Next ws


End Sub

