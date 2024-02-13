Attribute VB_Name = "Module2"
Sub summarize_trade_intermediate():

    Dim stock_vol As LongLong   'Running total on the volume of a stock. We get an overflow with "long" and rounding errors with floating point, so "longlong"
    Dim ticker As String        'This will store the current stock ticker symbol
    Dim first_price, last_price As Single       'These two variables will keep track of the first price of the year and closing price of the year
    Dim i As Long                '"i" will be our counter for looping through rows
    Dim result_row As Integer       'We need to keep track of which row we will be writing out our summary data to as we move through tickers
    Dim big_increase, big_decrease As Single    'As we total our stocks, we will also keep track of some maximum values
    Dim big_vol As LongLong                  'This includes our largest volume, not just our largest percent changes
    Dim ticker_big_increase, ticker_big_decrease, ticker_big_vol As String          'As we collect our max values, we will need the associated ticker symbols
    
    'Lets start by setting column K as having a percent format so we don't need to use a type cast conversion to a string later
    'And also setting the number format for column "K" to avoid scientific notation on large volumes and show comas to make it legible
    Columns(11).NumberFormat = "0.00%"
    Columns(12).NumberFormat = "#,###"
    
    first_price = Cells(2, 3).Value     'Get the opening value of the first stock on the first day as a starting point (we'll update it as we iterate through rows)
    stock_vol = 0       'Explicitly setting our opening variable values to zero...
    big_increase = 0
    big_decrease = 0
    big_vol = 0
    result_row = 2      'We are going to write our results into the result section of each worksheet starting at row "2". This value will increment with new tickers
    
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row     'We will run through all rows, using our "trick" to see how many rows we need dynamically
            ticker = Cells(i, 1).Value              'No matter what row we are in, our ticker is always the first column of our current row
            If Cells(i + 1, 1).Value = Cells(i, 1).Value Then   'if we are not about to change stock ticker symbol, we just need to add volume to the running total
                stock_vol = stock_vol + Cells(i, 7).Value
            Else                                                                    ' If we are on the last record for a ticker symbol, we will need to run our calculations and write our results
                stock_vol = stock_vol + Cells(i, 7).Value                           ' We still need to add this last day's volume in for the stock
                Cells(result_row, 9).Value = ticker     'Write out our stock ticker symbol
                last_price = Cells(i, 6).Value              'This is the last trade of the year for the ticker, so can grab the closing price for percent change
                Cells(result_row, 10).Value = Round(last_price - first_price, 2)   'Write out our change for the year from first open to final close
                ' Let's apply our color coding to the yearly change to indicate whether it is positive or negative...
                If Cells(result_row, 10).Value >= 0 Then             'Note that there is no situation where yearly change and percent change is negative, so the single if works
                    Cells(result_row, 10).Interior.Color = vbGreen   'Green for positive or zero values in yearly change
                    Cells(result_row, 11).Interior.Color = vbGreen    'Green for positive or zero values in percent change
                Else
                    Cells(result_row, 10).Interior.Color = vbRed
                    Cells(result_row, 11).Interior.Color = vbRed         'Red for negative values
                End If
                Cells(result_row, 11).Value = Round((last_price - first_price) / first_price, 4)         'Write out our percent change for the stock
                Cells(result_row, 12).Value = stock_vol                                                                  'Write out our total volume for the stock for the year
                
                
                'Now we have all our summary info for a given ticker... Let's see whether this one is any of our "max value" stocks...
                 If stock_vol > big_vol Then                             'Check to see if this is our largest volume stock to date
                    big_vol = stock_vol                                     'If it is, then update our running "largest stock volume" stock...
                    ticker_big_vol = Cells(i, 1).Value               '...including the ticker symbol. We don't need an else since we would "do nothing" with it.
                End If
                If ((last_price - first_price) / first_price) > big_increase Then       'Check to see if this is our highest increasing stock by percent so far
                    big_increase = ((last_price - first_price) / first_price)               'If it is, then replace our current record holder in the clubhouse
                    ticker_big_increase = Cells(i, 1).Value                                  'Including which stock ticker it was, not just the new value
                End If
                If ((last_price - first_price) / first_price) < big_decrease Then       'Do the same for the worst performing stock...
                    big_decrease = ((last_price - first_price) / first_price)
                    ticker_big_decrease = Cells(i, 1).Value
                End If
                
                'Now we need to reset our variousvalues as we move to the next stock ticker
                stock_vol = 0                                           'Resets our stock volume to zero as we start next stock
                first_price = Cells(i + 1, 3).Value         'Sets our opening day price for the next stock
                result_row = result_row + 1
            End If
        Next i
        
        'Now that we have looped throuh all our rows, written our summaries, and collected our max values along the way, let's write out our max values
        Cells(2, 15).Value = "Greatest % Increase"               'Writes a text string data label
        Cells(3, 15).Value = "Greatest % Decrease"               '...
        Cells(4, 15).Value = "Greatest Total Volume"            '...
        Cells(2, 16).Value = ticker_big_increase                 'Writes out the stock symbol that had the biggest increase
        Cells(2, 17).Value = big_increase                            'Writes out the actual % increase for this "winner"
        If big_increase > 0 Then
            Cells(2, 17).Interior.Color = vbGreen                   'Color codes the cell for biggest winner
        End If
        Cells(2, 17).NumberFormat = "0.00%"                      'Formats the cell as a percent so it will display properly without changing it to a string
        Cells(3, 16).Value = ticker_big_decrease                 'Writes out the stock symbol that had the biggest decrease
        Cells(3, 17).Value = big_decrease                            'Writes out the actual % decrease for this "loser"
        Cells(3, 17).NumberFormat = "0.00%"                      'Formats the cell as a percent
        If big_decrease < 0 Then
            Cells(3, 17).Interior.Color = vbRed                      'Colors the bigest "loser" as red
        End If
        Cells(4, 16).Value = ticker_big_vol                          'Writes the ticker for the highest volume stock
        Cells(4, 17).Value = big_vol                                     'Writes the volume for this high-volume stock
        Cells(4, 17).NumberFormat = "#,###"                      'So it won't show in scientific notation, sets the format for the cell'
 
End Sub
