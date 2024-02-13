Attribute VB_Name = "Module3"
Sub summarize_trade_simple():

    Dim stock_vol As LongLong   'Running total on the volume of a stock. We get an overflow with "long" and rounding errors with floating point, so "longlong"
    Dim ticker As String        'This will store the current stock ticker symbol
    Dim first_price, last_price As Single       'These two variables will keep track of the first price of the year and closing price of the year
    Dim i As Long                '"i" will be our counter for looping through rows
    Dim result_row As Integer       'We need to keep track of which row we will be writing out our summary data to as we move through tickers

    
    'Lets start by setting column K as having a percent format so we don't need to use a type cast conversion to a string later
    'And also setting the number format for column "K" to avoid scientific notation on large volumes and show comas to make it legible
    Columns(11).NumberFormat = "0.00%"
    Columns(12).NumberFormat = "#,###"
    
    first_price = Cells(2, 3).Value     'Get the opening value of the first stock on the first day as a starting point (we'll update it as we iterate through rows)
    stock_vol = 0       'Explicitly setting our opening variable values to zero...
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
              
                
                'Now we need to reset our various values as we move to the next stock ticker
                stock_vol = 0                                           'Resets our stock volume to zero as we start next stock
                first_price = Cells(i + 1, 3).Value         'Sets our opening day price for the next stock
                result_row = result_row + 1
            End If
        Next i
        

End Sub

