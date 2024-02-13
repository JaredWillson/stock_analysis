# stock_analysis
Script to summarize stock performance over the year

Script performs several actions...
  - Loops through a list of stock trading values by day (ticker, date, open, close, volume)
  - Stores the opening value of the stick for the first day of the year
  - Keeps a running total of the volume for the ticker for the year
  - On the last day a stock traded, determines closing value
  - Outputs summary information for each stock for the year including total volume, % change, $ change
  - Determines and outputs biggest winner and loser (by percent change) and largest volume stock
  - Applies conditional formatting to highlight positive and negative changes
  - Repeats all of the above for however many worksheets are in the workbook
