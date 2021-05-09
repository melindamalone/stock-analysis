# stocks-analysis

```

'1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero
    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i
                     
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1) = tickers(tickerIndex) Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
            
        '3b) Check if the current row is the first row with the selected tickerIndex
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) Check if the current row is the last row with the selected ticker. If the next row's ticker doesn't match, increase the tickerIndex
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
    
        '3d Increase the tickerIndex
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
        End If
        
    Next i
           
    '4) Loop through arrays to output the Ticker, Total Daily Volume, and Return
    Worksheets("All Stocks Analysis").Activate
    For i = 0 To 11
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i

```

![](Resources/all_stocks_analysis_2017.png)

![](Resources/all_stocks_analysis_refactored_2017.png)

![](Resources/all_stocks_analysis_2018.png)

![](Resources/all_stocks_analysis_refactored_2018.png)