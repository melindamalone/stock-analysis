# All Stocks Analysis

## Overview of Project

### Purpose

###### The purpose of Module Two and the All Stocks Analysis Refactored Challenge is to learn how to write and execute code in Visual Basic for Applications (VBA), how to refactor code in VBA, and the benefits of refactoring code in VBA.  For the All Stocks Analysis Refactored Challenge, the original All Stocks Analysis code was refactored primarily by introducing tickerIndex as a variable and using it to loop through multiple arrays. 

## Results

###### By introducing tickerIndex as a variable in the refactored code and using it to access the stock ticker index for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays, the elapsed run time for executing the code improved by approximately 79% for both 2017 and 2018 stock data from the original script to the refactored script.


### AllStocksAnalysisRefactored
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
        
        '3c) Check if the current row is the last row with the selected ticker. 
	'If the next row's ticker doesn't match, increase the tickerIndex
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

### All Stocks Analysis 2017 Elapsed Run Time (original script)
![](Resources/all_stocks_analysis_2017.PNG)

### All Stocks Analysis 2017 Elapsed Run Time (refactored script)
![](Resources/VBA_Challenge_2017.PNG)

### All Stocks Analysis 2018 Elapsed Run Time (original script)
![](Resources/all_stocks_analysis_2018.PNG)

### All Stocks Analysis 2018 Elapsed Run Time (refactored script)
![](Resources/VBA_Challenge_2018.PNG)

## Summary

###### In summary, advantages and disadvantages of refactoring code exist and in my opinion, exist on a case-by-case basis.  For example, an advantage of refactoring the All Stocks Analysis code improved the run time by approximately 79% allowing the user the benefit of time-efficiency. However disadvantages of refactoring code are possible if sufficient comments are not used in the original script. This would make it difficult for a new person to refactor the original code because they might not be familiar with the original code's intent. Also it is possible for a new person refactoring the code to change the desired output of the script or break it completely. These pros and cons apply to the All Stocks Analysis Challenge because the elapsed run time definitely improved from .97 seconds run time to .20 seconds run time. As a new user of VBA, it was a difficult challenge for me to originally refactor the All Stocks Analysis code. There were countless days where the refactored code would not work and therefore making it impossible to run the analysis on the 2017 and 2018 stock data.  Finally, and with great help and support from my tutor, Mark Fullton, I persevered and was able to identify the gaps in my refactored script. 