# Stocks-Analysis
## Overview of Project
The purpose of this analysis was to calculate the total volume traded and annual returns for 12 different stocks for the the years 2017 and 2018 using a refactored script to minimize the runtime.  
## Results
### 2017
![2017 Performance](https://github.com/nleva25/Stocks-Analysis/blob/main/Resources/2017%20Results.png) ![2017 Runtime Results](https://github.com/nleva25/Stocks-Analysis/blob/main/Resources/VBA_Challenge_2017.png)
Overall, 2017 was a good year for all the companies with the exception of TERP which saw a negative annual return of 7.2%. The other 11 comapanies had postitive returns and of those, 9 had double digit returns. Only RUN and AY had returns lower than 10%.

The overall runtime was lower as well. The refactored script ran in 0.1171875 seconds for 2017 compared to the original srcipt at 0.7109375 seconds. 

### 2018
![2018 Performance](https://github.com/nleva25/Stocks-Analysis/blob/main/Resources/2018%20Results.png) ![2018 Runtime Results](https://github.com/nleva25/Stocks-Analysis/blob/main/Resources/VBA_Challenge_2018.png)
2018 performance was much worse than 2017. Out of the 11 stocks 9 saw negative returns, and 2 had all of their 2017 gains erased: JKS ans SPWR. Only 2 stocks continued to have positive annaul returns: ENPH and RUN at 81.9% and 84.0% respectively.

The runtime was lower as well. The refactored script ran in 0.1015625 seconds compared to 0.7265625 seconds for the original script. 

### Refactored Script 
The reason for the shorter runtimes seems to stem from the fact that the refactored script doesn't use a nested For Loop. Instead of cycling through the each row in the spreadsheet **before** moving to the next ticker (see below for orignal script), the new script uses a variable (tickerIndex) and conditional formatting to track which ticker the conditional formatting uses. This decreases the amount of times the script runs whereas a nested For Loop compounds the amount of times a script runs, which in turn increases runtime. 

For i = 0 To 11
    
        ticker = tickers(i)
        
        totalVolume = 0
        
        '5) Loop through rows in the data
        
        Sheets(yearValue).Activate
        
        For j = 2 To RowCount
        
            '5a) get total volume for current ticker
            
            If Cells(j, 1).Value = ticker Then
            
                totalVolume = totalVolume + Cells(j, 8).Value
                
            End If
            
            '5b) Get starting price for current ticker
            
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                startingPrice = Cells(j, 6).Value
                
            End If
            
            '5c) Get ending price for current ticker
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                endingPrice = Cells(j, 6).Value
                
            End If
            
        Next j
        
        
        '6) Output data for current ticker
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = ticker
        
        Cells(4 + i, 2).Value = totalVolume
        
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
        
      Next i
## Summary
Advantages to refactoring code include: 

    1) Increased efficiency by lowering runtime
    2) Smaller file sizes
    3) Simpler and better organized script

Disadvantages include: 

    1) Time consuming 
    2) Identifying which lines of code are essential over which ones are not 

The new refactored VBS script is much simpler in that it removes the nested For Loop which in turn deecreases the runtime. The script is much simpler and easier to read overall (see below, notice the removal of the nested For Loop): 

Worksheets(yearValue).Activate
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
    Next i  

The disadvantages to the refactored script are that it doesn't include the output script within the nested For Loop in the original, so it had to be re-written within a seperate For Loop afterwards. 



