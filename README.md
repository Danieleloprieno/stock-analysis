# Refactored VBA Code for Stock Analysis

##Overview
An inital analysis of stock prices was conducted on historical stock data in order to inform future purchase decisions.  While inital results were promising, refactored code was created in order to optimize the code for larger data calls.  

##Results
In order to optomize code performance, references to ticker array by index position were used to nest the loops of the code.  See excerpt below for details:

```
1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12)  As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
   Worksheets(yearValue).Activate
   For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then

        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then

        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
        '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then

                   tickerIndex = tickerIndex + 1
        
                End If
            
        Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```

In 2017, all stocks in the specified array experienced positive returns with the excepetion of TERP, which had an overall return of -7.2%.  2018 was a far worse year for all stocks in question; all experienced negative returns excpet for ENPH and RUN.  See screenshots below for anaysis outputs:
![Table of Results for 2017](/Resources/Stocks 2017.png)
![Table of Results for 2018](/Resources/Stocks 2018.png)

Time tests indicate that the refactored code did perform, on average, about .25 or .30 secs faster than the inital code.  See screenshots below for refartoced code performance time for 2017 and 2018:

![Time Test for 2017 Results](/Resources/Refactored 2017.png)
![Time Test for 2018 Results](/Resources/Refactored 2018.png)

##Summary




