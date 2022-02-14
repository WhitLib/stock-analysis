# stock-analysis
Stock Analysis Project Using VBA

### Table of Contents
- [1 Overview](#1-overview)
- [2 Results](#2-results)
- [3 Summary](#3-summary)


## 1 Overview

This VBA stock analysis project was meant to assist Steve in analyzing various green energy stocks for his parents. Steve's parents are passionate about alternative energy solutions but have done minimal research in other investment opportunities aside from DAQO - the only stock currently in their portfolio. This project aimed to help Steve's parents diversify their portfolios and analyze other alternative energy companies as investment options.

In this segment of the project, Steve wanted to do some additional research and expand the dataset to include the entire stock market over the last few years. To do this successfully, looping through the data one time to collect information on 2017 and 2018 stock's ticker names, total daily volume, and return percentage.

## 2 Results 

### 2.1 Refactor VBA Code

The first task of this project was to refactor the VBA code from the completed module to accomplish each of the following steps:
- [x] Create a tickerIndex and set it equal to zero before looping over the rows

````
Dim tickerIndex As Integer
    tickerIndex = 0
````

- [x] Create arrays for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices 

````
    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
 ````
 
 ````
 '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
 ````
 
- [x] Ensure the tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays

````
        ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
            tickerStartingPrices(i) = 0
            tickerEndingPrices(i) = 0
        Next i
````

- [x] Make sure the script loops through stock data, reading and storing tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices 
Based on the results, overall, 2017 stocks outperformed 2018 stocks. In 2017, eleven out of twelve stocks analyzed had a positive return, whereas, in 2018, all but two stocks (ENPH and RUN) had positive returns. 

````
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then

            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
               
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
           
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then

            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
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
   ````
   
In 2017 in particular, some stocks - including DQ - had a return greater than 100%:
- DQ with 199.4%
- ENPH with 129.5%
- FSLR with 101.3%

Alternatively, in 2018, ten out of twelve stocks had negative returns (including DQ) and only two had positive: 
- ENPH  with 81.9%
- RUN with 84.0%


<img width="232" alt="Screen Shot 2022-02-14 at 8 50 37 AM" src="https://user-images.githubusercontent.com/95978097/153908742-b7c8d5a2-1478-4e8b-8097-4c5b44fe69e5.png">


<img width="232" alt="Screen Shot 2022-02-14 at 8 52 15 AM" src="https://user-images.githubusercontent.com/95978097/153909026-039132a2-dbe5-40bc-96ca-a87aace53175.png">
<img width="232" alt="Screen Shot 2022-02-14 at 9 05 26 AM" src="https://user-images.githubusercontent.com/95978097/153911507-1188dfe4-668d-42f2-9f88-6ae84294dd4e.png">
<img width="232" alt="Screen Shot 2022-02-14 at 9 05 42 AM" src="https://user-images.githubusercontent.com/95978097/153911544-f18e1647-cdee-4a43-8423-74845d221be9.png">
