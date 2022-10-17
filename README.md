# Stock-Analysis

## Overview of Project

### Purpose
There were two purposes for this project. The first was to help “Steve” determine which stocks would be a good investment for his parent’s based on the data set we were given. The second purpose was to refactor the original code to see if the changes made any improvement in how fast the code ran.

### The Data
The data that was used was for the years 2017 and 2018. There was a sample set of 12 different stocks which included their starting price and ending price and total volume for day of trade. From this data we created code to determine the total daily volume and rate for return for each year. 


## Results

### Analysis

#### Stock Investment Results
Originally “Steve’s” parents wanted to invest all their money into DQ a green energy company. Being in finance Steve wanted to run an analysis of the DQ stock to see if it would be a good investment. When the data showed it ran at loss, Steve wanted “us” to preform an analysis of which stocks would be a good investment based on the data we had. Although 11 of the 12 stocks showed positive returns in 2017, over all between both years of data only the tickers of ENPH and RUN both had positive returns.

#### Refactoring Results
The next task was to refactor the original code we created by streamlining the code to make it more efficient and use less memory. We were not tasked with showing if the code took up less memory, however, we were tasked to see if there was a change in how fast the program ran between the original and the refactored program. 
For the refactored code we were given specific guidelines to improve the efficiency of the code. The code that was used is as follows:

    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ‘2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
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
        
#### Refactoring Elapsed Time
After refactoring the elapsed time difference for 2017 was 0.640625 seconds. The elapsed time difference for 2018 was 0.0117187 seconds. In this scenario there was not much of a difference, however, in a larger program the time refactoring could make it run faster and more efficiently


<img src="https://github.com/dbrashears63/Stock-Analysis/blob/main/VBA%20Module%202%20Deliverables/Message%20Box%20Run%20Time%20For%202017.png" width=20% height=20%>
<img src="https://github.com/dbrashears63/Stock-Analysis/blob/main/VBA%20Module%202%20Deliverables/Message%20Box%20Run%20Time%20For%202017%20Refactored.png" width=20% height=20%>


## Summary

### Summary of All Stocks
Overall if this was going to be a real analysis there was not enough data to preform a proper analysis. However, for the project we found the total daily volumes for each ticker and what the yearly return was for each ticker. We used the results to make a recommendation for which stocks would be best for “Steve’s” parents to invest in.

### Summary of Refactoring Code
The goal of refactoring is to make the code cleaner and make it more organized so it will run more 
efficiently and take up less memory.


## Pros and Cons of Refactoring

### Pros
The main benefit of refactoring is having the code more organized and cleaner. In addition, by having clean well documented code it may help in the speed in which a program runs and help with debugging among other benefits. By having clean efficient code, it will make it easier to read and make any updates or changes

### Cons
Some of the disadvantages to refactoring may include not having all the original documentation for the code one is working with. By not having that information there is a risk of introducing bugs into the program. Refactoring should not be done if there is a short deadline or there is not enough manhours to have it don’t correctly.


    
