## Overview Of The Project

### Purpose

The goal is to provide our client an easy and efficient way to analyze the stocks in the worksheet for the years 2017 and 2018 and determine whether or not the stocks are worth investing. This process was originally completed in a similar format, however, the goal for this project is to increase the efficiency of the original code.

### The Data

The data that is presented includes two charts with stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. 

### Results

After completing the analysis of the stocks you can clearly see that 2018 was not the best year in the market for most of these stocks. The average return for the stocks on the worksheet that we analyzed was -8.5%. In comparison NASDAQ's return was -4.36% in 2018. However the stocks provided on this worksheet had a phenomenal 2017 with an average return of 67.3%, compared to NASDAQ's 27% return in 2017.

### Analysis

Before refactoring the code, I began by copying the code that was needed to create the input box, chart headers, ticker array, and to activate the appropriate worksheet. The steps were then listed out in order to set the structure for the refactoring. Below is the code after refactoring with the steps laid out as comments.
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row's ticker doesn't match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
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
# Summary

### Pros and Cons of Refactoring Code

Refactoring helps make our code cleaner and more organized. A few advantages of a cleaner code include design and software improvement, debugging, and faster programming. It may also benefit other users who view our projects because it becomes easier to read, as it is more concise and straightforward. However, we do not always have the luxury to refactor our code due to the one and only disadvantage I can think of, Time consumption. Refactoring code takes time and may involve a lot debugging.

### The Advantages of Refactoring Stock Analysis

The biggest benefit that occurred as a result of refactoring the macro code for all stocks analysis is the decrease in macro run time. The original analysis took approximately 0.33 second to run, whereas our new analysis only took about a fiffth of the time (approximately 0.06 seconds) to run. Attached below are the screenshots that indicate the run time for our old and new analysis.
