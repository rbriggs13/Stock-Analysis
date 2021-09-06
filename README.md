# VBA of Wall Street

## Overview of Project

### Purpose
The purpose of this analysis was to determine if having the code loop through all the data once as opposed to multiple times for each different stock would make the code run any faster. This was accomplished by refactoring the code written to loop through the data for each stock into code that did it all at once.

## Results

### Refactoring the Code
This is the original code for the For loop that required multiple loops through the data:

        'Loop through the tickers
        For i = 0 To 11
            ticker = tickers(i)
            Worksheets(yearValue).Activate
            totalVolume = 0
        
        'Loop through the data
        For j = rowStart To rowEnd
        
          'increase totalVolume
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
        
            'set starting price
             If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
        
            
                startingPrice = Cells(j, 6).Value
        
            End If
        
            'set ending price
             If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
        
                endingPrice = Cells(j, 6).Value
        
            End If
        
        Next j
        
        'output the data for the current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = ticker
        Cells(i + 4, 2).Value = totalVolume
        Cells(i + 4, 3).Value = endingPrice / startingPrice - 1
        
    Next i
    
As you can see this code stops looking through the data after the data involved with the current ticker in the array is over, then the entire loop starts over with the next ticker. Now looking at the refactored For loop:

    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
            End If
            
            
        
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
            
            
            

                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            
            
            End If
    
    Next i
    
This loop utilizes multiple arrays and If Then statements in order to only have to move down the data once. It accomplishes this by using the IF THEN statements to check once the data for a particular stock is finished and then move on to the next one.

### Analysis of the Execution Times
The time it took for the original code to execute this function was 0.703 seconds for the 2017 data and and 0.695 seconds for the 2018 data. While the time it took for the refactored code to execute the function was 0.172 seconds for 2017 and 0.133 seconds for 2018. These run times clearly indicated that the refactored code is much more efficient than the original.

## Summary 

### Pros of using Refactored Code
- Allows for you to use code that you've already written so you don't have to start from scratch, this is applied to this code because the basic structure of most of the different functions are unchanged.
- Allows the code to become better organized and more efficient. This is applied to this code by making the code both easier to follow and quicker to execute.

### Cons of using Refactored Code
- There is a risk of introducing bugs into the code. This can be applied to this code because during the process of refactoring the similarity between the old and new variables caused some errors. As well as forgetting to call the specific variable in an array because it used to just be a single variable.

