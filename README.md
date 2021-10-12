# VBA Challenge - All Stocks Analysis (Refactored)

##  Purpose
Steve wanted a way to run analysis that would include the entire stock market over the last few years. The years provided were 2017 and 2018, separated into their own tabs. The development of this refactored VBA script needed to be able to handle a large amount of data for any year, including the two provided for the entire stock market.

## Results
The VBA code created allows Steve to see the total daily volume and percent return for each stock (ticker) all on a separate tab in the workbook. A ticker index and three output arrays were created. When the code runs, the ticker index will increase if the next row's ticker doesn't match.

>For Example

    For i = 0 To 11
        tickerIndex = tickers(i)

    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single, tickerEndingPrices As Single
 
A similar process happens for volume. The macro then needs to find the total volume of each ticker by finding the starting price and end current ending price. 

        Worksheets(yearValue).Activate
        tickerVolumes = 0

        For j = 2 To RowCount
    
        If Cells(j, 1).Value = tickerIndex Then

            tickerVolumes = tickerVolumes + Cells(j, 8).Value
            
        End If
        
          If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

            tickerEndingPrices = Cells(j, 6).Value


It will loop through the arrays to output the **Ticker**, **Total Daily Volume**, and **Return**. 
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickerIndex
        Cells(4 + i, 2).Value = tickerVolumes
        Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1


It will also format the percent return by highlighting it green for positive and red for negative.


Buttons were installed so that Steve or his parents can easily run the analysis on any tab year that they would like.

Running the macro will result in images similar to what is seen below.

**2017**
It's easy to see that TERP was the only stock that had a negative return that year. The macro was also able to run in 0.93 seconds.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/90485451/136884848-0533865c-c543-4037-8654-952b86b226c0.png)



**2018**
ENPH and RUN were the only stocks in 2018 with a positive return. The macro was also able to run in 0.87 seconds.

![VBA_Challenge_2018](https://user-images.githubusercontent.com/90485451/136884808-4cf3f2ff-10d9-405e-b2dc-3046661fc92c.png)



## Summary

1. What are the advantages or disadvantages of refactoring code?

2. How do these pros and cons apply to refactoring the original VBA script?
