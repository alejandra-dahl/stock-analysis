# VBA Challenge - All Stocks Analysis (Refactored)

## Overview of Project

### Purpose
Steve wanted a way to run analysis that would include the entire stock market over the last few years. The years provided were 2017 and 2018, separated into their own tabs. The development of this refactored VBA script needed to be able to handle a large amount of data for any year, including the two provided for the entire stock market.

## Results
The VBA code created allows Steve to see the total daily volume and percent return for each stock (ticker) all on a separate tab in the workbook. Each stock was assigned a ticker in an array. The code will run in a loop and assign each stock a ticker. 
>For Example

    tickers(0) = "AY"

    tickers(1) = "CSIQ"
 
The macro then needs to find the total volume of each ticker by finding the starting price and end current ending price. It will also format the percent return by highlighting it green for positive and red for negative.

Buttons were installed so that Steve or his parents can easily run the analysis on any tab year that they would like.

Running the macro will result in images similar to what is seen below.

**2017**

![VBA_Challenge_2017](https://user-images.githubusercontent.com/90485451/136884848-0533865c-c543-4037-8654-952b86b226c0.png)



**2018**

![VBA_Challenge_2018](https://user-images.githubusercontent.com/90485451/136884808-4cf3f2ff-10d9-405e-b2dc-3046661fc92c.png)



## Summary

1. What are the advantages or disadvantages of refactoring code?

2. How do these pros and cons apply to refactoring the original VBA script?
