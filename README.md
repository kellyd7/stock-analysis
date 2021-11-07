# stock-analysis
## Overview

### Background

Steve loves the workbook we prepared for him in `green_stocks`. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although our code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

### Purpose

In this challenge, I’ll edit/refactor, the Module 2 code solution, in `green_stocks`, to loop through all the data one time in order to collect the same information that I did in this module. Then, I’ll determine whether refactoring the code successfully made the VBA script run faster.

## Results

I will explain how the Module 2 code was edited/refactored in series of steps:

1. First, I created a `tickerIndex` variable and set it equal to zero. Next, I created three output arrays, `tickerVolumes`, `tickerStartingPrices` and `tickerEndingPrices`. Each output array was also assigned a data type, as seen below:

  ![This is an image](https://github.com/kellyd7/stock-analysis/blob/main/Resources/1.png)

2. Second, I created two `for` loops. One to initialize the `tickerVolumes` to zero and one that will loop over all the rows in spreadsheet. However, before creating the second `for` loop, I activated the `yearValue` worksheet, which is the worksheet that corresponds to the year selected  in the input promt at the start of the analysis. See code below:

  ![This is an image](https://github.com/kellyd7/stock-analysis/blob/main/Resources/2.png)

3. Now that I have a `for` loop that will loop over all the rows in the spreasheet, I used an `if-then` statements to increase the `tickerVolumes` for the current ticker and add the ticker volume for the current stock ticker. I aslo used `if-then` statements to check if the current row is the first row with the selected tickerIndex and if it is, then assign the current starting price to the `tickerStartingPrices`. Another `if-then` statement I created was one to check if the current row is the last row with the selected tickerIndex and if it is, then assign the current closing price to `tickerEndingPrices`. One thing to note in the code below, is that part 3d is commented out of the script. I found that this portion of the coded was not necessary to executed the task at hand. it only added to the memory load and slowed performance time.

  ![This is an image](https://github.com/kellyd7/stock-analysis/blob/main/Resources/3.png)

4. Finally, the last edit made, was to create a loop through the arrays to output the Ticker, Total Daily Volume and Return. To do this I activated the `All Stocks Analysis` worksheet, which is where these outputs will be reported. Next, I assigned `Cells().Value` to each output array. And, in order to get the code to loop, this is all placed outside of the `for` loop with a `j` counter and inside the `for` loop with an `i` counter. See code below:

  ![This is an image](https://github.com/kellyd7/stock-analysis/blob/main/Resources/4.png)

When the VBA script is ran for each year, 2017 and 2018. We get the following outputs in the `All stocks Analysis` worksheet below. Notice how the run times are 0.65secs whereas in the original `green_stocks` VBA script we see higher run times.

2017

  ![This is an image](https://github.com/kellyd7/stock-analysis/blob/main/Resources/2017.png)
  
2018

  ![This is an image](https://github.com/kellyd7/stock-analysis/blob/main/Resources/2018.png)
 
Green Stocks 2017

  ![This is an image](https://github.com/kellyd7/stock-analysis/blob/main/Resources/green_stocks_2017.png)

Green Stocks 2018
 
 ![This is an image](https://github.com/kellyd7/stock-analysis/blob/main/Resources/green_stocks_2018.png)
  


## Summary

As a result of performing this analysis, I was able to discover some advantages and disadvatages of code refactoring.

### Advantages:
    * Code refactoring allow for a more smooth processing experience.
    * Makes the code more extensible, meaning more functions can be easily added afterwards.
    * The Code becomes easier to read/understand and easier to maintain.

### Disadvanteages:
    * Code refactoring is very time consuming
    * There is also a lot of room for error. Whether, it's from adding new code to the script or deleting needed coded mistakely.
    
Keeping this advantages and disadvantages in mind and thinking about the `green_stocks`script and the `VBA_challenge.xlsm` script, it is evident that a key advantage of refactoring the code was faster performance time and easy to read code. However, the only disadvantage I can speak to is that the refactored code was very time consuming for me. But, I believe it was necessary and worth the time, given the results.



