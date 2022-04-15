# Overview of Project
## Purpose
This project aimed to refactor a Microsoft Excel VBA code to improve the implementation and efficiency of the original VBA code for collecting stock information in both 2017 and 2018.  The purpose of this code is to determine which stocks may be worth investing in and can analyze an entire dataset of stocks. Even though the current focus is on stocks from Green Companies, this newly refactored code will also be able to work for larger datasets in a timely manner.  For this workbook specifically, we created two charts that compare stock information from 12 different companies using the ticker value, the total daily volume, and the yearly return.

### The Data
stock analysis of green energy company
## Results
Using my original code, I was able to achieve the desired results, yet with a run time of 0.9453125 for stocks from 2017 and a run time of 0.984375 for stocks from 2018. To improve this time, some changes needed to be made. Before refactoring the code, I first copied the original code from my “All Stocks Analysis” macro. I made no changes to the code that set the runtime, created an input box, labeled the chart headers, added a ticker array, activated the worksheet, and to find the number of rows to loop over. Key changes were then made to add a tickerIndex variable and three more arrays for ticker volumes, ticker starting prices, and ticker ending prices were added before looping through all rows and running our conditionals. You can see these changes in the image below.

![Arrays_Added](https://user-images.githubusercontent.com/102122063/163629199-61d8205a-9b83-43f7-98af-59f1b37c44cd.JPG)

We were able to use the tickerIndex variable to access the stock ticker index for all our arrays. Within the script loops, we were able to successfully read and store data from each row for the Ticker, Ticker Volumes, Ticker Starting Prices, and Ticker Ending Prices. Once complete, the code for formatting the cells was included within the same macro to automate the formatting when switching between years. You can view the refactored code for the loops and formatting below. 

![Loops_Formatting](https://user-images.githubusercontent.com/102122063/163630477-941d5c9b-f111-46b3-94a8-0056022a1f4d.JPG)


## Summary
### Pros can Cons of Refactoring Code

### The Advantages of Refactoring Stock Analysis
