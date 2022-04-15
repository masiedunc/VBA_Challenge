# Overview of Project
## Purpose
This project aimed to refactor a Microsoft Excel VBA code to improve the implementation and efficiency of the original VBA code for collecting stock information in both 2017 and 2018.  The purpose of this code is to determine which stocks may be worth investing in and can analyze an entire dataset of stocks. Even though the current focus is on stocks from Green Companies, this newly refactored code will also be able to work for larger datasets in a timely manner.  For this workbook specifically, we created two charts that compare stock information from 12 different companies using the ticker value, the total daily volume, and the yearly return.

### The Data
stock analysis of green energy company
## Results
Using my original code, I was able to achieve the desired results, yet with a run time of 0.9453125 for stocks from 2017 and a run time of 0.984375 for stocks from 2018. To improve this time, some changes needed to be made. Before refactoring the code, I first copied the original code from my “All Stocks Analysis” macro. I made no changes to the code that set the runtime, created an input box, labeled the chart headers, added a ticker array, activated the worksheet, and to find the number of rows to loop over. Key changes were then made to add a tickerIndex variable and three more arrays for ticker volumes, ticker starting prices, and ticker ending prices were added before looping through all rows and running our conditionals. You can see these changes in the image below.

![Arrays_Added](https://user-images.githubusercontent.com/102122063/163631247-a3edd404-b66f-4e84-9091-2eead4ece011.png)

We were able to use the tickerIndex variable to access the stock ticker index for all our arrays. Within the script loops, we were able to successfully read and store data from each row for the Ticker, Ticker Volumes, Ticker Starting Prices, and Ticker Ending Prices. Once complete, the code for formatting the cells was included within the same macro to automate the formatting when switching between years. You can view the refactored code for the loops and formatting below. 

![Loops_Formatting](https://user-images.githubusercontent.com/102122063/163631260-4550543a-d977-4815-9b24-39db989d8876.png)

After refactoring our code, debugged and tested our new macro to ensure it created the same results, yet faster. We found that the run time decreased to 0.2421875 for 2017 and 0.1835938 for 2018, successfully creating a faster and more efficient code. The run time for both years can be viewed in the images below. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/102122063/163631020-aff0ef3e-0d98-45f1-8394-d622ade48ad7.PNG)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/102122063/163631029-560084a1-53f4-4e9d-b894-8a8bd63afa68.PNG)

## Summary

After reviewing this project, refactoring code in general has it's advantages. It can help make it cleaner and more organized which can lead to improved software efficiency through reduced run time and use of less storage, it becomes easier to read and debug and can allow for faster programming overall. 
Yet, while the efficiency and time for the refactored code improve, one disadvantage that comes with refactoring is how much time the task takes to be complete. There may also be cases of working with larger programming code which can pose the risk of getting lost in the code or not understanding where to go next after making changes.

Looking specifically at our VBA script for the stock analysis, we can see that one important advantage of refactoring the original code was the reduction in run time for the macros. In reference to our results, the run time of our original code decreased from a run time of 0.9453125 for stocks from 2017 with the original code to 0.2421875 seconds. The run time of 0.984375 for stocks from 2018 respectively decreased to 0.1835938 with the refactored VBA code.
