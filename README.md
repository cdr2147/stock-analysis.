# VBA Challenge
## Overview of Project
Analyze stock market trends across multiple years
### Purpose
Refactor VBA code to loop through stock data and return the Total Daily Volume and Rate of Return of individual stocks for a given year. Using code, the user can simply click a button and easily see what stock tickers had a positive or negative return for the given year. The code has been refactored to loop through the data once and run more efficiently than looping multiple times, and could be updated to account for additional stock tickers.
## Results
After running the analysis on the data of 12 stock tickers, one can see that the rate of return on most stock tickers in 2017 was better than the rate of return in 2018. In 2017, all but one of the stocks had a positive rate of return, but in 2018 only two stocks had a positive rate of return. Using refactored code allowed the script to run more quickly and it can updated to account for additional stock tickers in a given data set. 

Instead of looping through the data multiple times for each ticker, in the refactored code a tickerIndex was created and set to zero so that the code could loop through the data only one time and produce results more quickly. Reducing run time is useful if additional stock tickers or larger set of data needs to be analyzed. 

After creating the tickerIndex, we set three output arrays, initiated tickerVolumes to zero and looped through the data once. As the code loops through the date, the tickerVolumes increases the volume for a given stock ticker, and the code retrieves the starting and ending price of a given ticker so that we can output a rate of return for each stock. See below for the refactored code.
![Code Refactored](https://user-images.githubusercontent.com/99205688/156932493-67cbf87b-01a5-4280-8d05-518901038372.JPG)

The refactored code ran in less time than the original code, which can be seen in the run time screenshots below. Additionally, one can see the difference in rates of return of the given stocks in 2017 compared to 2018.

**Refactored Run Time**
![VBA_Challenge_2017](https://user-images.githubusercontent.com/99205688/156932526-8b910b10-2e9f-4345-93fc-454ad3831f90.JPG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/99205688/156932521-010941b9-1409-4b03-a33a-82d436351d9e.JPG)

**Original Run Time**
![Run Time_2017_Original](https://user-images.githubusercontent.com/99205688/156932541-606eb91f-3e22-41a6-9521-da93b20d8188.JPG)
![Run Time_2018_Original](https://user-images.githubusercontent.com/99205688/156932539-5b2648c2-6c68-4f5b-b07f-30634ec4f686.JPG)

### Advantages and Disadvantages of Refactoring Code
Refactoring code has several advantages and disadvantages. Refactoring code can make code clean, run more efficiently, and easier to maintain. Refactoring code can simplify code and make it easier for the user and future users to understand. It can also make code more scalable and reduce errors for future projects. Nevertheless, refactoring code can have some disadvantages, including that it can be quite time consuming to test and edit. The reduction in time when running the refactored code may not be time saving enough to warrant the time commitment put into the refactoring itself. 

### Advantages and Disadvantages of this VBA code
The advantages of refactoring this particular code is that the code is cleaner and simplified, as it only needs to loop through the data once to produce results. This brought down the run time and is useful in the event that the client wishes to analyze larger sets of data with additional stock tickers. However, a disadvantage of refactoring this specific code is that if the data set is very large, there are still some time consuming updates that need to be made to the script to update the array for additional tickers. 
