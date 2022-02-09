# Stock-Analysis

Project Overview
Background
Steve asked me/us(we) to analyze a group of 12 green stocks to support his parent’s investment decisions. To do so, we designed an interactive, user friendly, workbook using Visual Basic Application (VBA) within Excel to provide each stocks annual volume and return on investment .

He was able to analyze each stock at the click of a button and now wanted  to go with his research beyond the 12 green stocks.

Purpose
Steve wants to analyze a higher number of stocks and we are trying to help him . This may increase the amount of time it takes the analysis to produce results and we would  like to maintain or, improve it. Now  to improve the workbooks efficiency by refactoring the VBA coding and to  ensure that we are in right direction, we will have to  compare the new execution(refactored) time with the original workbook
.

#Results
Refactoring the Code
To make the code more efficient according to instructions we , created 3 new arrays: -tickerVolumes(12) to hold volume -tickerStartingPrices(12) to hold starting price -tickerEndingPrices(12) to hold ending price

The above 3 arrays store performance data for each stock when a for loop runs analysis on them. The tickers array that we created in the original establishes a ticker symbol that can be called on for each stock.

Matching the 3 performance arrays with the ticker array is done by using a variable called the tickerIndex.

Now that we have created these arrays, we used Nested For Loops and variables to loop through the data and complete the analysis.

See the Refactored vs Original coding in VBA module 2 vs VBA module 1 script in excel.

#Results:

2017 vs 2018 stock Performances -

there is vast noticeabut of the 12 stocks,noticeable difference in 2017 PERFORMANCE of GREEN STOCKS vs 2018.
Only 2  of the 12 stocks had maintained Positive Returns for both years.
All other stocks mostly had decline in the volume.

Steve should look at both economic and industry related influences before advising his parents on their investment decision. They would  be better off with another industry.


#Execution time
Improving the efficiency of the code was a successas per the results Execution time improved from 0.7734375 seconds to 0.15625 seconds for 2017, and, 0.765625 to 0.1054688 for 2018 which is obviously more than 50% reduction in execution time
#Summary
Advantages of refactoring code
The obvious advantage of refactoring code is that it makes it more efficient if you get it right. An more than 50% reduction in execution time can be huge if analyzing thousands of rows of data.

Disadvantages of refactoring code
A huge risk with refactoring is that your errors may destroy an already working code. It is highly recommended that you save your original code and any changes you make frequently in case you run into any issues. That way you can always go back a step without needing to start completely over, I did have issue withnested loops issues during refactoring and found that using the msgBox script made it easier that was in Module 2 formated already  to fill in appropriatly, as well as, run performance outputs individually helped me identify  my errors at steps.