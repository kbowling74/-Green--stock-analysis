# Stock Analysis using VBA

## Overview of Project
Using refactored VBA code to conduct analysis of a small sample size of stock market data and determining efficiency improvements for time of VBA code.

## Purpose
This analysis was completed to compare performances of particular "green" stocks over the years 2017 and 2018. The original VBA code used for previous analysis was refactored to improve run time of the code for this analysis. 

## Results

### Analysis of Stocks Based on Yearly Performace

With data provided from the Vanderibilt Bootcamp the analysis was performed on 12 "green" stocks finding their return on a yearly basis. The years of 2017 and 2018 were used for measuring returns. All but one stock, TERP, provided a positive return for 2017 with the highest being DQ at a return of 199.4%. <br>

However, analysis from 2018 showed a very different outcome for these stocks. All but two had a negative return with ENPH and RUN the exceptions. <br>



## Summary
While the results of both the original and refactored code are the same, improvements behind the scenes were made by refactoring the code. The refactored code is easier to read and follow when looking through it. It is a much cleaner presentation once items are condensed and quicker to find a specific line of code. 

The most important element in the analysis that was improved by refactoring the code was the time taken to run the program. <br>

For 2017 the time improved from .5625 sec to .09765625 sec.<br>

For 2018 the time improved from .578125 sec to .1054688 sec. <br>

Refactoring code can be beneficial in circumstances such as this where time for processing is reduced. Although this analysis only used 12 stocks, in a situation where there are thousands of data points the improvement of time can be very important. However, since this was such a small size of data the time taken to refactor the code out-weighed the time reduction of the actual analysis. If it is only being used once for a very quick application, it could be detrimental to take that much time to refactor the code to gain milliseconds of improvement. 
