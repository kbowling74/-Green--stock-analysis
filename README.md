# Stock Analysis using VBA

## Overview of Project
Using refactored VBA code to conduct analysis of a small sample size of stock market data and determining efficiency improvements for time of VBA code.

## Purpose
This analysis was completed to compare performances of particular "green" stocks over the years 2017 and 2018. The original VBA code used for previous analysis was refactored to improve run time of the code for this analysis. 

## Results

### Analysis of Stocks Based on Yearly Performace

With data provided from the Vanderibilt Bootcamp the analysis was performed on 12 "green" stocks finding their return on a yearly basis. The years of 2017 and 2018 were used for measuring returns. All but one stock, TERP, provided a positive return for 2017 with the highest being DQ at a return of 199.4%. <br>
<img width="252" alt="2017_VBA_returns" src="https://user-images.githubusercontent.com/106560606/175837804-b3e0dfdf-2626-48ba-8b2e-3f2386c75940.png">

However, analysis from 2018 showed a very different outcome for these stocks. All but two had a negative return with ENPH and RUN being the exceptions. <br>
<img width="251" alt="2018_VBA_returns" src="https://user-images.githubusercontent.com/106560606/175837806-9f972c13-8a54-45a0-a4ad-33bedf2ccd9f.png">

## Summary
While the results of both the original and refactored code are the same, improvements behind the scenes were made by refactoring the code. The refactored code is easier to read and follow when looking through it. It is a much cleaner presentation once items are condensed and quicker to find a specific line of code. 

The most important element in the analysis that was improved by refactoring the code was the time taken to run the program. <br>
<img width="433" alt="2017_Original" src="https://user-images.githubusercontent.com/106560606/175837803-39071fbb-f1f3-43c6-bbc3-0d5c7c7d28e6.png">
For 2017 the time improved from .5625 sec to .09765625 sec.<br>
<img width="431" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/106560606/175837807-35db47ac-0060-4b8d-8729-869fc846bdf8.png">


<img width="433" alt="2018_Original" src="https://user-images.githubusercontent.com/106560606/175837805-d933fa2a-d0af-4974-97ee-02291f319098.png">
For 2018 the time improved from .578125 sec to .1054688 sec. <br>
<img width="437" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/106560606/175837808-a0642b43-57ff-468d-b547-28834d5d2480.png">

Refactoring code can be beneficial in circumstances such as this where time for processing is reduced. Although this analysis only used 12 stocks, in a situation where there are thousands of data points the improvement of time can be very important. However, since this was such a small size of data the time taken to refactor the code out-weighed the time reduction of the actual analysis. If it is only being used once for a very quick application, it could be detrimental to take that much time to refactor the code to gain milliseconds of improvement. 
