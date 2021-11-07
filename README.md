# Stock-Analysis

## Overview
	
The purpose of the stock analysis is to click “Run Analysis” button in All Stocks Analysis worksheet to analyze an entire dataset for year of 2017 or 2018. The analysis provide the best return result to Steve for making the right decision. This project is refactoring the original module exercise to loop through all the data one time in order to run the VBA script faster. 

## Results
	
This VBA challenge start with creating the new worksheet name “All Stocks Analysis” in the excel. In VBA challenge script, start activate the worksheet, create out arrays for ticker volumes, ticker starting prices, and ticker ending prices. As we can see the script below step by step refactoring my code successfully made the VBA script run faster. The results of this analysis run time and length of script is better than the original module project. The difference between module exercise and challenge analysis script is use only one iterator (i) where as module exercise use two iterators (i and j). Therefore, the run time for loop is much more higher than single iterator. 

## Summary
	
The advantages of refactoring code can minimize the run time, use shorten code, and easy to debug when encounter. The length of refactoring code is short which allow readers to follow though easily compare with the normal script. On the other hand, the disadvantages of refactoring when we have big size of application, and not having proper way of test cases for the codes.
	
Let’s compare both 2017 and 2018 macro run time original VBA script with refactoring script. In original VBA script, macro run time for 2017 and 2018 took approximately 2 seconds. In refactoring VBA script, the macro run time only took 0.12 seconds shown below.

<img width="257" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/92540198/140666538-401cd11a-1263-4074-8c75-af5150cdfef3.png">
<img width="257" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/92540198/140666539-d18af663-fe36-4027-af08-9d2561f15a1c.png">
