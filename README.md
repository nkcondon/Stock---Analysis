# Kickstarting with Excel

## Overview of Project
	Provide a stock performance analysis of 12 stocks for years 2017 and 2018. The analysis is to identify daily total volumes and yearly return rate. Both years of 
daily total volumes were within 5% of each other but the return rates in comparison was significanty different.  In addition, it was important for the customer to be 
able to run the analysis many times and be able to select a certain year.

### Purpose
	The purpose is to review the green stocks dataset, covert to VBA and write a macro to run the All Stocks Analysis for the 12 stocks. Arrays sand loops were used to 
quickly run through over 3000 lines of data. Daily total volumes were added for each stock and the yearly return was calculated using ending price over starting price. 
Conditional color formatting was used for the ease of recognizing positive (green) and negative (red) yearly returns of the stock.
The program produced a message for the enduser to select a 2017 or 2018 for the results.


## Results
### Stock Performance

Best return on for the 12 stocks was in 2017. In 2017, 92% of the stocks had a positive return for yearly return in comparison to 2018. In 2018, the stocks
did poorly; only 2 stocks showed a positive return.  The total volumes traded was within 5% of each year (over 3,100,000,000 shares). TERP stock was
the only stock that had a negative return in both years




##Summary
### Refactoring Code in General
Advantages of refactoring code are improvement of design, function, efficiency and code readability; decrease of complexitiy and identification of bugs.
Disadvantages of refactoring is that it can be time consuming and increase cost; possibly creating bugs in the code or getting somewhere and not
be able to get out of the code.



### Refactoring Code in VBA Script
When reading about refactoring code, my experience between the original and refactored script was a little frustrating.  Both scripts were ran quickly;
original 0.99 secs / refactored 1.2 secs - not much of a significant difference.  Adjusting the coding for the refactored was time consuming (creating a new variable and using it) and I boxed 
myself in on some of the coding and was not sure how to get out. Maybe with more experience, my refactoring will improve and I utilize more of its advantages.


