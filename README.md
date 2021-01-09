# Module-2-Challenge-Stock-Analysis
## Overview of Project
This project refactored code that was created to list tickers, trade volume and rate of return for various stocks contained in the provided data.
### Purpose
The original code looped through each ticker to capture the information required whereas the refactored code looped through the information once to capture the same information.
### Results
The data contained information for both 2017 and 2018.  
#2017
The original code ran in 1.035156 seconds versus the refactored code which ran in 0.3242188 seconds.  Although the difference is not significant it is clear that the refactored code is more efficient. 

https://github.com/bedwardssmith/Module-2-Challenge-Stock-Analysis/blob/main/Original%20Code%202017.png
https://github.com/bedwardssmith/Module-2-Challenge-Stock-Analysis/blob/main/VBA_Challenge_2017.png

#2018
The original code ran in 1.0625 seconds versus the refactored code which ran in 0.22890625 seconds.  As with 2017 the difference is not significant, however, the refactored code is more efficient.

https://github.com/bedwardssmith/Module-2-Challenge-Stock-Analysis/blob/main/Original%20Code%202018.png
https://github.com/bedwardssmith/Module-2-Challenge-Stock-Analysis/blob/main/VBA_Challenge_2018.png

### Summary
A first attempt at writing code to accomplish a specific goal may not necessarily be the most efficient or even logical method for achieving the results.  Refactoring original code allows for a more logical approach to ensure efficiency and manageability of the code as the basics are now documented.    In my view there are no disadvantages to refactoring the code provided the funcionality of the original code remains.

The results in this anlaysis shows a difference of less than 1 second in running the VBA scripts.  This is an insignificant difference, however, it is also relative to the amount of data.  Had the number of stocks been significantly higher our refactored code would have saved a significantly greater amount of time.  This is due to the original code looping through the data for each stock to capture the required information rather than looping through the data once.  In my view the only disadvantage of the refactored code is that it is not as intuitive to write.

