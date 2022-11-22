# Stock Analysis (VBA Challenge)

## Overview of Project

### Purpose
Explain Purpose

## Results

### Stock Performance
After running this code it’s clear the industry at a whole changed quite a bit between these 2 years. Without looking at 2017 or 2019 as a comparison, it’s hard to determine if the green market boomed in 2017 or took a large hit in 2017. The total volume for all green stocks did not change much as it was only a $200 million increase to a $3.1 billion dollar market. There was a more dispersed market in 2018 compared to 2017 where 2 stocks handled most of the volume.

Steve’s parents did a great job investing in DAQO in 2017 with the stock providing the largest return across these green stocks in the analysis with a 199.45% return. This stock was also the least traded in 2017 with the lowest Total Daily Volume. Though in 2018 DAQO lost a lot of its momentum with the lowest return across the green stocks with a -62.60% return. DAQO also remained low traded in 2018 compared across these stocks. Steve’s parents would have benefited in taking the large return in 2017 and diversifying their money in 2018. Though they would have likely still have seen a negative return it would have been a lot smaller.

For further analysis I would start to look at 2019 and possibly look into 2016 to understand some of the volatility of the market and DAQO. I would also suggest looking at other characteristics of ENPH to better understand why their growth didn’t fluctuate almost at all through 2017 and 2018. They were clearly the best stock across the 2 years, providing the second highest return in 2017 and the highest in 2018. The market also followed this stock in 2018 with it being the highest traded stock in 2018.

'![Test](path)


### Refactoring the code
Explain difference

### Run-time Comparison
As seen in the pop up message above in the Stock Performance section, refactoring the code was able to generate a quicker return time which helps Steve conduct his analysis faster. The speed of this report increased by over 500% for both years. The original code’s average return time was 1.13 seconds while the refactored code averaged 0.21 seconds. The goal when refactoring the code is to keep the run-time minimized as more values are added to the analysis.

| Year & Code | Run-Time (in seconds) |
| --- | --- |
| 2017 Original | 1.09 |
| 2018 Refactored | 0.20 |
| 2018 Original | 1.17 |
| 2018 Refactored | 0.22 |

## Summary

### Advantages and disadvantages on Refactoring Code
The main advantages to refactoring code are that the code will operate more efficiently and will be easier to follow in the future. This can save time spent running the code and will make it easier for others to understand and update the code in the future. The main disadvantages of refactoring is it takes time to complete and you’re editing a working script with a chance of error. If you have time to refactor it can be beneficial but if you’re one tight timeline this might not be the best idea especially because rushing through the refactoring increases the chances of an error. If you save out the original code before beginning you can always revert back to that code if any issues arise. 

### Impact on VBA Challenge Script from Refactoring
For the script in this challenge, the code does operate more efficiently and should be easier to follow and add to in the future. The advantage of a quicker run-time is not that significant now, but could be substantial if more information is needed for analysis. This could be if we want to add 2019 data or more if more values are needed to be added for analysis. Not only will adding to this in the future be easier, but as more data and values are analyzed the run-time of the code will stay minimized. For now the code does decrease the run-time by over 500% but in reality that’s only 1 second in time. 

The disadvantage here was also less significant because there was plenty of time to test and complete the refactoring of the code. It was definitely challenging to complete the refactoring of code with plenty of debugging done to ensure the code would run correctly. This would have been very difficult to complete if Steve had a very high turnaround time request on the refactoring.

