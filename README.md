# Stock Analysis with Excel & VBA

## Overview of Project

We're working with our client Steve to perform a stock analysis of various companies he is interested in. We decided to base our analysis around the total trading volume and annual returns for 2017 and 2018. We'll code our analysis process using VBA so that we can automate the analysis for Steve to decide what stocks to invest in.


### Results 

We build the script with flexibility. The user of the script can enter the specific year for the analysis to be done by entering into an InputBox:

```
yearValue = InputBox("What year would you like to run the analysis on?")
```

Loops were then used to add the volume together with each stock and determine the price change from the start of the year to the end of the year. 

Lastly, I added formatting to display the returns as a percentage and to highlight positive returns as green while negative returns as red.

Based on the results, we can clearly see all stocks in 2017 aside from TERP had a positive return. In contrast, 2018 was a difficult year for the stocks, and only ENPH and RUN had similar positive returns. When looking at both 2017 and 2018, ENPH is clearly the best-performing stock.

<img src="https://raw.githubusercontent.com/hdolci/stock-analysis/main/resources/VBA_Challenge_2017%202.png" width="50%" height="50%"/><img src="https://raw.githubusercontent.com/hdolci/stock-analysis/main/resources/VBA_Challenge_2018%202.png" width="50%" height="50%"/>

The execution time for the script is about .54 seconds for both the original and refactored code. It is worth noting that the script takes 1 second to complete upon a fresh Excel boot up. Screenshots are displaying performance for running the script after we just opened Excel vs. subsequent runs.


<img src="https://raw.githubusercontent.com/hdolci/stock-analysis/main/resources/VBA_Challenge_2017.png" width="50%" height="50%"/><img src="https://raw.githubusercontent.com/hdolci/stock-analysis/main/resources/VBA_Challenge_2018.png" width="50%" height="50%"/>

## Summary

Generally speaking, refactoring code can optimize speed during run time, and the benefit can be felt with larger data sets. Refactoring code can also lead to reconsidering other ways to solve the problem with a more efficient solution.

Sometimes the code can already be optimized and the time spent to look for improvements can be considered wasted. Although debatable, it is also possible to reduce the code to the point where it would be hard to read.

In this scenario, the main advantage to refactoring the code was that it is easier to read. This code can be revisited years from now if Steve needs to make adjustments, and a different developer can quickly understand the code and make changes.

The disadvantage will be the time commitment if no other changes are to be made. Lastly, the performance difference was negligible at this stage for the size of the dataset.
