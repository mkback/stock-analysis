# Stock Analysis with VBA

## Overview of Project

#### Steve is a newly graduated financial advisor and his parents are his first clients. In order to make a smart investment, he wants to analyze a number of stocks' total volume and returns. In this project we used VBA to pull this information together for Steve to look at.   

## Results


Based on the 2017 results, all but one stock saw positive results, but when looking at 2018 results only two saw positive results. Stocks ENPH and RUN both saw positive return in both years so I would recommend Steve invest in these. See below for results each year. 

![Alt Image Text](https://github.com/mkback/stock-analysis/blob/master/Resources/Results.png)

We ran this code twice, the first time it took about .27 seconds to run. After we refactored the code, it took the run time down to .25 seconds. This may only be a small difference, but this is a small piece of code. With larger code running, refactoring can show much more improvement. See below for the run times of the refactored code vs not refactored. 

![Alt Image Text](https://github.com/mkback/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png)
![Alt Image Text](https://github.com/mkback/stock-analysis/blob/master/Resources/VBA_NotRefactored.png)

The difference between the code is making it loop only once and adding an index. See below screen shot for part of the refactored code. In this screenshot you can see one large for loop for i that goes through each line once and a smaller for loop for j that goes through the 12 tickers. 
![Alt Image Text](https://github.com/mkback/stock-analysis/blob/master/Resources/Code.png)


## Summary

- What are the advantages or disadvantages of refactoring code?

An advantage to refactoring code is making the code run faster and more efficiently. Refactoring can make it easier to understand and can be adapted easier. One disadvantage is it could cause some bugs in the code. Making code refactor could be complicated to set up and could cause issues to arise.

- How do these pros and cons apply to refactoring the original VBA script?

Refactoring this code made it more efficient by going through each line once. It also made it more efficient by adding in variables like the tickerIndex. If we needed to update the tickerIndex we just needed to change the variable at the top rather than replacing in each step of the code.  
 
