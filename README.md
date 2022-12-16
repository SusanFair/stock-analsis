# stock-analsis

## Overview of Project
This project will compare coding styles to assess the most efficient way to perform a stock performance analysis.  Since this sub routine may be run hundreds of times on thousands of stock combinations, effficiency is paramount.  We will compare original coding vs the result of refactoring the code.  This analysis was designed to answer the following questions:
* Was the refactored code faster and/or ore efficient to write?
* Was the refactored code faster to run?
* Advantages or disadvantages between the two coding styles.

### Background
The initial analsis of data provided information on a variety of green energy stocks over the years 2017 and 2018.  While there was primary interest in the DAQO New Energy Corp it was deemed prudent to review a variety of stocks and their returns on investment.  An initial analysis was performed.  It was felt that there was perhaps a better way to perform this same analysis and as a result this refactoring project was initiated.

### Approach and Analysis
While the two sub routines performed the same actions AllStocksAnalysis (the initial code) and AllStocksAnalysisRefactored (refactored code) took slightly different approaches.   The refactored code relied much more on arrays, creating arrays for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. 

## Results

### Initial Stock Analysis
The inital analysis was on the stock DAQO (ticker DQ).  While in 2017 DQ had a 199.4% return on investment it's daily volumes were low.  In 2018 while the volumes increased the return on investment dropped sisgnificantly.  The 2018 return on investment was a low -62.6%.   

![Alt text](https://github.com/SusanFair/stock-analysis/blob/main/Resources/VBA_Challenge_2017_table.PNG)

![Alt text](https://github.com/SusanFair/stock-analysis/blob/main/Resources/VBA_Challenge_2018_table.PNG)


### Run Time Comparison
We can see the results of the refactoring in the run times for each the original code and the refactored code.

#### Original code run time:
 ![2017 original](https://github.com/SusanFair/stock-analysis/blob/main/Resources/green_stocks_original/green_stocks_original_2017.PNG)  ![2018 original](https://github.com/SusanFair/stock-analysis/blob/main/Resources/green_stocks_original/green_stocks_original_2018.PNG) 


#### Refactored Code run time:
![2017 refactored](https://github.com/SusanFair/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)          ![2018 refactored](https://github.com/SusanFair/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

Benefit:
Able to store data using a single variable name with a changing index
The variable plus index information can be stored.  This means that the macros doesn't have to write each result for each index as they are defined.  Write activities are saved until the end.  Reducing the incidence of writing to the results worksheet improves the overall run time.

### Summary
detailed statement on the advantages and disadvantages of refactoring code in general 
detailed statement on the advantages and disadvantages of the original and refactored VBA script











