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
The inital analysis was on the stock DAQO (ticker DQ).  While in 2017 DQ had a 202.0% return on investment it's daily volumes were low.  In 2018 while the volumes increased the return on investment dropped sisgnificantly.  The 2018 return on investment was a low -61.1%.   The investors would do well to consider researching other stock that had a more consistent positive return.

![Alt text](https://github.com/SusanFair/stock-analysis/blob/main/Resources/VBA_Challenge_2017_table.PNG)   

![Alt text](https://github.com/SusanFair/stock-analysis/blob/main/Resources/VBA_Challenge_2018_table.PNG)


### Run Time Comparison
We can see the results of the refactoring in the run times for each the original code and the refactored code.  The refactored code has significant improvements in run time.  

##### Original code run time:
 ![2017 original](https://github.com/SusanFair/stock-analysis/blob/main/Resources/green_stocks_original/green_stocks_original_2017.PNG)  ![2018 original](https://github.com/SusanFair/stock-analysis/blob/main/Resources/green_stocks_original/green_stocks_original_2018.PNG) 


##### Refactored Code run time:
![2017 refactored](https://github.com/SusanFair/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)          ![2018 refactored](https://github.com/SusanFair/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)


## Summary
In the beginning we asked 3 questions about the pros or cons of refactoring.  The answers to these questions provide a clear answer.

### Was the refactored code faster and/or ore efficient to write?
Although new to coding, which may have impacted this analysis, writing the code using the method deployed in the refactoring project seemed longer and more prone to errors due to the complexity in defining the new arrays and indexes.  Familiarity with arrays and indexes may reduce this overall coding times.  Efficiency is another factor.  While it may have taken longer to write the maintainability and flexibility of the refactored code must be considered.  The refactored code with more arrays would allow for more flexibility in maintaining and modifying the code in the future.

### Was the refactored code faster to run?
While it may have taken longer to write the benefits were undeniable when it comes to run times.  The refactored code was significantly faster to run as noted in the above comparisons. This run time benefit would only be amplified with larger datasets.

### Advantages or disadvantages between the two coding styles.
The refactored code in VBA particularly does seem to have advantages.  Using the arrays and allowed us to store data with a single variable name and while changing the index.  Doing so allowed storing of the rests of the analysis e.g.: volumes, starting price and ending price.  Being able to store a number of values per ticker allowed us to only write to the spreadsheet at the end of the analysis.  In the original code the macro would write to the output file after each loop. Reducing the incidence of writing to the results worksheet improves the overall run time.

#### Original code

![Alt text](https://github.com/SusanFair/stock-analysis/blob/main/Resources/Write_original_code.PNG)


#### Refactored Code

![Alt text](https://github.com/SusanFair/stock-analysis/blob/main/Resources/Write_refactored_code.PNG)

Overall refactoring as a general rule has it's advantages in maintainability, flexibility and timing. This also held true for the VBA script.







