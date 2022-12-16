# stock-analsis

## Overview of Project
This project will compare coding styles to assess the most efficient way to perform a stock performance analysis.  Since this sub routine may be run hundreds of times on thousands of stock combinations, effficiency is paramount.  We will compare original coding vs the result of refactoring the code.  This analysis was designed to answer the following questions:
* Was the refactored code faster and/or ore efficient to write?
* Was the refactored code faster to run?
* Advantages or disadvantages between the two coding styles.

### Background
The initial analsis of data provided information on a variety of green energy stocks over the years 2017 and 2018.  While there was primary interest in the DAQO New Energy Corp it was deemed prudent to review a variety of stocks and their returns on investment.  An initial analysis was performed.  It was felt that there was perhaps a better way to perform this same analysis and as a result this refactoring project was initiated.

## Approach and Analysis
While the two sub routines performed the same actions AllStocksAnalysis (the initial code) and AllStocksAnalysisRefactored (refactored code) took slightly different approaches.   The refactored code relied much more on arrays, creating arrays for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. 

## Results
We can see the results of the refactoring in the run times for each the original code and the refactored.

### Original code:

![2017 original](https://github.com/SusanFair/stock-analysis/blob/main/Resources/green_stocks_original/green_stocks_original_2017.PNG) 

![2018 original](https://github.com/SusanFair/stock-analysis/blob/main/Resources/green_stocks_original/green_stocks_original_2018.PNG)


### Refactored Code:
![2017 refactored](https://github.com/SusanFair/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

![Alt text](https://github.com/SusanFair/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

Benefit:
Able to store data using a single variable name with a changing index
The variable plus index information can be stored.  This means that the macros doesn't have to write each result for each index as they are defined.  Write activities are saved until the end.  Reducing the incidence of writing to the results worksheet improves the overall run time.



### Summary
detailed statement on the advantages and disadvantages of refactoring code in general 
detailed statement on the advantages and disadvantages of the original and refactored VBA script
### Limitations of this dataset?




### Conclusions of analysis of Outcomes based on Launch Date?
















Kickstarting with Excel

## Overview of Project
This project is an analysis performed for Louise.  She is a budding entrprenuure who is interested in crowdfunding campaigns of verious types in the field of theatrical presentations.  While she started a campaign for a play called "Fever", she is now moving on with other adventures.  Wisely Louise is making decisions based on realiable and relevant historical data.  A database of previous campaigns has been assembled to be used as a basis for the analysis. She has asked us to perform an analysis on this data to provide insights for her future campaign.

## Overview of Project
As Louise moves to her next project we will provide her with specific analysis on results of different campaigns' success factors based on campaign lauch date.  Is there a better season or month to start a campaign?  Did campaigns in certain months have more success? Also Louise wants more detailed analysis on the success factors based on the funding goals of previous campaigns.  How much is too much?  Do larger goals increase or decrease the probability of a successful campaign.  Is there a sweet spot between lower and higher goals that will give her a clearer pathway to a successful campaign.

## Analysis and Challenges

### Analysis of Outcomes Based on Launch Date
In the analysis based on launch date we viewed the data set by the month of the campaign launch and looked at the "Successful", "Failed", and "Canceled" campaigns.  The scope of the analysis was theatrical campaigns over the complete timefame of the data available. Using a pivot table to narrow and filter the data and the accompanying chart we can clearly show the results of this analysis. Using these results we can provide Louise with the planning information to help her with her future campaign.

The raw data can be found in the Excel workbook linked below on the workshet entitled "Kickstarter". The data Analysis can be found on the worksheet entitled "Theater Outcomes by Launch Date".  This file can be downloaded.
*   https://github.com/SusanFair/kickstarter-analysis/blob/main/Kickstarter_Challenge.xlsx


### Analysis of Outcomes Based on Goals
In the analysis based on campaign funding Goal sizes we sorted the data set by funding dollar amounts within specified monitary ranges. The scope of the analysis was theatrical play campaigns which were Successful, Failed or canceled over the complete timefame of the data available. Using the analysis tools at hand this project will provide Louise with guidance on the most successful campaign goal size for her future campaign. Using COUNTIFS formulas to narrow and filter the data we derived a table and the accompanying chart so as to clearly show the results of this analysis.

The raw data can be found in the Excel workbook linked below on the workshet entitled "Kickstarter". The data Analysis can be found on the worksheet entitled "Outcomes Based on Goals". This file can be downloaded.

* https://github.com/SusanFair/kickstarter-analysis/blob/main/Kickstarter_Challenge.xlsx

The goal ranges used were as follows:

* Less than 1000
* 1000 to 4999
* 5000 to 9999
* 10000 to 14999
* 15000 to 19999
* 20000 to 24999
* 25000 to 29999
* 30000 to 34999
* 35000 to 39999
* 40000 to 44999
* 45000 to 49999
* 50000 or More

### Challenges and Difficulties Encountered :
#### Challenges - Analysis of Outcomes Based on Launch Date
The challenge of drilling down to months instead of years was solved by adjusting the rows category in the PivotTable Fields configuration. 

![PivotTable Rows field](https://github.com/SusanFair/kickstarter-analysis/blob/main/Resources/PivotTable_Field_Config.PNG)

#### Challenges - Analysis of Outcomes Based on Goals :
The challenge in this analysis was the new formula for COUNTIFS where there were multiple criteria used to filter the count.  The provided video gave helpful guidance in creating the required filters.

## Results

### Conclusions of analysis of Outcomes based on Launch Date?
From the analysis we can see that Louise is on the right track!!  Using the event category of Theatre there were more Successful campaigns than Failed or Canceled campaigns.  Looks like Theatre campaigns are hot right now!

Other valuable conclusions can be made.  The campaigns that launched in the month of May had the highest number of successful campaigns.  This was closely followed by the months of June and July.  It can also be noted that November and December had the lowest number of successful campaigns. 

As a general observation, the months of May and June had the highest number of Successful campaings however they also had the highest number of Failed campaigns.  In general those months simply had the highest number of campaigns overall.  Further analysis of other factors would therefore be necessary to discern specific success factors.

The following chart will show the overall results:

![Outcomes Success based on Launch Date](https://github.com/SusanFair/kickstarter-analysis/blob/main/Resources/Theater_Outcomes_vs_Launch.png)


### Conclusion of anlysis of Outcomes based on Goals?
The analysis for successfull outcomes based on campaign goal sizes was a bit broader.  Campaign goals below $1000 had high success rates with 75% being successful.  However other ranges saw similar rates with goals between 35000 to 39999 and 40000 to 44999 each having a success rate of 67%.  It could then be concluded that either lower goals (less than $1000) or midrange goals (between $35000 and $49999) would have the highest chances of success.  It was also evident that the higher dollar ranges of goals had distintcly lower success, with ranges over 45000 having a success rate of less than 14%. It can also be noted that the failed campaigns were the mirror image of the successful ones.

The following chart will show clearly the overall results:

![Outcomes Success based on Launch Date](https://github.com/SusanFair/kickstarter-analysis/blob/main/Resources/Outcomes_vs_Goals.png)


### Limitations of this dataset?
Thes is one clear limitation of this dataset.
* Current data - In the Theater category the most recent year of data is 2017 with only 31 campaigns in that year.  Having more current data, reflecting the current economy, would be advisable to gain meaningful analysis.


### Options for further analysis
Since in the Theater Outcomes by Launch Date showed that the months of May and just had the highest number of Successful campaings as well as the highest number of Failed campaigns further analysis would be helpful.  An analysis of the highest % of campaigns per month would be an interesting analysis.  This would remove the prevelence of campaign starts and provide an actual success % by month.

Another interesting and possibly helpful analysis would be on the number of backers and how this affects the success or failure of campaigns.