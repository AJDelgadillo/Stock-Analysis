# VBA_Challenge Written Analysis
## Overview and Purpose:
In this challenge we are given a dataset consisting of stock prices and volumes recorded over two years. The data in organized on two sheets, separating data recorded in 2017 and data recorded in 2018. The goal of the challenge is to create a VBA Macro calculating the total daily volume and return percentage of each type of stock over each year, that runs more efficiently than the Macro we created while completing Module 2. 

### Methods:
Our method of finding which Macro runs more efficiently is using the timer function to find how many seconds it took for the Macro to run. The timer begins after inputting which year, 2017 or 2018, we want the Macro to run for in the InputBox. The timer stops at the very end of the Macro, after the results have been entered in the Stock Analysis sheet and the sheet has been formatted. 

## Results:
By following the guidelines given in the Challenge Starter Code I was able to code a Macro that runs faster than the AllStockAnalysis Macro created while following the Module 2 instructions. 
When running the AllStockAnalysis Macro for the 2017 sheet the analysis was completed in 0.4453125 seconds. This can be seen in the image titled Unrefactored_Code_2017: 
![Unrefactored_Code_2017](Unrefactored_Code_2017.png) 

When the refactored Macro, AllStockAnalysisRefactored, was run for the 2017 sheet the analysis was completed in only 0.296875 seconds. This can be seen in image titled VBA_Challenge_2017: 
![VBA_Challenge_2017](VBA_Challenge_2017.png) 

The same trend was seen when running both Macros for the 2018 sheet.
When running the AllStockAnalysis Macro for the 2018 sheet the analysis was completed in 0.40625 seconds. This can be seen in the image titled Unrefactored_Code_2018: 
![Unrefactored_Code_2018](Unrefactored_Code_2018.png)

When the AllStockAnalysisRefactored Macro was run for the 2018 sheet the analysis was completed in only 0.3007812 seconds. This can be seen in image titled VBA_Challenge_2018: 
![VBA_Challenge_2018](VBA_Challenge_2018.png)  

From these results I concluded that when running for the 2017 sheet the AllStockAnalysisRefactored Macro completed 33% faster than the AllStockAnalysis Macro. For the 2018 sheet the refactored Macro completed 26% faster than the original Macro. These figures show that I was able to successfully meet the goal of the challenge by creating a Macro that runs more efficiently than the original Macro created while following the Module 2 lessons. 

## Summary:

### Advantages and Disadvantages of Refactoring Macros
From this challenge we learn the importance of refactoring code. Being able to edit and re-work code allows us to create Macros that run more quickly and smoothly. This is an advantage because it can help avoid any frustration that a client may feel if the script was taking a longer time to run. 

While writing the code I tried to run it several times to make sure that I was using the correct functions and syntax. In the events where I wrote something incorrect the Macro was stalling and taking up to 20 seconds to complete. In these moments I assumed that I wrote something incorrectly or the excel page was stuck. I could imagine that if a client was experiencing a Macro running this slowly they would get frustrated and concerned as well. Even if the Macro was to run successfully, this lag in time may be concerning for the client.  

The only disadvantage I could see in refactoring code is that the person writing it could make the script more complicated than it needs to be. As we learned in class, there are many ways to write pieces of code that will give us the same outcome; although this is a good thing it could possibly leave more room for error while creating and editing Macros. 

### Advantages and Disadvantages of Original and Refactored Macros
Both Macros, AllStocksAnalysis and AllStocksAnalysisRefactored have their advantages. The biggest advantage being that they are utilizable to analyze an infinite amount of datasets organized the same way as the 2017 and 2018 sheets. By using the InputBox and not making the Macros specific for only one sheet, they could be utilized for any future analysis in the upcoming years. If our client was to create a sheet consisting of stock data from the year 2019 they would not need to change anything about either Macro to run the script for that new dataset. They would simply need to input “2019” in the InputBox. This would be possible for any upcoming year. 

The AllStocksAnalysisRefactored Macro has an additional advantage of running more quickly and efficiently than its predecessor, AllStocksAnalysis.

The only disadvantage for these Macros would be that they are written very specifically for the data given in the 2017 and 2018 sheets and would need to be entirely refactored if we wanted to consider any other information added onto the dataset. For example, if we wanted to keep track of how many units of individual stock the client was trading we would need to add additional columns to our data sheets and refactor our Macros to utilize the new data. Also, if the client began trading a new type of stock we would need to update the tickerIndex array and the ticker index to recognize that new data. 
