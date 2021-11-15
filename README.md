# Stock Analysis Using Excel VBA

## Overview

This project was created to analyze collected stock information in 2017 and 2018 using refactored code in Excel VBA. The analysis was used to determine if certain stocks were worth investing in by comparing their return values over the past year of collected data. This data was originally in a similar format with the purpose of this project to make the original code run more efficiently.

## Results
### Analysis

The code is run in nested loops that run through all rows of stock data after the user inputs the given year they would like to analyze. The second column of data shows the toal volume of the stock traded over the course of the given data set. Using an array of the different stocks, the code checks if a row is the first row of a certain stock in the array. If it is, the starting price will be recorded as a variable and the code will run a following if-then check to record the last instance of that stock's price in the data as its ending price. The output of the third column is the percentage gain or loss of each stock in the array over the course of the year.

## Summary
### Benefits and Consequences of Refactoring Code

Taking the time to refactor code makes it more organized and cleaner to read. This can be helpful when debugging and sharing code with others who might not understand messy design decisions. The time saved in having code run less loops can use less system resources as well as improving the end user experience, overall making a better product. However, there isn't always enough time to correctly refactor code which can lead to increased dev time. It's usually prudent to be able to deliver a less-than-perfectly optimized solution compared to missing a release deadline. Refactoring can also lead to unforseen issues with the code, meaning a working solution is now unusable. Overall, refactoring code is great if time and resources allow, but time is the most important resource when deciding if refactoring is the right decision.

### Advantages and Disadvantages of Old vs. Refactored Code
The most apparent advantage of the refactored code is the shortened run time of the program. Checking stocks in the 2017 data sheet went from 0.8 seconds to 0.12 seconds. This saved time running the code and will be helpful as the dataset expands over time and more stocks are added to the array.
