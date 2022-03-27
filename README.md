# VBA of Wall Street

## Overview of Project

### Background
Steve is doing research on the stock market for his parents. I created a workbook that analyzed 12 stocks' performances in 2017 and 2018, but Steve wants to expand the dataset to include the entire stock market over the last few years. I will look to improve on the initial code to potentially help Steve's expanded research. 

### Purpose
The purpose of the analysis was to improve upon a VBA code used to analyze 12 stocks in each year. The code is refactored to make it more efficient by taking fewer steps and improving the logic. To prove this, the refactored code includes the timer to help measure performance against the original code. 

## Results

### Stock Performance: 2017 vs 2018
11 out of 12 stocks performed positively in 2017, while only 2 out of the 12 stocks did so the following year. ENPH and RUN performed well in both years, with RUN improving upon its 2017 performance. 

![This is an image](https://github.com/csambrano/stock-analysis/blob/main/Resources/2017_Stock_Performance.png?raw=true)

![This is an image](https://github.com/csambrano/stock-analysis/blob/main/Resources/2018_Stock_Performance.png?raw=true)

### Execution Times: Initial vs Refactored Script
The refactored script proved to be more efficient by taking fewer steps and improving the logic.

Initial script's execution time for 2017: 


![This is an image](https://github.com/csambrano/stock-analysis/blob/main/Resources/2017_Original_Execution_Time.png?raw=true)

Refactored script's execution time for 2017: 


![This is an image](https://github.com/csambrano/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png?raw=true)

Initial script's execution time for 2018: 


![This is an image](https://github.com/csambrano/stock-analysis/blob/main/Resources/2018_Original_Execution_Time.png?raw=true)

Refactored script's execution time for 2018: 


![This is an image](https://github.com/csambrano/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png?raw=true)

## Summary

### Refactoring Code: Advantages and Disadvantages

Refactoring a script helps improve efficiency of the code, making it significantly faster. This will be helpful for much larger data sets. For example, the refactored script provided a ~.76 seconds improvement on the execution time, which translates to an 81% decrease. A refactored script can make the initial script less complicated and easier to understand. A disadvantage of refactoring code is figuring out issues/errors in your new script. This applied specifically to my situation, in which I ran into a few errors in my refactored VBA script that took some time to unravel and overcome. Human error may also result in drastically different results between the refactored and initial scripts. 
