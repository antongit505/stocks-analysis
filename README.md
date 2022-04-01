# Analyzing stocks for investment

## Overview
A financier is interested in analizing stocks to recommend his clients where they should invest their money.
A dataset was provided with information from twelve different companies in order to analyze their stocks behaviour through different years.
Since the clients are insterested to invest in green energy companies they think that DAQO New Energy Corp (DQ), is the best choice but before they invest, the financier needs to support this decision with data.

The dataset contians information such as Ticker, Date, Open ,High, Low, Close, Adj Close and Volumene.
![](resources/extra_resources/Dataset.PNG)

A program in VBA is required to analize the dataset. 
Three things need to be accomplished:

* Show the time that it takes to make the calculations
* Calculate the TotalVolume which is the total number of shares traded throughout the day, it measures how actively a stock is traded and calculate the Yearly Return which is the percentage difference in price from the beginning of the year to the end of the year.
* Format the results in a readable table.

First we set the variables and the format where the results will be shown.
The program has to be able to analize any year that we require, therefore a variable to ask for the year of interest is created using an inputbox. We initialize the timer and activate the worksheet where we will set the format.

![](resources/extra_resources/Variables.png)





This program has to be able to analize a random number of tickers therefore the code
