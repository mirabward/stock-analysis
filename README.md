# stock-analysis

## Overview of Project:

For our client, Steve, we took a workbook that had an extensive amount of information regarding the stock market and created our own macros to help analyze this data. Thbiks analysis allowed us to see not only what the daily volume for the stock tickers were, but also the return rate. This information can help Steve determine if certain tickers are work investing in or if certain investments should be declined. Using the macro we created allows us to execute the analysis for thousands of rows of data, repeatedly, in less than a second. Compared to manually calculating the data using formulas, macros is proved to be much more efficient! 

## Results: Compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

In 2017, the stock returns were extremely high. The DQ ticker had an astounding return rate of 199.4% with SEDG close behind at 184.5%. This leads us to believe that DQ and SEDG are tickers that should be invested into, because statistically, you'll receive a better return on investment than investing with another ticker, such as TERP that had a return loss of 7.2%. In 2018, the stock market had a much more difficult year. RUN and ENPH were the only tickers that had a return, both in the 80% range. DQ that previosuly had the highest return now has the lowest return of -62.6%. With this huge drop in return, we would recommend avoiding investing with DQ due to the alarming rate the ticker's return has dropped. In comparison, RUN is the only ticker that had a substantial raise in it's return, leading us to recommend this ticker for investing instead. In my original VBA script, the 2017 code took about 1.2 seconds to run and my 2018 script took about 1.1 seconds. After refactoring the code, my 2017 dropped down to 0.8 seconds and my 2018 script to 0.81 seconds.

## Summary: In a summary statement, address the following questions.

### What are the advantages or disadvantages of refactoring code?
Refactoring your code defintiely has more advantages than disadvantages. While rafactoring your code does require spending more time on your code instead of moving onto a new project, it allows you to create a more efficient macro. While a difference of about .3 seconds doesn't seem like much, when it comes to using a much larger data set (possibly in the millions range), running a macro can take an extensive amount of time. Rafactoring it can help cut down on that time, which leads to much more time saved in the long run than spent on refactoring. Especially if the macro needs to be ran over and over. 

### How do these pros and cons apply to refactoring the original VBA script?
My original VBA script had multiple repeats of activating worksheets, which was not necassary for the macro and eats up memory space. Rafactoring the code allows us to see where we can cut out these unnecessary repeats and catch errors we may not have originally seen when running the beta version of the macro. It also provides a chance to look over your macro and make sure that your code is pulling the correct data that's needed and not anything that could hinder getting the data you're looking for. 
