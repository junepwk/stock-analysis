# VBA of Wall Street

## Overview of Project 

### Purpose

The goal was to assist Steve in analyzing stock data to help his parents commit to the right investment using VBA script in Excel.  The script written compared different green energy stock data for the years 2017 and 2018 by calculating the total daily volume and yearly return for each stock.  The script would prompt users to input the year they would like to analyze and provide the time it took to execute with a single click of a button.  To further expand on this analytic request, a refactored code was written to allow Steve to run analyses on larger stock datasets over a shorter execution time.

## Results

### Stocks Analysis

In 2017, the green energy stock market showed promising growth in yearly return with the exception of TerraForm Power (TERP) who had a -7.2% decrease in value with approximately $139 million in total daily volume.  Steve's parents were interested in Daqo's (DQ) stocks and in the year 2017, DQ had the highest yearly return growth (199.4%) and lowest daily volume of $35 million. 

In 2018, majority of the green energy companies experienced a negative dive in their stock's yearly returns.  DQ held the highest net gain by the end of 2017 but it appears to hold the highest net lost in 2018 with -62.6%.  On the contrary, Enphase Energy (ENPH) and Sunrun (RUN) both yielded net gains.  Additionally, while remaining positive, ENPH's returns were not as abundant as observed the previous year.  On the other hand, RUN's yearly return skyrocketed from 5.5% to 84%.  

![stocks_2017](https://github.com/junepwk/stock-analysis/blob/main/resources/stocks_2017.png)  ![stocks_2018](https://github.com/junepwk/stock-analysis/blob/main/resources/stocks_2018.png)

### Refactored Vs. Original Script

As shown in the pictures below, the refactored script performs significantly faster than the original script.

**Refactored execution time**

![vba_challenge_2017](https://github.com/junepwk/stock-analysis/blob/main/resources/vba_challenge_2017.png) ![vba_challenge_2018](https://github.com/junepwk/stock-analysis/blob/main/resources/vba_challenge_2018.png)

**Original execution time**

![old_script_time_2017](https://github.com/junepwk/stock-analysis/blob/main/resources/old_script_time_2017.png)  ![old_script_time_2018](https://github.com/junepwk/stock-analysis/blob/main/resources/old_script_time_2018.png)

In the refactored script, a variable, tickerIndex, was created with the purpose of being used to access the correct stock's index during For loops.  In addition, unlike the original script, variables for daily volumes, starting prices and ending prices were all initialized as output arrays.  As a result, this script is not limited to just the stocks dataset provide in this module.  The refactored script could gather the same information for larger datasets from different stock markets efficiently. 

```VBA

Dim tickerIndex As Integer
    tickerIndex = 0
    
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

```

## Summary

- What are the advantages or disadvantages of refactoring code?

One of the advantages to refactoring codes is that it broadens the capability of the code. It allows the code to be used to solve a multitude of problems.  Putting the original script side by side to the refactored script, the original script felt crammed with subroutines.  I found myself getting lost scrolling through the codes when I needed to reference a specific part of the code. The refactored script is condensed and easier to read at a glance for coders. Another pro for refactoring codes is the ability to cut execution time short which is ideal when working with an extensive dataset.  

The disadvantage to refactoring is the amount of time spent debugging errors and figuring out the correct line to place a code in order for the loops to run correctly, especially in VBA editor where error messages are vague. Reimaging and restructuring an existing solution and logic was also extremely challenging. 

- How do these pros and cons apply to refactoring the original VBA script?

For this particular instance, the original script worked well with the green energy dataset provided.  It served as a solution for Steve as his parents were interested in investing in a green energy company.  With the refactored code, Steve would be able to utilize the same code to analyze any stock market in the future.  

Aside from the hardship stated above, the disadvantages of refactoring the original script seems to be non-existent or minute as the refactored code improved the overall functionality.  The original script was much easier to write but the refactored code's end result outweighs the cons.

