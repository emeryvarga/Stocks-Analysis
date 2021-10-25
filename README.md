# Stocks-Analysis

I.	Overview

Steve recently graduated with a finance degree and is helping his parents diversify their investment portfolio into a handful of green energy stocks. He is using Microsoft Excel VBA code to automate his analyses and determine whether certain stocks are worth investing in.


II.	Purpose

The purpose of this project is to refactor Steve’s VBA code to collect stock information for the years 2017 and 2018. The refactored code will make his analyses more efficient by taking fewer steps, using less memory, and improving the logic of the code.

III.	Results

The original code and refactored code use the same logic to create an input box, chart headers, ticker array, and activate the appropriate worksheet. However, the refactored code contains a tickerIndex which is used as a reference variable throughout the rest of the code, as seen in Figure 1.


**Figure 1**: Refactored VBA Code using tickerIndex

As we can see in Figure 2a and Figure 2b, the original code takes nearly a second to execute. For the year 2017 and 2018, the code ran in ~0.76 seconds and ~0.78 seconds respectively.
 
**Figure 2a**: Original code run time for the year 2017
 
**Figure 2b**: Original code run time for the year 2018

However, when we refactor the code by using the tickerIndex as a reference, the run time improves drastically, as seen in Figure 3a and Figure 3b. The refactored code takes ~0.11 and ~0.10 seconds for the year 2017 and 2018, respectively.

 
**Figure 3a**: Refactored code run time for the year 2017


 
**Figure 3b**: Refactored code run time for the year 2018

IV.	Summary

Refactoring not only makes code clearer and more organized, but also improves software performance by using up less memory and making it easier to reference in the future. However, the process of refactoring code uses up coders’ time and energy that could be applied elsewhere in the workplace. Before deciding to refactor code, its necessary to determine how often it will be used and to what extent. If the task at hand is part of basic reporting that will be used after employees rotate out of their current positions, refactoring code will likely be beneficial in the long run.

The refactored code in Steve’s stock analysis drastically improved its run time (approximately 0.7 seconds). It will also benefit Steve in the long run because he can use the same code, perhaps with minor tweaks, for his future clients.
