# Stock_Market_Analysis
# "The VBA of Wall Street"

# Purpose

The purpose of my VBA script is to:

    * analyze stock market data found in the "Multiple_year_stock_data" spredsheet
    * loop through all the rows within a sheet
    * loop through all the tabs within the overall spreadsheet

and return the following analysis as new columns within each sheet:

    * unique ticker symbols
    * yearly change for each ticker symbol
    * percentage change from opening price at the start of the year to ending price at the end of the year
    * total stock volume

Please note: For the Yearly Change data, the script applied conditional formatting to indicate positive change in greend and
negative change in red.

Also, I added functionality to the script to include three additional analytical data points:

    * stock with the "Greatest % increase"
    * stock with the "Greatest % decrease"
    * stock with the Greatest total volume"

## Detailed Explanation of VBA Script

My script begins with pseudo-code to help me logically plot out what I needed to accomplish.
The first thing I did was activate each sheet within the worksheet in order to loop through all the tab/year data.
Then I called out my variables, set my values (e.g., the counters), determined the last row and initatied the forloop.
At this point, I scripted the "TotalVolume" so that it fell outside of the first "IF" statement in order to calculate
    the cumlative total for a given ticker symbol. The purpose of the first "IF" statement is to find unique ticker symbols. 
Next, I started my overall conditional "IF" statement to find the unique ticker symbol and the yearly change.
Next, I nested a second "IF" statement to calculate the percentage change.
Then I reset the counters before the next loop.
The next "IF" statement applied the conditional formatting to the Yearly Change column.
And finally, the last forloop found the Greatest % Increase/Decrease and Greatest Total Volume within the results of
    the first "IF" statement.
