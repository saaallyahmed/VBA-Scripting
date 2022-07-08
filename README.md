# The VBA of Wall Street

## Background

Our team  used VBA scripting to analyze real stock market data. Depending on our comfort level with VBA to challenge ourselves with a few of the challenge tasks.


### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - Use this while developing your scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.

### Stock Market Analyst

![stock Market](Images/stockmarket.jpg)

## Instructions

* We created a script that will loop through all the stocks for one year and output the following information:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* You should also have conditional formatting that will highlight positive change in green and negative change in red.

* The result should look as follows:

![moderate_solution](Images/moderate_solution.png)

## In Adcance

* Our solution also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:

![hard_solution](Images/hard_solution.png)

* We Made the appropriate adjustments to our VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

## Other Considerations

* We used the sheet `alphabetical_testing.xlsx` while developing our code. This data set is smaller and allowed us to test faster. our code runs on this file in less than 3-5 minutes.

* We made the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with a click of the button.
