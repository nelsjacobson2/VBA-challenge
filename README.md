# Stock Market Data Analysis VBA
This VBA code is designed to analyze stock market data that is contained within an Excel workbook. The code loops through each worksheet in the workbook and applies the StockData subroutine to each sheet.

The StockData subroutine calculates several statistics for each ticker symbol in the sheet including:

Yearly change in stock price
Percent change in stock price
Total volume of shares traded
After calculating these statistics, the subroutine outputs them to a table in the worksheet with the following headers:

Ticker Symbol
Yearly Change
Percentage Change
Total Volume
The subroutine also applies conditional formatting to the Percentage Change column, with positive values displayed in green and negative values displayed in red.

Finally, the subroutine calculates the following statistics for the entire workbook:

The ticker symbol with the greatest percent increase in stock price
The ticker symbol with the greatest percent decrease in stock price
The ticker symbol with the greatest total volume of shares traded
The RunStockData subroutine runs the StockData subroutine on every worksheet in the workbook.

How to Use
To use this code, simply copy and paste it into a new VBA module in your Excel workbook. Then, run the RunStockData subroutine to analyze the data in all worksheets in the workbook.

Requirements
This code requires Microsoft Excel to be installed on your computer in order to run. Additionally, the workbook being analyzed must contain data on each worksheet in the following columns:

Column A: Ticker symbol
Column B: Date
Column C: Open price
Column D: High price
Column E: Low price
Column F: Close price
Column G: Volume of shares traded
