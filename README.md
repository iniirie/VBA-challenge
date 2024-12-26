# VBA-challenge
DU_BCS_Module_2_challenge

StockAnalysis VBA Script
This repository contains a VBA script designed to perform stock analysis for multiple worksheets in an Excel workbook. The StockAnalysis processes stock market data, calculates yearly changes, percentage changes, total volumes, and identifies key statistics like the greatest percentage increase, decrease, and volume for each stock ticker.

Purpose
The script is intended to automate the analysis of stock data. It calculates the following for each ticker:

Yearly Change: The difference between the stock's opening and closing price.
Percent Change: The percentage change in stock price over the year.
Total Stock Volume: The total trading volume for the stock over the given period.
Greatest Increase/Decrease/Volume: The stock ticker with the highest increase, decrease, and volume.
The script then outputs these results into a new summary table for each worksheet.

Features
Loops through each worksheet in the workbook, analyzing data for each stock ticker.
Calculates:
Yearly change (Closing price - Opening price)
Percent change (Yearly change / Opening price)
Total volume (Sum of volumes for the year)
Identifies and highlights the following:
Greatest percentage increase
Greatest percentage decrease
Greatest total volume
Outputs results in a summary table with color-coded yearly changes (green for increase, red for decrease).
Generates a "Greatest Values" summary section with the highest increase, decrease, and volume.
How to Use

Steps to Run the Script
Open the Workbook: Ensure that your Excel workbook contains the stock data in the expected format.

Open VBA Editor.

Insert a Module: In the VBA editor, go to Insert > Module.

Paste the Script: Copy the entire StockAnalysis script and paste it into the new module.

Run the Macro: Press F5 or go to Run > Run Sub/UserForm to execute the script.

View Results: The results will be written to new columns on each worksheet, with a summary of the greatest changes and volumes displayed at the end of each sheet.


The script will generate the following results for each worksheet:

Summary Table (Columns I to L):
Ticker
Yearly Change
Percent Change (formatted as percentage)
Total Stock Volume
Greatest Values Summary (Columns O to R):
Greatest % Increase
Greatest % Decrease
Greatest Total Volume
The cells will be color-coded:

Green for positive yearly change (increase)
Red for negative yearly change (decrease)

Greatest Values:
Metric	Ticker	Value
Greatest % Increase	AAPL	20.00%
Greatest % Decrease	MSFT	-10.00%
Greatest Total Volume	AAPL	500,000
Error Handling
If the data is missing or improperly formatted (e.g., non-numeric values in the "Open" or "Close" columns), the script may not function as expected.
Make sure that each worksheet contains the expected stock data in the correct columns.
