Stock Analysis VBA Script

Overview

This VBA script analyzes stock data stored in an Excel workbook. It loops through each worksheet, calculates quarterly changes, percentage changes, and total volumes for each stock ticker, and outputs the results into a summary table. Additionally, it identifies stocks with the greatest percentage increase, decrease, and maximum trading volume.

Features

	•	Main Subroutine: Main()
	•	Loops through each worksheet in the workbook, formats data using FormatAllSheets, and processes stocks using LoopThruStocks.
	•	Formatting Sheets: FormatAllSheets(ws As Worksheet)
	•	Formats specified columns as numbers with two decimal places across all worksheets.
	•	Processing Stocks: LoopThruStocks(ws As Worksheet)
	•	Calculates quarterly changes, percentage changes, and total volumes for each stock ticker.
	•	Identifies and updates the summary table with the results.
	•	Highlights positive changes in green and negative changes in red using conditional formatting.
	•	Output: Summary table includes:
	•	Ticker symbol
	•	Quarterly change from opening to closing price
	•	Percentage change from opening to closing price
	•	Total stock volume
	•	Additional Functionality:
	•	Identifies stocks with the “Greatest % Increase”, “Greatest % Decrease”, and “Greatest Total Volume”.
	•	Outputs results in accordance with provided example image.

Usage

	1.	Excel Environment:
	•	Ensure Excel (Version 16.86) is installed with macros enabled.
	2.	Running the Script:
	•	Open your Excel workbook containing stock data.
	•	Access the VBA editor (Alt + F11), insert a new module, and paste the script.
	•	Run the Main subroutine to execute the analysis across all worksheets simultaneously.
	3.	Conditional Formatting:
	•	The script applies conditional formatting to highlight:
	•	Positive changes in green.
	•	Negative changes in red.


Dependencies

	•	Excel Version: Developed and tested on Microsoft Excel (Version 16.86).
	•	Resources Used: YouTube tutorial (https://www.youtube.com/watch?v=3OfVIsKy59c&ab_channel=EverydayVBA), Microsoft Community post (https://answers.microsoft.com/en-us/msoffice/forum/all/how-do-i-convert-a-date-formatted-as-text-as/a89d4a1d-cb39-459d-84d6-d473d90f4182), and assistance from ChatGPT.

License
This project is licensed under the MIT License.
