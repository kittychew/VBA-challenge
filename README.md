# VBA Stock Market Analysis Challenge

## Project Overview
This project demonstrates the use of **VBA scripting** to automate the analysis of stock market data for multiple quarters. The script calculates key financial metrics, including:
- **Quarterly Change**: The difference between the opening and closing price for each quarter.
- **Percentage Change**: The percentage change from the opening to the closing price.
- **Total Stock Volume**: The total number of stocks traded during the quarter.

Additionally, the script identifies the **stock with the greatest percentage increase**, **greatest percentage decrease**, and **greatest total volume** for each quarter.

## Technologies Used
- **VBA (Visual Basic for Applications)**: For automating the analysis in Excel.
- **Microsoft Excel**: For managing and manipulating the stock data.

## Key Features
- **Data Retrieval**: The script loops through each row of stock data, capturing the ticker symbol, open price, close price, and stock volume.
- **Data Processing**: 
  - Quarterly change in stock price is calculated.
  - Percentage change is computed based on the opening and closing prices.
  - Total stock volume is summed for each stock.
- **Conditional Formatting**: 
  - Positive changes are highlighted in green, while negative changes are highlighted in red.
- **Key Insights**: 
  - The script identifies the stock with the greatest percentage increase, greatest percentage decrease, and the greatest total volume.
- **Multi-Sheet Analysis**: The script is capable of processing data across multiple worksheets, each representing a different quarter.

## How to Run the Code
1. Download the **alphabetical_testing.xlsx** file and save it to your local system.
2. Open **Excel** and press `Alt + F11` to open the VBA editor.
3. Insert a new module by selecting `Insert → Module`, and paste the provided VBA script.
4. Run the macro by pressing `Alt + F8` and selecting the macro name to process the data.
   
The code will loop through the worksheets in the workbook and output the calculated metrics in each sheet.

## Files
- **vba_stock_analysis_script.bas**: Contains the VBA code to process the stock data.
- **alphabetical_testing.xlsx**: A sample dataset used to test the script.
- **README.md**: This file containing the project details.

## Calculations and Outputs
- **Quarterly Change**: Difference between the opening and closing price for each quarter.
- **Percentage Change**: Calculated as `(Closing Price - Opening Price) / Opening Price * 100`.
- **Total Stock Volume**: Sum of the trading volume for each quarter.
- **Greatest Increase/Decrease/Volume**: Identifies the stock with the highest change and volume.

  [![2-VBA-screenshot.jpg](https://i.postimg.cc/V6LVdNKv/2-VBA-screenshot.jpg)](https://postimg.cc/HJR2NTD1)
  

## Results and Insights
The script outputs the following:
1. **Ticker symbol**, **quarterly change**, **percentage change**, and **total volume** for each stock in each quarter.
2. **Greatest percentage increase**, **greatest percentage decrease**, and **greatest total volume** across the entire dataset.

   [![2-VBA-returns.jpg](https://i.postimg.cc/L4q38fmw/2-VBA-returns.jpg)](https://postimg.cc/z3rgdybF)

## Future Improvements
- The code can be enhanced to support larger datasets and more complex analyses.
- Additional calculations, such as moving averages or price trends, can be added for deeper insights.


