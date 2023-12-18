# VBA Challenge

![stockmarket](https://user-images.githubusercontent.com/112406455/194172615-177dcc2b-4c3b-4911-a681-a46bcc4ce3e9.jpg)

## Background
In the ever-evolving landscape of the financial markets, the ability to swiftly analyze vast amounts of data can provide a significant edge. This project leverages the power of VBA (Visual Basic for Applications) scripting to dissect and interpret generated stock market data with efficiency and precision. Our VBA script is designed to automate the analysis process, rapidly parsing through datasets to extract key financial metrics and performance indicators. By integrating these scripts into our workflow, we aim to uncover actionable insights from the stock market data, enabling more informed investment decisions and fostering a deeper understanding of market dynamics.
## Objective
The primary goal of this project is to develop a robust VBA script capable of iterating across an entire year's worth of stock data to compute and output critical performance metrics. The script will systematically loop through each stock and gather the following information:

* **Ticker Symbol**: Uniquely identifying each stock.
* **Yearly Change**: The difference between the opening price at the year's commencement and the closing price at the year's end.
* **Percentage Change**: This represents the yearly change as a percentage of the opening price.
* **Total Stock Volume**: The sum of the traded volume for the stock throughout the year.

Additionally, the script will be equipped to identify key performers in the dataset by reporting stocks with the `Greatest % Increase`, `Greatest % Decrease`, and `Greatest Total Volume`. This will provide a snapshot of the most significant movements within the market over the course of the year.

To ensure a comprehensive analysis, the script will be designed to run across multiple worksheets, which will allow for an analysis of data from different years in a single execution. This approach aims to provide a more holistic view of stock performance trends over time, enhancing the decision-making process for investors.
## Data
The dataset utilized in this project encompasses a detailed record of stock market transactions over several years, provided by edX Boot Camps LLC strictly for educational purposes. It is structured into an Excel workbook titled `Multiple_year_stock_data.xlsx`, with each worksheet corresponding to a full year's trading data for a variety of stocks. Below is an outline of the data columns present in each worksheet:
| Column  | Description                                              			       |
|---------|------------------------------------------------------------------------------------|
| Ticker  | The unique stock identifier                                                        |
| Date    | The trading date                                                                   |
| Open    | The price of the stock at the market open                                          |
| High    | The highest price of the stock during the trading day                              |
| Low     | The lowest price of the stock during the trading day                               |
| Close   | The price of the stock at the market close                                         |
| Volume  | The number of shares traded during the day                                         |
## Implementation
### Creating the Subroutine
The first step in building our VBA script is to create a subroutine. This is where all of our code for analyzing the stock data will reside. 
```VBA
Sub Stock_Data_Analysis()
```
### Declaring Variables
After establishing our subroutine, the next step is to declare the variables we will use throughout our script. This is crucial for structuring our data and ensuring smooth execution of the script
```VBA
'Define all variables'------------------------------------------------------------------------------------------------------'(Ticker Symbol, Open Price, Closing Price, Percentage Change, Total Stock Volume,Yearly Change, Greatest Total Volume)Dim ticker As StringDim open_price As DoubleDim closing_price As DoubleDim pc As DoubleDim tsv As DoubleDim yc As DoubleDim gtv As Double'Other variablesDim PreviousStockPrice As LongDim table_summary_row As LongDim greatest_increase As DoubleDim greatest_decrease As Double'Declare Worksheet as "ws" and Loop through each worksheetDim ws As WorksheetFor Each ws In Worksheets
```
In this section:
* We declare various types of variables like `String`, `Double`, and `Long`. Each type serves a specific purpose, such as handling text, decimal numbers, or large integers.
* `ticker`, `open_price`, `closing_price`, etc., are used to store and manipulate stock data.
* `PreviousStockPrice` and `table_summary_row` are used for tracking and managing data during the script execution.
* `ws` is declared as a Worksheet object, allowing us to iterate over each worksheet in the Excel workbook.

Declaring variables at the beginning of our script enhances readability and maintenance, making the script easier to understand and modify.
### Labeling Column Headers and Tables
Once the variables are declared, the next step in our script is to label the column headers and tables. This is important for organizing the output of our analysis in a readable and accessible format.
```VBA
'Label Column Headers and Tablesws.Range("P1").Value = "Ticker"ws.Range("Q1").Value = "Value"ws.Range("O2").Value = "Greatest % Increase"ws.Range("O3").Value = "Greatest % Decrease"ws.Range("O4").Value = "Greatest Total Volume"ws.Range("I1").Value = "Ticker"ws.Range("J1").Value = "Yearly Change"ws.Range("K1").Value = "Percent Change"ws.Range("L1").Value = "Total Stock Volume"
```
In this code block:

* We are using the `Range` property of the worksheet object `ws` to access specific cells.
* Each cell is assigned a value that serves as a header for different data columns and tables.
* For example, `ws.Range("P1").Value = "Ticker"` sets the value of cell P1 to "Ticker", which will serve as the header for the ticker symbols in our analysis.

By labeling the columns and tables clearly, we ensure that the output of our script is easy to understand and interpret, facilitating better analysis of the stock data.
### Analyzing Each Stock: Yearly Change, Percent Change, and Total Stock Volume
After setting up our worksheet, the next key task in our script is to analyze each stock to find the Yearly Change, Percent Change, and Total Stock Volume.
```VBA
'For each stock find the Yearly Change, Percent Change, and Total Stock Volume'Assign values to variables for loop to starttsv = 0table_summary_row = 2PreviousStockPrice = 2'Set value of the last row for column AEndRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row'Loop through first rows for stock infoFor i = 2 To EndRowA        'Find the value of the Total Stock Volume        tsv = tsv + ws.Cells(i, 7).Value                'Execute to record for change in stock ticker in the summary table with ticker name and tsv and reset tsv back to zero        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then                        ticker = ws.Cells(i, 1).Value                                ws.Range("I" & table_summary_row).Value = ticker                                ws.Range("L" & table_summary_row).Value = tsv                                tsv = 0                                'Next find the year open and close price, yearly change, and percentage change                open_price = ws.Range("C" & PreviousStockPrice)                                close_price = ws.Range("F" & i)                                yc = close_price - open_price                                ws.Range("J" & table_summary_row).Value = yc                                'Start another if statement to determine percent change                If open_price = 0 Then                                    pc = 0                                    Else                    open_pice = ws.Range("C" & PreviousStockPrice)                                    pc = yc / open_price                                    End If                                'Place value of percentage change in summary table using the % format                ws.Range("K" & table_summary_row).Value = pc                                ws.Range("K" & table_summary_row).NumberFormat = "0.00%"                                'Make another if statement for conditional formating the cells of yearly change (green=positive/red=negative)                If ws.Range("J" & table_summary_row).Value >= 0 Then                    ws.Range("J" & table_summary_row).Interior.ColorIndex = 4                                    Else                    ws.Range("J" & table_summary_row).Interior.ColorIndex = 3                                    End If                                'Initiate task to go to next row for summary table and previous stock price                table_summary_row = table_summary_row + 1                                PreviousStockPrice = i + 1                            End If                        Next i
```
This code block accomplishes the following:

* Iterates through each row of the stock data.
* Accumulates the Total Stock Volume (`tsv`) for each stock.
* Identifies when the stock ticker changes and then calculates the Yearly Change (`yc`) and Percent Change (`pc`).
* Records these values in a summary table.
* Applies conditional formatting to the Yearly Change values for visual representation of positive (green) or negative (red) changes.

By looping through each stock, this script segment effectively computes and organizes key financial metrics, making it easier to analyze the stock market data.
### Identifying Key Performers: Greatest % Increase, Greatest % Decrease, and Total Volume
After analyzing each stock, the script then focuses on identifying the stocks with the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume. This is achieved through another loop in our script.
```VBA
'Make another loop for Greatest % Increase, Greatest % Decrease, and Greatest Total Volume'Assign values to variables for loop to startgtv = 0greatest_increase = 0greatest_decrease = 0'Set value of the last row for column KEndRowK = ws.Cells(Rows.Count, 11).End(xlUp).RowFor i = 2 To EndRowK    'First determine the Greatest Total Volume    If ws.Range("L" & i).Value > gtv Then       gtv = ws.Range("L" & i).Value       ws.Range("Q4").Value = gtv       ws.Range("P4").Value = ws.Range("I" & i).Value           End If        'Next determine Greatest % Increase    If ws.Range("K" & i).Value > greatest_increase Then        greatest_increase = ws.Range("K" & i).Value        ws.Range("Q2").Value = greatest_increase        ws.Range("P2").Value = ws.Range("I" & i).Value            End If        'Last determine Greatest % Decrease    If ws.Range("K" & i).Value < greatest_decrease Then        greatest_decrease = ws.Range("K" & i).Value        ws.Range("Q3").Value = greatest_decrease        ws.Range("P3").Value = ws.Range("I" & i).Value            End If        'Change format to "%" for Greatest % Increase and Decrease    ws.Range("Q2").NumberFormat = "0.00%"        ws.Range("Q3").NumberFormat = "0.00%"    Next i
```
This section of the script:

* Loops through the Percent Change column to find the stocks with the greatest increase and decrease in value.
* Records the Ticker Symbol and the corresponding values for these key performers.
* Formats the Percent Change values to display them as percentages.
* Identifies the stock with the Greatest Total Volume and records its details.

By highlighting these key performers, the script provides valuable insights into which stocks had the most significant positive and negative changes, as well as which had the highest trading volume.


