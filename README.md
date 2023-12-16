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
