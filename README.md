# stock-analysis

## Overview of Project
**Background**
This project is to help Steve who wants to diversify his parents funds in investing in the _right_ and _profitable_ stocks. He wants to analyze many green stocks including DAQO stocks for his parents. He created an Excel file for the analysis but I used **VBA** to automate the task as it interacts with Excel in a way that would enable Steve to just **push a button** and analyze many stock.

### Purpose
**Goal**
The goal of this project is to find the **total daily volume and yearly return** for each stock. The purpose is to create an **efficient code** to analyze multiple stocks. The first code performed the analysis of the 12 stocks but I refactored the code to make it more efficient to work with Steves' data. The goal of the project is to test if refactoring improves the performance of the code.


**Approach/Method**
Using **Visual Basic Application (VBA)** in Excel to find the stocks **total daily volume and annual return** as shown below
[VBA_Challenge.xlsm](https://github.com/FUNMIIB/stock-analysis/blob/main/VBA_Challenge.xlsm)


### Results
I made the code more efficient by creating 4 different arrays: tickers, tickerVolume, tickerStartingPrices, and tickerEndingPrices as follows:

Dim tickers(12) As String
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

I used **tickerIndex** to match the tickers array with the 3 arrays. The refactored code performed more efficiently as shown in the image below.
![VBA_Challenge_2017.png](https://github.com/FUNMIIB/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png) 
![VBA_Challenge_2018.png](https://github.com/FUNMIIB/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)


##  Summary
**Advantages of refactoring**
Refactoring the code improved the performance of the code so task is executed faster
- This would allow the end user to analyze different stock data in short period of time.

**Disadvantages of refactoring**
- Code needs to be re-run to test functionality of the script.
- Time consuming.


