# VBA Challenge: Stock Market Analysis
## Overview
A VBScript was created to loop through all the stocks in the provided Multiple-year file. For each stock in a given year, its' ticker symbol, yearly change, percent change, and total stock volume were obtained and stored adjacent to the raw data. Yearly change was determined using the opening and closing prices for the given year, and the percent change was calculated based on the yearly change. Conditional formatting highlighted cells with positive yearly change green and negative yearly change red, leaving cells with no net change uncolored. 

Once yearly change, percent change, and total stock volume were obtained for each stock ticker in the given year, the ticker pertaining to greatest percent increase, greatest percent decrease, and greatest total volume was acquired. 

The VBScript encompasses both sections entirely before moving on to the next worksheet (the next calendar year) in the workbook. 
## Results
Attached are snipets of the first page of the multiple-year file after the VBScript has been run. From Columns A to G are the raw data, Columns I to L the calculated yearly change, percent change, and total volume per stock ticker, and Columns O to Q are the filtered tickers and their corresponding values. 
### 2018
![2018](https://user-images.githubusercontent.com/100882943/161402680-05d9fb21-4670-4aef-a60f-1277889a6f45.JPG)

### 2019
![2019](https://user-images.githubusercontent.com/100882943/161402686-5364b017-6535-4a69-b62d-713fe3c3f0ba.JPG)

### 2020
![2020](https://user-images.githubusercontent.com/100882943/161402689-f8887467-a322-4776-93a2-35cd8edfe98a.JPG)
