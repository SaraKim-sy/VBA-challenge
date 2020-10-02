# VBA CHALLENGE

## Table of contents
  * [About this project](#about-this-project)
  * [Getting started](#getting-started)
  * [Analysis Result](#analysis-result)



## <a name="about-this-project"></a> About this project
This project is to develop a VBA script to analyze stock market data. The raw data, which includes 3 years of stock market data in each tab, is in a Microsoft Excel file.
Using the VBA script will allow you to analyze each year and summarize meaningful analysis in new columns in each worksheet all at once.


## <a name="getting-started"></a> Getting started

### <a name="step-one"></a> Step 1
Download the files required.
The VBA script, 'Stock_Market_Analysis_Script.bas,' was uploaded in this repository. 

### <a name="step-two"></a> Step 2
Open the excel workbook which has raw data. In the Developer tab, click Visual Basic. 

### <a name="step-three"></a> Step 3
Import the VBA script 'Stock_Market_Analysis_Script.bas' file (File > Import File).
Then, press F5 button on the keyboard or click the play button to run the script.

* The script will loop through all the stocks for one year and output the following information:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    * It also has conditional formatting that will highlight positive change in green and negative change in red.
    
  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* The script will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". 


## <a name="analysis-result"></a> Analysis Result
After running the script on the Workbook, the result should look as follows:

### 2014 Analysis Data
![Image of 2014 Analysis](https://github.com/SaraKim-sy/VBA-challenge/blob/master/2014%20Data.png?raw=true)

### 2015 Analysis Data
![Image of 2015 Analysis](https://github.com/SaraKim-sy/VBA-challenge/blob/master/2015%20Data.png?raw=true)

### 2016 Analysis Data
![Image of 2016 Analysis](https://github.com/SaraKim-sy/VBA-challenge/blob/master/2016%20Data.png?raw=true)
