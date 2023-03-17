# VBA_Challange
Module_2_Challange
VBA-challenge 

Stock Market Analysis

Background 
    The analysis used two data sets. One is called Alphabetical Testing. This data set includes (7) worksheets labeled A-P. The other data set is called Multiple year stock data. It is a larger file and contains (3) worksheets labeled 2018,2019, and 2020. Due to the size of the data set Multiple year stock data , the Alphabetical testing data set is used for script testing. The VBA script is found in the file Scripts.VBS.

Solution
    The script loops through all worksheets in the Multiple year stocks data set and the following information displayed.

Solution 1 Ticker Symbol
    The script will sort for the distinct ticker symbol from column “A” and place it in column “J” with a column header "Ticker”.

Solution 2 Yearly Change
    The script will execute Yearly change from opening price at the beginning of a given year to the closing price at the end of that year and put the value in the "K" column. The code added conditional formatting that highlighted positive change in green and negative change in red.

Solution 3 Precent Change
    The script also performs Percent change from opening price at the beginning of a given year to the closing price at the end of that year and put the value in the "L" column.

Solution 4 Total stock Volume
    The script adds up all the total stock volume of a given ticker name and inserts this data on in the “M” column.
Solution 5 Greatest
	The script also provides the "Greatest % increase", "Greatest % decrease" and "Greatest total volume" of each year. 


