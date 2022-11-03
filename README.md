# Week2 - VBA-Homework: The VBA of Wall Street
# Author - Jyothi Palle

Objective of the work is to output the below 
1. Ticker sysmbol
2. Yearly change fropm opening price at the begining of a given year to the closing price at the end of that year
3. The percent change from opening price at the beginning of a given year to the closing price at the end of that year
4. The total stock volume of the stock.
5. Bonus - Find the "Greatest % increase", "Greatest % decrease", and "Greatest total volume" for each year

Solution -

1. Initialise the required variables 
2. Find the number of rows in the sheet
3. Write the condition to check if its the opening row for the year then store opening value, loop through all the rows for the particular row adding stock value0 and at the last row capture closing value
4. Calcuate yearly change as closing - opening
5. Apply confitonal format for yearly change if the change >0 then green else red
5. Percent change as yearly change / opening
6. Output the values to output table defined in the same sheet
7. Repeat the same process for all tickers within the same sheet
8. Repeat steps 2 to 7 for all the sheets within workbook

Uploaded the below documents to the VBA-Challenge repository
1. Readme file
2. VBA-Challenege Script file
3. Output Screenshot for 2018 sheet
4. Output Screenshot for 2019 sheet
5. Output Screenshot for 2020 sheet


