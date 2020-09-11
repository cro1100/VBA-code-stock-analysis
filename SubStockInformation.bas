Attribute VB_Name = "Module1"
Sub StockInformation()

'For this project, we have multiple worksheets which have stock information.
'On each sheet we have a list of stocks with each row trading information on a particular
'day.  Our project is to summarize this data by looping through to find changes in each stock
'over the time frame.

'The data is a list of each stock's prices on a particular day, sorted by stock first, then
'by trading day.  In addition to the day's prices, you also have the volume of shares traded.

'There is a challenge option to find the stocks with the greatest changes
'in price, percent and total volume

'I will use 3 loops
'1.) A loop to go through all the sheets
'2.) A loop to go through the individual sheet, all the rows
'3.) A loop to go through each stock

'1.) Looping all sheets
'Will create the headers for each column and cycle through the entire project

'2.) Looping through each individual sheet
'Will go through all of the data on the sheet.  This will also allow us to address the challenge
'portion of the project.

'3.) Looping through the stocks
'Will go through each stock by identifying a change in ticker symbol, will sum the stock volume
'over each iteration, and will save the opening and closing prices.  Once this loop exits, the
'output will be created on the line under the previous stock.

'------------------------------------------------------------------

'Declare variables
Dim ws As Worksheet

'Begin with the loop through each sheet

For Each ws In Sheets

    ws.Range("I1") = "Ticker"
    ws.Range("J2") = "Yearly Change"
    ws.Range("K2") = "Percent Change"
    ws.Range("L2") = "Total Stock Volume"

Next ws

End Sub
