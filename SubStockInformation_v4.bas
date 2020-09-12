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
Dim HighestVolumeTicker, HighestPerChangeTicker, LowestPerChangeTicker As String
Dim BeginPrice, EndPrice, LastTrade As Integer
Dim SumOfVolume, HighestSumVolume As Long
Dim HighestPerChange, LowestPerChange As Double

'Begin with the loop through each sheet, print the column names

For Each ws In Sheets

    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
'Find the last row.  When on to using all sheets, need to mark this as ws.Cells(Rows.Count, 1).End(xlUp).Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set Counters to the first row of data
    i = 2
    j = 2
    
'Set formatting for percents
    ws.Columns("K").NumberFormat = "0.00%"

'Do loop which goes through each row in the sheet; originally I was going with a FOR loop here
'I got help from someone on stackoverflow suggesting i go with a Do While
    Do While i <= LastRow
        
        'start with setting the BeginPrice and SumOfVolume to the first values
        BeginPrice = ws.Cells(i, 3).value
        SumOfVolume = ws.Cells(i, 7).value
    
'Create a Do loop through the stocks
        Do While ws.Cells(i, 1).value = ws.Cells(i + 1, 1).value
            i = i + 1
           
'Sum the stock volume
            SumOfVolume = SumOfVolume + ws.Cells(i, 7).value
        Loop
 
'Set the EndPrice
        EndPrice = ws.Cells(i, 6)
 
'Enter in the Stock Ticker Symbol
        ws.Cells(j, 9) = ws.Cells(i, 1)
        
'Populate with stock price change, SumOfVolume and format change with green if
'positive and red if negative
        ws.Cells(j, 10) = EndPrice - BeginPrice
        
        If ws.Cells(j, 10) < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
        ElseIf Cells(j, 10) > 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
        End If
        
        ws.Cells(j, 12) = SumOfVolume
        
'Challenge option: find the largest amount of shares traded.
'        If HighestSumVolume <= SumOfVolume Then
'            HighestSumVolume = SumOfVolume
'            HighestVolumeTicker = ws.Cells(j, 9)
'        End If
        
        
'Populate with sotck percentage change. If statement here if trying to divide by zero
        If BeginPrice <> 0 Then
            ws.Cells(j, 11) = ws.Cells(j, 10) / BeginPrice
            
'Test for challenge options: Highest and lowest change in percentage
            If HighestPerChange <= ws.Cells(j, 10) / BeginPrice Then
                HighestPerChange = ws.Cells(j, 10) / BeginPrice
                HighestPerChangeTicker = ws.Cells(j, 9)
            End If
            If LowestPerChange >= ws.Cells(j, 10) / BeginPrice Then
                LowestPerChange = ws.Cells(j, 10) / BeginPrice
                LowestPerChangeTicker = ws.Cells(j, 9)
            End If
                        
        Else
            ws.Cells(j, 11) = "Not Applicable"
        End If
        
'increase the next interation
        i = i + 1
        j = j + 1
    Loop
    
    'MsgBox (BeginPrice & " " & EndPrice)

Next ws

'Select sheet to put in highest/lowest numbers
Sheets("A").Select

'Enter in the titles for the challenge option
Range("P1") = "Ticker"
Range("Q1") = "Value"
Range("O2") = "Greatest % Increase"
Range("O3") = "Greatest % Decrease"
Range("O4") = "Greatest Total Volume"

'Enter in the values for the challenge option
Range("P2") = HighestPerChangeTicker
Range("P3") = LowestPerChangeTicker
'Range("P4") = HighestVolumeTicker

Range("Q2") = HighestPerChange
Range("Q3") = LowestPerChange
'Range("Q4") = HighestVolume

End Sub

Sub TestLoopThroughSheet()
'Find the last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'For loop which goes through each row in the sheet
    For i = 2 To LastRow
        Range("H2") = LastRow
    Next i
        
End Sub
