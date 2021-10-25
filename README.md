# VBAModule

Project Overview
The purpose of this analysis is to refactor the module 2 solution code. The refactored code should loop through all of the stock data one time to collect information on each stock's total daily volume and the return. Additionally, we will determine the new analysis run time and evaluate how it compared to the run time of the subroutine before the code was refactored.

Results
Almost all of the stocks we examined in Steve's workbook had a positive percent return in 2017. The only stock that had a negative percent return in 2017 was "TERP". But in 2018, almost all of the stocks had a negative percent return. There were 2 stocks, "ENPH" and "RUN", that had positive returns in 2018. Overall, the majority of the stocks had better performance in 2017 than 2018.
 
After refactoring the code, the run times were reduced by about 4X. The new run times were 0.0703125 seconds for both year 2017 and 2018. The previous run times from the module were about 0.28 seconds for both year 2017 and 2018.
 
 
When we refactored the code, we created output arrays for each ticker's volume, starting price, and ending price. This improved our overall logic and allowed us to write the conditional statements more efficiently, utilizing a tickerIndex.

#--------------------------------------------------------------------------------------------------------------------------------
Code:
Sub Refactored()

    Dim startTime As Single
    Dim endTime As Single
    
'User inputs the year they want to analyze
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer
    
'Format the output sheet on All Stocks Analysis worksheet
    
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
'Create a header row for the summary table

    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
'Activate data worksheet
    Worksheets(yearValue).Activate
    
'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
'1a) Create a tickerIndex variable and set it equal to zero before iterating over all the rows.
    
    Dim tickerIndex As Integer
        tickerIndex = 0

  
'1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
'2a) Create a for loop to initialize the tickerVolumes to zero for all 12 elements in the array
        For i = 0 To 11
            tickerVolumes(i) = 0
        Next i
 
    '2b) Loop over all the rows in the spreadsheet.
        
        For j = 2 To RowCount
            
            '3a) Increase the volume for current ticker
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
 
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            End If
           
            '3c) Check if the current row is the last row with the selected ticker
            'If the next row's ticker doesn't match, increase the tickerIndex.
            '3d) Increase the tickerIndex
            
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                tickerIndex = tickerIndex + 1
            End If
            
        Next j
            
   
'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        
        For m = 0 To 11
            Worksheets("All Stocks Analysis").Activate
            
            Cells(4 + m, 1).Value = tickers(m)
            Cells(4 + m, 2).Value = tickerVolumes(m)
            Cells(4 + m, 3).Value = tickerEndingPrices(m) / tickerStartingPrices(m) - 1
    
        Next m

'Formatting the summary table headers
    
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
'Conditional formatting in the returns column
    
    dataRowStart = 4
    dataRowEnd = 15
    
    For k = dataRowStart To dataRowEnd
        If Cells(k, 3) > 0 Then
            Cells(k, 3).Interior.Color = vbGreen
        Else
            Cells(k, 3).Interior.Color = vbRed
        End If
    Next k
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
#------------------------------------------------------------------------------------------------------------
 
Summary
Refactoring code is an important part of the coding process. Usually the first draft of code can be streamlined and improved upon. There are many advantages to refactoring code; these advantages include improving the logic, organization, and readability of the code. Also, removing bugs and vulnerabilities in the code. On the other hand, there are also disadvantages to refactoring code. Some disadvantages include that refactoring code can be time consuming or increase the chance of mistakes as the code is worked over multiple times.
For this project, refactoring the code improved the organization, readability, and run time. The refactored code will also help Steve in the future, when he needs to analyze multiple stocks and needs the analysis to run quickly. The disadvantage in this case was refactoring the code was somewhat time consuming.

