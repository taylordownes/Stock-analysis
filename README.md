# Stock-analysis
## Overview of Project
  The purpose of this project was to generate, edit, and refactor VBA script for the Green Stocks dataset that helped us determine how well stocks performed in the past. Our goal was to generate a macro that looped through the data set by the click of a button to compare the stock tickers to their totaly daily volume and return rates. To acheive this, we used a combination of loops, If statements, static and conditional formatting, and more to generate a clean report for analysis.  
### Analysis and Challenges
Our project was based on helping "Steve" create an analysis for his parents to determine how well Daquo Energy Corporation (DQ) performed in the past before they decided to invest their money. Aside from helping Steve with an analysis solely for DQ, we also created an analysis that compared other green energy corporations to DQ to get a better understanding of how well DQ stock performed. 
The series of procedures were as followed:
- Used the green stocks data to generate a report for DQ that calculated the yearly return for DQ in 2018 on the DQ Analysis tab in the Green_Stocks.xlsm file.
- Created an analysis for all stocks to see if different stocks performed better than DQ on the All Stocks Analysis tab in the Green_Stocks.xlsm file.
- Created buttons to run and clear the code we wrote for the analysis.

Some of the challenges in this analysis were degbugging the macros and creating a clean readable VBA script. The debugging process required percise analysis of the code to determine where the errors were generating. Creating a clean and readable VBA script became much more clear in the refactoring stage of the analysis.
- **Refactored the code in the VBA_Challenge.xlsm file to make the code more readable and efficient.**
## Results
In terms of our refactoring reults, we started by copying the code for creating header rows, intializing arrays for tickers, activating the worksheet, and getting the number of rows to loop over. The refactoring steps and code (starting at 1a) were as follows:


    Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    Dim TickerIndex As Single
    TickerIndex = 0
    
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
    
    
    '1a) Create a ticker Index
        
        TickerIndex = 0
    
    '1b) Create three output arrays
    
        Dim TickerVolumes(12) As Long
        Dim TickerStartingPrices(12) As Single
        Dim TickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
        For i = 0 To 11
        TickerVolumes(i) = 0
    
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    
        For i = 2 To RowCount

    
        '3a) Increase volume for current ticker
        
            TickerVolumes(TickerIndex) = TickerVolumes(TickerIndex) + Cells(i, 8).Value
            
            
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(i - 1, 1).Value <> tickers(TickerIndex) Then
            TickerStartingPrices(TickerIndex) = Cells(i, 6).Value
            
            
            End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        
            If Cells(i + 1, 1).Value <> tickers(TickerIndex) Then
            TickerEndingPrices(TickerIndex) = Cells(i, 6).Value
            
            
            '3d Increase the tickerIndex.
            
            TickerIndex = TickerIndex + 1
            
            'End If
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        
        Worksheets("All Stocks Analysis").Activate
        
        For i = 0 To 11
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = TickerVolumes(i)
        Cells(4 + i, 3).Value = TickerEndingPrices(i) / TickerStartingPrices(i) - 1
        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    range("A3:C3").Font.FontStyle = "Bold"
    range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    range("B4:B15").NumberFormat = "#,##0"
    range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub


#### 2017 Results
<img src="https://user-images.githubusercontent.com/84201614/122683271-c3477480-d1c3-11eb-94a4-c8fa87a0084c.png" width="325" height="275">

### 2018 Results
<img src="https://user-images.githubusercontent.com/84201614/122683437-c858f380-d1c4-11eb-9ffa-c9e72e653eea.png" width="325" height="275">

As we can see, DQ performed extremely well compared to the rest of the stocks in 2017 with a return rate of 199.4%, however, DQ experienced a large dip in 2018 with the lowest return rate of -62.6%. Steve's parents should be careful about investing all their money in Daquo energy based on these results. It would also be ebenficial to collect more data to see if there are additional trends to return on investment for DQ.

## Summary

### Advantages and Disadvantages of Refactoring Code
Advantages
- Makes code easier to understand
- Helps with the debugging process
- Can make code run faster

Disadvantages
- Can be time consuming
- Can create more debugging errors

### How do the advantages and disadvantages apply to the original VBA script?

The refactored code was easier to follow especially with the comments showing what each step in the code was doing. It was also easier to debug based on the orderly steps. The macro run time in the refactored code was about 1 second faster than the original script with the run times below:

<img src="https://user-images.githubusercontent.com/84201614/122684395-c7c35b80-d1ca-11eb-8984-a7d177731a83.png" width="415" height="225">

<img src="https://user-images.githubusercontent.com/84201614/122684468-27216b80-d1cb-11eb-92ba-49c21c9c364d.png" width="415" height="225">

Although the refactoring process was time consuming and created additional errors at first, the refactoring process made the code more efficient overall.
