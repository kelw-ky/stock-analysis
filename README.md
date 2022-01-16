# Stock Analysis using VBA (Excel Macro)

## Overview of Project
 
### Purpose
Creating a VBA code on excel to help Steve analyze stocks over the past years in an even more effective way. We will be focusing on the Total Daily Volume and the return for twelve stocks during 2017 and 2018 to determine which stocks should be invested. 

## Results

### Refactor of the Code
The refactor of the original code is shown below; the reason for refactor is to make the code to run more efficiently by taking less step.

'1a) Create a ticker Index
 ticketIndex = 0

'1b) Create three output arrays
 Dim tickerVolumes(12) As Long
 Dim tickerStartingPrices(12) As Single
 Dim tickerEndingPrices(12) As Single
    
''2a) Create a for loop to initialize the tickerVolumes to zero.
For i = 0 To 11
  tickerVolumes(i) = 0
  tickerStartingPrices(i) = 0
  tickerEndingPrices(i) = 0    
Next i
            
''2b) Loop over all the rows in the spreadsheet.
For i = 2 To RowCount
    
  '3a) Increase volume for current ticker
  tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
  '3b) Check if the current row is the first row with the selected tickerIndex.
  If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
  End If
         
  '3c) check if the current row is the last row with the selected ticker
  If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
  End If
            
  '3d Increase the tickerIndex.
  If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
    tickerIndex = tickerIndex + 1
  End If
    
Next i
    
'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For i = 0 To 11
  Worksheets("All Stocks Analysis").Activate
  Cells(4 + i, 1).Value = tickers(i)
  Cells(4 + i, 2).Value = tickerVolumes(i)
  Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1        
Next i
    
### Analysis
From 2017 data, we can clearly see that all except for TERP had a postive return. DQ, SEDG, and ENPH had the highest return of 199.4%, 184.5%, and 129.5% respectively. The stocks with the highest Total Daily Volume were SPWR, FSLR and CSIQ with amounts of 782,187,000, 684,181,400, and 310,592,800 respectively.  

![2017_Chart](/Charts/2017_Chart.png)

However, the 2018 data are very different. From the 2018 chart below, we could see that most of the stocks has a negative return except for RUN and ENPH, which were the only two stocks with a positive return of 84% and 81.9% respectively. These two stocks also had the highest amount of Total Daily Volume of 502,757,101, and 607,473,500 respectively. 
![2018_Chart](/Charts/2018_Chart.png)

In the screenshots below, we can see the first execution times for the original code and the refactored code for 2017 and 2018 data. There is a huge difference in terms of efficiency as the execution times decreased from 1.1875 seconds to 0.2617188 seconds and 1.140625 seconds to 0.25 seconds for 2017 data and 2018 data respectively. 
![2017_ExecutionTimes](/Resources/VBA_Challenge_2017.png)
![2018_ExecutionTimes](/Resources/VBA_Challenge_2018.png)

## Summary

### Advantages and Disadvantages of Refactoring Code in general

Refactoring Codes makes running the code significantly more efficient as it will allow the code to execute quicker. It will also allow the code to be cleaner; refactoring will allow the fixer to detect duplicating code or longer methods. It will also allow debugging to happen and will help the code to run more smoothly. Disadvantages that could arise when one tries to refactor code is that it usually will take a longer as one will need to go through the whole code. Not everyone has the same amount knowledge in codes and might cause a bug as coding is like a fingerprint. On some occasions, it might cause more to refactor the code than just to start from scratch. 

### Advantages and Disadvantages of Original and Refactored VBA Script 

In this situation, none of the disadvantages arose since the person who created the original code was also the person refactoring the code. There were also no expenses associated to this project and the deadline given allowed the creator to have ample amount of time to refactor the code. It just allowed a higher efficiency, and a cleaner code to work and understand. 
