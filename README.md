# VBA of Wall Street

## Overview of Project
The puropse is to take existing VBA code for analyzing stock market price and volume data and refactor, to make the code more efficient, in terms of running faster and having more concise code.

### Purpose
By making the coding more efficient, we create the opportunity to more easily add greater functionality to the existing code.

## Analysis and Challenges
The initial subroutine was limited in funcionality to the particular data set given.  This is because we hard-coded the sheet we'd be working on.  By utilizing the InputBox() function, the user can manually input the name of the sheet they wish to use.
```
yearValue = InputBox("What year would you like to run the analysis on?")
```
The biggest issue with the original code was the excessive number of loops.  I used one loop with nested conditional statements to make the code more efficient:
```
 For i = 2 To RowCount
    
	tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
       	 If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            
		tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
           
		tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            	tickerIndex = tickerIndex + 1
            
        End If
    
   Next i
```
I then outputed the desired values to the All Stocks Analysis worksheet using the arrays in a loop:
```
 For i = 0 To 11
        
    	Worksheets("All Stocks Analysis").Activate
        
     	Cells(dataRowStart + i, 1).Value = tickers(i)
    	Cells(dataRowStart + i, 2).Value = tickerVolumes(i)
     	Cells(dataRowStart + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
 Next i
```

### Analysis of Performance
Below are screenshots highlighting performance.

**BEFORE:**

![2017 Before Refactor](https://github.com/maddenc33/stock-analysis/blob/main/Resources/yearValueAnalysis%202017.png?raw=true)

![2018 Before Refactor](https://github.com/maddenc33/stock-analysis/blob/main/Resources/yearValueAnalysis%202018.png?raw=true)

**AFTER:**

![2017 After Refactor](https://github.com/maddenc33/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png?raw=true)

![2018 After Refactor](https://github.com/maddenc33/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png?raw=true)

## Results
As can be seen in the screenshots of the pop-up messages, performance was dramatically improved once the refactoring of the code was completed.
Using the 2017 worksheet as an example, we can observe a change in run-time from 0.6640625 seconds to 0.1132813 seconds.  Our code is now roughly 6x faster!
By using conditional loops more efficiently, one can essentially make less work for the computer to do in terms of iterating the code and therefore finish faster.  Efficiency will become increasingly more important as we use larger and larger data sets.

