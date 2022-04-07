# Green Stock Analysis Using VBA
## Overview

### Purpose

Steve wants to help his parents invest in stock in a company called Green, for this, he needs a fast and reliable way to do an analysis over the years and be able to determine if DQ is a good option, which his parents have previously chosen.

### Background

For this requirement, we used VBA going through the learning process on how to use loops and calculations, intetrgate interaction elements with the user and  conditional formats to accomplish the target

On a second stage, we improved the code to make it faster by using variable arrays that helped the code not go back and forth between different sheets, just holding values and applying them to reduced loops.

## Results

### Analysis

<img align="left" width="300" height="300" src="https://github.com/adolfoxitlan/stock-analysis/blob/main/Resources/Resultados2017.jpg"> After finishing the first process, we can determine that the stocks, in general, had gains in 2017.

<br clear="left"/>
<br clear="left"/>

<img align="left" width="300" height="300" src="https://github.com/adolfoxitlan/stock-analysis/blob/main/Resources/Resultados2018.jpg"> But in 2018 there were significant changes, which show that DQ is quite unstable, but also we can see that RUN and ENPH are good investing candidates.

Another detail that we can see from this analysis is that the higher the volume, the higher the losses, so definitely, DQ is no go.
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
### Code

First of all I would like to show a how it looks the user interface.

<img src="https://github.com/adolfoxitlan/stock-analysis/blob/main/Resources/Green-Stock.gif">

Even though both codes give the same result, I refactored the code to eliminate nested loops; it was a simple change that took me long to assimilate. In addition, adding arrays before entering the loop reduces performance utilization considerably.. 

Nested For

    '4) Loop through tickers
      For i = 0 To 11
      ticker = tickers(i)
      totalVolume = 0

    '5) loop through rows in the data
        Sheets(yearValue).Activate
        For j = 2 To RowCount
           '5a) Get total volume for current ticker
                If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
                
                End If
            
           '5b) get starting price for current ticker
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
    
                End If
           
           '5c) get ending price for current ticker
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
                
                End If
           
        Next j
        
Using Arrays to eliminate j Loop (For)

    '1a) Create a ticker Index
        Dim tickerIndex As Integer
            tickerIndex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
             
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(tickerIndex) = 0
            tickerStartingPrices(tickerIndex) = 0
            tickerEndingPrices(tickerIndex) = 0
        
    Next i

The refactored code had a remarkable impact on performance, it was not so in the number of lines; it remained similar.

The execution time reduction is considerable:
- 2017 from 1.023438 to 0.125 seconds.
- 2018 from 0.1328125.1.023438 to 0.125 seconds.

Quite Impressive!!!!

## Summary

### Advantages and Disadvantages

There are many advantages of refactoring to improve code.
- Is clean to read.
- Easy to understand.
- More efficient.

Disadvantages in general:
- Time consuming.
- You need to be careful and re-do to validate that your code is in good shape and you didn't add an error by mistake.
- Sometimes you mess up something that was already working.

In Steve's case, in particular, carrying out the activity had benefits that I consider important since it gives:
- Flexibility.
- Better performance to add other types of stocks and more data, which can degrade the performance in future analysis.
- In both cases (Orignal and Refactored Scripts), the result are the same, but the performance of the refactored is enormous.  

The only disadvantage I see is that the programmer has to spend more time deciphering how to enhance it.

