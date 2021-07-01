# VBA Green Stock Analysis

## Overview of Project
The goal of this project was to assist a hypothetical client/friend named
Steve, who recently graduated with a finance degree.  His parents wanted 
to invest in alterntive energy production, but they hadn't done much research.
Steve intended to analyze various stocks to discover with of the alternative
energy companies would be the best investment for his parents.  A VBA macro was
created to make this search easier and more efficient for Steve to work with so 
all Steve had to do was press a button and the total volumes and yearly returns 
would pop up for each stock.  Although this original script worked well, if Steve 
wanted to process a lot more data, it could be time consuming.  In order to make
the process more efficient, the code was refactored.

## Results

### Analysis of Green Stock Performances in 2017 and 2018
<img src = "/resources/volume_chart.png" alt = "Total Volume in 2017 and 2018" width = "49%"> <img src = "/resources/returns2_chart.png" alt = "Returns in 2017 and 2018" width = "49%"> 

As you can see from the charts above, returns are negative for all stocks except
ENPH and RUN in 2018 whereas they were all positive in 2017 except for TERP which
only had a -7.2% return. Although the returns in 2018 were predominantly negative,
the total volume increased for DQ, ENPH, HASI, RUN, SEDGE, TERP and VSLR. However, 
of those 7, the volume for ENPH increased by almost 200% and the volume of RUN 
increased by almost 100%.  Although DQ increased by 200%, it's starting total volume was 
only $37 million, while ENPH started with $222 million and RUN started with around $268 million.
Also, DQ had a drop of 63% in returns in 2018.

<p align ="center">
  <img src = "/resources/green_stocks_2017.png" alt = "Stock Returns and Total Volumes in 2017" width = "35%"> <img src = "/resources/green_stocks_2018.png" alt = "Stock Returns and Total Volumes in 2018" width = "35%" > 
</p>

This data demonstrates that ENPH and RUN were the only stocks that had positive returns, which were 82% and 84% respectively. 
Considering this, in conjunction with their increase in total volume by large margins, ENPH and RUN
seem to be the only stocks worth investing in.  

### Analysis of VBA Refactoring
The start of both scripts were the same. Some variables were initialized, the input box for the  year was created, formatting for the table itself was created and a ticker array was initialized and filled with the stock tickers. Then the number of rows in the data table was counted.  The next part where it cycles throught the data to gather the desired value was where the original had efficiency issues.

In the original, there was a nested loop with the outer loop cycled through the tickers output the data to the table once it was gathered by the inner loop.  The inner loop cycles through *every* row of data  for each ticker. The loop was written as follows:
```     
    '4) Loop through Tickers
     For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        '5) Loop through rows in the data
        Worksheets(YearValue).Activate
        For j = 2 To RowCount
        
            '5a) get total volume for current ticker
            If Cells(j, 1).Value = ticker Then
            
                totalVolume = totalVolume + Cells(j, 8).Value
            
            End If
            
            '5b) get starting price for current ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                
                startingPrice = Cells(j, 6).Value
            
            End If
            
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
            
                closingPrice = Cells(j, 6).Value
            
            End If
            '5c)get closing price for current ticker
            
        Next j
        
        '6) Output data forcurrent ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = ticker
        Cells(i + 4, 2).Value = totalVolume
        Cells(i + 4, 3).Value = (closingPrice / startingPrice) - 1
        
    Next i
```
Although this technically works, due to the fact that the inner loop cycles through the whole data set and the outer loop cycles through the 12 tickers, the script cycles through all the data 12 times. In order to track the time it took to generate this code, a timer was included.  Using the original code, the time for both years generated the following results:
<p align = "center"> 
  <img src = "/resources/original_timer_2017.png" alt = "2017 Original Script Time" width = "30%">&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<img src = "/resources/original_timer_2018.png" alt = "2018 Original Script Time" width = "30%">
</p> 
As you can see, this is a substantial amount of time considering it's only cycling through 3,013 rows of data with 8 columns. When the hypothetical client, Steve, had to work with larger data sets, the time could become much more of a nuisance. With this in mind,the code was refactored so that the loop ony cycled once and gathered the data as it went.  The first step included was a ticker index with a value of 0.  Then, the volume, starting price and ending price that were regular variables in the original were turned into arrays with 12 indices each. The last preparation for the loop was to fill the total volume array with 0 as a starting value.  These steps are displayed here:

```
   '1a) Create a ticker Index
    
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolume(12) As Long
    Dim tickerStartingPrice(12) As Single
    Dim tickerEndingPrice(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolume(i) = 0
    Next
```

Throughout the loop, the tickerIndex will be used as the index for each of the three output arrays. The loop itself cycles through each row of the data set and adds the volume to the current ticker.  The next If block checks to see if its the first row of the ticker so it can add the starting price to the starting price array.  The next block checks to see if its the last row of the current ticker to add the ending price to the ending price array *and* if so, it will increment the tickerIndex variable by 1. By the end, each of the three arrays created above will house the data for each ticker.

```
      '2b) Loop over all the rows in the spreadsheet.
       For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
         tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + Cells(i, 8).Value
        
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
         If Cells(i, 1) = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then
        
            tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If the next row‚Äôs ticker doesn‚Äôt match, increase the tickerIndex.
         If Cells(i, 1) = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
            
            tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
            
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
          End If
    
    Next i
```
In the original script, the data was output within the outer loop.  In this case, there is a tickers array with the 12 stock tickers, the total volume array, the starting price array and the ending price array.  The last step wast to write a simple loop to output the tickers, total volume and, using the starting and ending price, a percentage for the return values.

```
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1) = tickers(i)
        Cells(i + 4, 2) = tickerVolume(i)
        Cells(i + 4, 3) = (tickerEndingPrice(i) / tickerStartingPrice(i)) - 1
   
    Next i
```
The rest of the script is simply for matting but the importance of the simplification of this loops becomes evident in the time it takes to process. The following are images of the processing time of this script for 2017 and 2018.

<p align = "center"> 
  <img src = "/resources/refactored_timer_2017.png" alt = "2017 Refactored Script Time" width = "30%">&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<img src = "/resources/refactored_timer_2018.png" alt = "2018 Refactored Script Time" width = "30%">
</p> 

Although, the times shows throughout are highly dependent on the individual computer's processing speed, the difference between the times is extremely significant. The original script took about 0.773 seconds longer than the refactored script for 2017, and 0.535 second longer for 2018. On a larger scale, this could make a huge difference. 

## Summary

### What are the advantages and disadvantages of refactoring code in general?
Generally, refactoring code, or at least seeing if you can, is extremely important. The main goal of programmers is to find the structure code as efficiently as possible.  Ultimately, the codes will run faster and save time with refactoring.  I could imagine it's a good idea to have someone else look at your code before refactoring to see if different eyes will see what you haven't.  In this case, structuring and commenting code clearly is extremely important.  This could be the cause of one significant disadvantage (and the cause of a lot of issues in general) for refactoring.  Even if one is refactoring there own code, if it's disorganized and unclear, it will cause issues.

### What are the advantages and disadvantages of the original and refactored VBA script?
Clearly, refactoring the script was very advantageous in this case because it made the process a lot faster.  Hopwever, in this case the goal was to make the code more efficient for any stock data fed through it.  In the refactored code, there are still only the tickers for this data set. A next step would be including a loop that would capture the ticker name and create an array to be used from that.  This would mean that any stock data with the same columns could be run.  Considering the point was to make it so easy for the hypothetical client, Steve, to use the macro for other data so all he has to do is press the buttons on the worksheet and not need to mess with the macro itslef, this script could still use a little work. I may follow up with more refactoring in the future.

