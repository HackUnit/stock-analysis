# Stock Analysis


## Overview
The purpose of this analysis projected from an original desire by a friend to find the total daily volume and yearly return on each stock in a data set of stocks. The scope of the project begun by just focusing on a curiosity of your friends’ parents to a fledged-out macro that utilized the datasets provided. From a diminutive script to a fledged out dynamic report, this project’s original purpose was finally realized of your friend’s original intent. Furthermore, the code was refactored to help tidy and tighten up the original code to have a more efficient macro that scaled better.



## Results of the Stock Analysis

### The Beginning
The beginning steps of creating the code started with specifically finding the return on a DAQO stock in 2018 from provided datasets that covered a variety of stocks—including the DAQO stock—and data that could be used to create a solution. Using the data scraped through scripting, the total daily volume and yearly return was found. The numbers were then formatted into a separate final worksheet. At this point, only the DQ stock data had been scraped and compiled, which left the grand majority of the datasets unused.

### The Project Scales
At this point the project turned to creating code that would include the whole dataset, which in turn would be boiled down to their essential parts in relation to their specific stock. The resulting code was maybe not as elegant of a solution, but the point then was to get a script working that was capable of the task. For this first attempt, `for` loops were heavily utilized to perform the loops necessary to capture the whole dataset.

```
For i = 0 To 11
                
                ticker = tickers(i)
                totalVolume = 0
                    
                    '5) loop through rows in the data
                    Worksheets(yearValue).Activate
                    
                    For j = 2 To RowCount
                    
                        '5a) Get total volume for current ticker
                        If Cells(j, 1).Value = ticker Then
                        
                            totalVolume = totalVolume + Cells(j, 8).Value
                            
                            
                        End If
                        
                            '5b) Get starting price for current ticker
                            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                                
startingPrice = Cells(j, 6).Value
                            
                            End If
                        
                                '5c) Get ending price for current ticker
                                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
 
                                    endingPrice = Cells(j, 6).Value
                                
                                End If
                 
Next j
```
<sub>		* Original code *</sub>

If statements were used within the `for` loop to collect the necessary data as the macros went over the data within each set. This key script was fleshed out so when run it would query an end user for which year they wanted, then using the data from that specific yea, use `for` loops to answer what the total daily volume and return for each stock was. This data was then displayed into an efficiently designed table with static formatting, and conditional formatting to get information across faster. Finally, ease of use was implemented for the end user with buttons to run the macros.

[Pic 2017 report]

Using the completed macro, the performances of all the stocks could be compared between the two years provided.

[Pic 2018 report]

Noticeably that 2017 was a better year across the board for nearly every stock. But this was not the end of the analysis. At least not the analysis on the macro itself. 

The code was slow, relatively. It could be tightened up further, especially if it was to be scalable. This would be a necessary step, due to the nature of the data that was being computed. Stocks are not small data sets. 

[Pictures of times for original code both years]

OG 2017 = 626.25 ms
OG 2018 = 671.87 ms

Within the refactored code, the main difference was the removal of the nested `for` loops, and its replacement with `if-then` statements utilizing arrays. The made the task of collecting all of the data a tad faster.

```
For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
        End If
        'If the next row’s ticker doesn’t match, increase the tickerIndex
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
```

The difference between the original code can be clearly seen in the actual milliseconds. Roughly six times faster for both datasets.

Refactored 2017 = 109.37 ms
Refactored 2018 = 93.75 ms

[Pics of refactored times]


 ## Summary
	
### Advantages and Disadvantages of refactoring code

### On the advantages and disadvantages of the original and refactored VBA script

