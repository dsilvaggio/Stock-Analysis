# Stock-Analysis
## Overview of Project
  An analysis of different stock prices and volume was completed in order to determine how the stocks of different green energy companies performed over the years 2017 and 2018.
### Purpose
  Steve is a recent finance graduate who is helping his parents decide which green energy companies they should be buying stock in. They had originally bought stock in DQ, but after finding that this company has not been performing well, data on 11 other green energy companies was perfomred in order to help his parents choose the best company to invest in. Steve was also interested in adding data on how the entire stock market has performed over the years 2017 and 2018, in order to provide his parents with more research. 
## Analysis of Stock Performance
  Based on the code that was ran, I was able to see the total daily volume and percent return for each of the 12 stocks included in this data set for the years 2017 and 2018. A table of the 2017 daily volume and percent return can be seen below:
  ![This is an image](https://github.com/dsilvaggio/Stock-Analysis/blob/4afa2d43b102bf801613e1974b7d5cd1b191d368/Resources/2017%20_data.png)
  
  All but 1 of these stocks in 2017 had a positive yearly return, meaning that all but 1 of these stocks experienced an increase in their price. This means that these 11 stocks would have netted an increase in your investment. 
  However, the year 2018 showed a much different picture. A table of the 2018 daily volume and percent return can be seen below:
  
![This is an image](https://github.com/dsilvaggio/Stock-Analysis/blob/e1da3a7919ebceecf866c633b42ad3b38a199cc9/Resources/2018_data.png)

  In 201t8, only 2 stocks on this list had a positive yearly return, which means that most of the stocks listed experienced a decrease in their stock price. We can also see that the total daily volume for these 10 stocks dramatically decreased from 2017 as well. This means that these 10 stocks were traded significantly less in 2018 then they were in 2017.
   There were 2 stocks that maintained a positive yearly return as well as significantly increased their total daily volume between 2017 and 2018. These 2 companies are referred to with the ticker symbol ENPH and RUN. These would be companies that I would suggest Steve's parents look at investing into.  
## Analysis of Refactored Code
  Prior to refactoring the code, our orignal code was only running across 1 array that we had named "ticker". This was allowing us to quickly see the daily volume and percent return of each of the 12 tickers listed in the spread sheet. When running this code, my computer was taking around 0.7 seconds to display the information for both 2017 and 2018. I then decided to add 3 more arrays that would allow us to run this code across multiple different ticker values. This would allow us to use the data of 1,000's of different stocks instead of just the 12 included in this worksheet. When refactoring the code, I did not need to write completely new code. I was able to reuse the original code, but I just needed to replace the original variables of "total volume", "startingPrice" and "endingPrice" with the new arrays that I had created. The time it took to run the new code is below: 
  
 ![This is an image](Resources/2017_run_time.png)
 ![This is an image](https://github.com/dsilvaggio/Stock-Analysis/blob/53efd5cb433a5609de63c921a495cbed0e5a23f2/Resources/2018_run_time.png)
 
 This was significantly less time than my first code. My refactored code can be seen below. 
 
```
  '1a) Create a ticker Index
        tickerIndex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
        Next i
        
    '2b) Loop over all the rows in the spreadsheet.
    Worksheets(yearValue).Activate
    For i = 2 To RowCount

 '3a) Increase volume for current ticker
            If Cells(i, 1).Value = tickers(tickerIndex) Then
            	tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            End If

   '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value 
        End If
        
   '3c) check if the current row is the last row with the selected ticker
	 If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
    '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1 
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
   Worksheets("All Stocks Analysis").Activate
    For i = 0 To 11
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
```

## Summary
The biggest advantage that I noticed with refactoring code is that it significantly decreases the amount of time it takes to run the code. I could see this being extremely helpful if you are working with 1,000s of data points and would not want your computer to take to long or time out when trying to run the code that you have written. Another advantage of refactoring is that you do not need to rewrite the entire code to make it something completely new. I was able to make edits to much of the code that I had previously written. 
Some disadvantages or challenges that I had when refactoring was understanding the difference between a single variable and an array. At first, I was trying to replace the original variables with the new "tickerVolume", "tickerEndingPrices", and "tickerStartingPrices" that I had created. When this was not working, I realized that when referencing an array you also need to include the index that you are referencing within the array. This is when I was fully able to see what was happening with the tickerIndex that I had created. When I went back and added the tickerIndex in paranthesis next to the arrays I was referencing, my code was able to run successfully. When replacing singular variables with arrays, it can be challenging to separate when you are referencing a single variable and when you are referencing an array. 
Overall, the biggest advantage of this particular refactored VBA script is that it ran in significantly less amount of time. We would also be able to add additional information on various other stocks and still run the same code in a short amount of time. 

