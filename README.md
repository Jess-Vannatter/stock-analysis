# stock-analysis
Module 2
# VBA of Wall Street

## Overview of Project

### Purpose and Background
-   The Main objective of this project was to "refactor" a Macro written to take a deeper look in to a specific data set of "Green Energy Stocks" () and make it more efficient/ applicable to a larger data set such as the entire market. the indended purpose of the Macro was to  establish if these "Green" stock options were a viable choice for our friend, Steve's parents to invest in. We were provided 2 data sets, seperated by year (2017 and 2018) of 12 different "Green Stock" options. The data sets were provided to us in one workbook, on two worksheets labled "2017" and "2018". The data was formulated in a large table displaying each stock's (labled as tickers) preformance for each "trading day" of the year. With there being 252 "trading days" within a year, each stock had 252 rows of data to work with. For these table's, the stock's preformances were characterized in the columns with lables such as the specified date the data in that row was pulled from, along with "opening" and "closing" ( also adjusted closing value ) values for the day, high and Low points of the day, and a daily volume figure as well.

- As stated above, our goal was to extract the data from these data sets and create a much more "User friendly" comparable table that provides an annual analysis of the 12 stocks. A macro was already created for this analysis but we wanted to create a more optimized version of that macro for our friend to provide to his clients (his parents). The Macro created was able to successfully extract the data for each stock/ ticker and provide the "total Daily Volume" (This was done by getting the sum of each days volume for the specified stock 'Cells(4 + i, 2).Value = tickerVolumes(i)') and rate of "return" (this was done by establishing the daily "opening" and "closing" values for each stock and dividing the "closing" price by the "opening" price and subtracting 1 'Cells(4 + i, 3).Value = tickerEndingPrice(i) / tickerStartingPrice(i) - 1'.) for the whole year. The extracted data was then "output on to a table on a new sheet labled "All Stocks Analysis" that detailed the "total Daily Volume" and Annual "rate of return" ( in percentage ) for each stock in a dynamic fashion that allows the user to choose between which year (2017 or 2018 ) you would like to run the macro on to determine the stocks preformance "All Stocks Analysis"().

## Results

### Comparison of 2017 and 2018 Stock preformace
- Once the Macro was successfuly implemented and ran for the two years ( 2017 and 2018 ) we were able to see that for the most part the 12 stocks preformed , as a whole better in 2017 as compared to 2018 even though the :total volume in 2018 was higher overall than in 2017. In 2017, all but one of the stocks had a negative rate of return ("TERP" at -7.2%) (All Stocks Analysis). While in 2018, only 2 of the stocks had a positive "rate pf return ("ENPH" at 81.9% and "RUN" at 84%). The Average "total volume" for the stocks in 2017 was $263,886,591.67 "2017 descriptive Statistics table"(), while in 2018 the Average was $275,503,183.33 '=AVERAGE(B4:B15)'"2018 Descriptive Statistics table" (). Which is an increase of $11,616,591.66. But while the annual total volume of the set of stocks did increase from 2017 to 2018, the "rate of return" decreased significantly from 67.3% to -8.5% in 2018 '=AVERAGE(C4:C15)'. Essentially, this means as a whole the set of stocks were traded/ had more activity in 2018 compared to 2017, but preformed better in 2017. 
- IF we were to look at specific stocks in this data set that steve could possibly suggest a client ( his parents ) to invest in, there are two in particular that were highlighted above. Both "ENPH" and "RUN" had positive "returns" in 2017 and 2018, meaning they both made money for their investors for each year. IN addition, they both showed pretty significant growth in they "Total Daily Volume" As well. With "ENPH" increasing from a "total Daily Volume" of $221,772,100.00 in 2017 to a "Total Daily Volume" of $607,473,500.00 in 2018. Similiarly, "RUN" presented an increase in "Total Daily Volume" as from $267,681,300.00 in 2017 to $502,757,100.00 in 2018. These two stocks stand out as options for Steve to suggest to his clients ( his Parents ) to invest in (AllStocks Analysis()).
The overall outlook on these data sets could have been a result of outside factors, such as the preformance of the overal market in 2018, as compared to 2017.But for the purpose of our analysis, it looks like the two specified stocks, "ENPH" and "RUN" preformed very well in both years, in spote of a possible "*down*" year for the "Green Stocks" market in 2018.

- ![2017 Descriptive Statistics Table](https://user-images.githubusercontent.com/117245167/204158143-ce1d8b5a-a16e-4c5f-9d28-1033664517d2.png)
 
- ![2018 Descriptive Statistics Table](https://user-images.githubusercontent.com/117245167/204158146-5b80d316-c9c9-48e4-954d-3d6cb34dd9fb.png)


### Analysis of Script Run times 
- When running the Refactored code we saw a significant difference in efficiency. Initially, when running the original "All Stocks Analysis" ( written in the module work ) the run time for the 2017 data set was "0.5507813 seconds" (Original 2017 All Stocks Run Time.png) and "0.5429688 seconds" for the 2018 data set (Original 2018 All Stocks Run Time.png). As compared to the Refactored code, where our code for the 2017 data set ran in "0.1835938 Seconds" (VBA_Challenge_2017.png) and "8.984375E-02 Seconds" (or 0.08984375 Seconds) (VBA_Challenge_2018.png). I believe this to be a significant difference in efficiency which is why I did not look to create another data set with average run times across a number of run to compare. Our Refacroted Macro ran 0.3671872 seconds faster for the 2017 data set, and 0.045312505 seconds faster for the 2018 data set [Resources] (https://github.com/Jess-Vannatter/stock-analysis/tree/main/Resources).


- Our Refactored Macro for the "All Stocks Analaysis" was more efficient essentially becuase instead of using a nested For Loop to run through the data set 12 times (for each Ticker) to gather the data needed for our output ( see below) (All Stocks Analysis()).
```  
'For i = 0 To 11
        
        ticker = tickers(i)
        
        totalVolume = 0
        
        
            '5)Loop through rows in the data.
    
        Sheets(yearValue).Activate
        
            For j = 2 To RowCount
            

            '5a)Find the total volume for the current ticker.
    
                If Cells(j, 1).Value = ticker Then
        
                    totalVolume = totalVolume + Cells(j, 8).Value
                    
                End If
                
    
            '5b)Find the starting price for the current ticker.
    
                If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
        
                    startingPrice = Cells(j, 6).Value
            
                End If
                
    
            '5c)Find the ending price for the current ticker.
       
                If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
            
                    EndingPrice = Cells(j, 6).Value
            
                End If
        
            Next j
            
    
            '6)Output the data for the current ticker.

        Worksheets("All Stocks Analysis").Activate
    
        Cells(4 + i, 1).Value = ticker
    
        Cells(4 + i, 2).Value = totalVolume
    
        Cells(4 + i, 3).Value = EndingPrice / startingPrice - 1
    
    Next i
```

 - Instead we added 3 arrays (in addition to our initial Ticker array) for our desired outout ("tickerVolumens, tickerStartingPrce" and "tickerEndingPrice") and used a variable (tickerIndex) to track our position within our desired output data so that we just had to go through the data set once.
 
```
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    
        tickerIndex = 0
        

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    
    Dim tickerStartingPrice(12) As Single
    
    Dim tickerEndingPrice(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        
        tickerVolumes(i) = 0
        
    Next i
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
                
            End If
            
            
         '3c) check if the current row is the last row with the selected ticker
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
                'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                        
                        
                        '3d Increase the tickerIndex.
                    tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
                    
                    tickerIndex = tickerIndex + 1
            
            
            End If
    
    Next i
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
       Worksheets("All Stocks Analysis").Activate
                       
            Cells(4 + i, 1).Value = tickers(i)
    
            Cells(4 + i, 2).Value = tickerVolumes(i)
    
            Cells(4 + i, 3).Value = tickerEndingPrice(i) / tickerStartingPrice(i) - 1
        
    Next i
```

## Summary

- What are the advantages or disadvantages of refactoring code?
  - Advatage 
            - The main advantage to refactoring code is to make a specific code more efficient. Its allowing your code to work smarter, not harder. Which in turn does the same for your computer. This can be even more important when working with larger data sets/ more complex codes since they may take more resources for it to run. By Making your code more efficient, it could allow it to work with larger data sets that you may not have been able to work with before. IN addition, Refactoring a code script by removing/ replacing steps with more logical functions can make a code easier to read by possible co-workers or peers.
   - Disadvantage
            - Coding can take time. So when considering the time it took to write our initial Macro (that did what we wanted it to do). we then spent more time and resources on the same code. So even though the code, once refactored is of better quality. In some situations we may not have the time to Refactor and go throught the code all over again.


- How do these pros and cons apply to refactoring the orignal VBA Script?
    - We were able to successfuly "refacter" the initial macro and make it run faster/ more efficient. By doing this, we allow Steve to possibly apply this Macro to a larger data set/ Larger number of stocks and/ or previous years. Although, our code may be limited in other situations where the data set we are working with is not as organized as the one we recieved from Steve. We wre able to write this macro mainly because there was a set of "identifiers" (Tickers) that we could organize alphabetically. Without such an identifer we wouldnt have been able create the initial array which was an essential part of both of the Macros we created. In addition, As a whole our Macro is pretty specific in dealing with stocks/ the stock markt and may not be applicable to other data sets.
