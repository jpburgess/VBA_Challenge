# VBA_Challenge

## Overview
-demonstrate the usefulness of refactoring code
-we made a vba macro that analyzes data for a few stocks and our client liked it so much they want to use it for analyzing the entire stock market over a few years


## Results
=original code

  For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           If Cells(j, 1).Value = ticker Then
               totalVolume = totalVolume + Cells(j, 8).Value
           End If

           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               startingPrice = Cells(j, 6).Value
           End If

           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               endingPrice = Cells(j, 6).Value
           End If
       Next j
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
   Next i


=refactored code

    For i = 2 To RowCount
        ticker = tickers(tickerIndex)
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
            tickerStartingPrices(tickerIndex) = tickerStartingPrices(tickerIndex) + Cells(i, 6).Value
        End If
        
        If Cells(i + 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
            tickerEndingPrices(tickerIndex) = tickerEndingPrices(tickerIndex) + Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        End If
    Next i
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        ticker = tickers(i)
        Cells(i + 4, 1).Value = ticker
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i

=stock data


## Summary
refactoring
pros
-increase performance
-improve readability
-Can consolidate repetitive sections into one 
cons
-Diminishing returns
-Chance of breaking working code


original code
pros
-Completes the objective
cons
-slow / this will compound as more stocks are analyzed 

refactored code
pros
-faster
-more readable
