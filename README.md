# The VBA of Wall Street

## Overview of Project

### Purpose

In this project, the purpose is to analyze various stock market tickers to help a client determine how best to invest their funds. In order to do this, we take a history of ticker updates for a given year, sorted ascending by name and secondarily by date, and create and run a macro, created with Visual Basic, that produces a total volume of stocks moved for the given year, as well as a percent return over the given year taken from the first recorded price and the last recorded price. However, the previous macro we made to do this took a considerable amount of time to complete, so for this project, we decided to try a macro with a single iterative loop to see if it saves time.

## Results

### Analysis of First Macro

For the first version of our analysis macro, we decided to go with nested 'for' loops. Since we have twelve tickers, the outermost loop would cover all twelve tickers one by one, while the innermost loop would cover the data proper. In order to make it so our macro could use all of the data for any given year, we opted to format the for loop as such:

```
startRow = 2
'endRow code found at https://stackoverflow.com/questions/18088729/row-count-where-data-exists
endRow = Cells(Rows.Count, "A").End(xlUp).Row

For i = 0 To 11
     ...

     For j = startRow to endRow

          ...

     Next j
Next i
```

From this, we comb through our data for a number of pieces of data. Namely, we want volume amounts for each instance of each respective ticker, as well as the first and last price instances for that ticker. The former we sum together to find a total volume as such:

```
totalVolume = 0

...

If Cells(j, 1).Value = ticker Then
     totalVolume = totalVolume + Cells(j,8).Value
End If
```

For the latter, we do two other checks, one to see if it's the first instance of the ticker in question, and the other to see if it's the last instance of the ticker. If the former is true, then its closing price is recorded as the starting price, and if the latter is true, its closing price is recorded as the ending price. In VBA, this is done as such:

```
If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then

     startingPrice = Cells(j, 6).Value

End If


If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then

     endingPrice = Cells(j, 6).Value

End If
```

At the end of each loop of the data, we print the data to the "All Stocks Analysis" under each respective ticker, and since our macro consists of twelve tickers, we do twelve loops of the data. At the end of those twelve loops, we produce a message box telling the user how long it took to go through the macro. In the case of the data we have from 2017, the code runs in about 0.44 seconds, while the data we have from 2018 causes the code to run in roughly 0.43 seconds. In terms of computer time, this is incredibly slow and adds up when multiple data sets need to be analyzed, and these data sets only comprise twelve tickers across around 3,000 entries for either data set. With roughly 250 entries per ticker, if we include more tickers with similar numbers of entries for each, even looking at twice as many tickers, we're looking at around 1.7 seconds, since we're combing a data set twice as big, twice as many times. The New York Stock Exchange, or NYSE, itself, trades for over 2,800 companies, [according to ADVFN](https://www.advfn.com/nyse/newyorkstockexchange.asp), so using this code, we could expect to analyze both the 2017 and 2018 tickers of the NYSE in roughly 23,411 seconds apiece, or 13 hours and 22 minutes, or both in *about 26 hours and 44 seconds.* Clearly, if this were to be used in large-scale stock analysis, it would be catastrophically slow. Suffice it to say, combing through the data repeatedly takes time that a client might not have.

### Analysis of Single-Iteration Macro

With the realization that multiple iterations over sufficiently large data sets in mind, we set about making a macro with a single iteration through the data set. In order to do this, we had to figure out another way to comb through the data. Thankfully, a way was already put into the code we'd already written -- we just needed to implement it. If we set up a ticker index, we'd simply need to dictate when the ticker index changed.

To do this, we simply needed to change this:
```
If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then

     endingPrice = Cells(j, 6).Value

End If
```
To this:
```
If Cells(j + 1, 1).Value <> ticker Then

     endingPrice = Cells(j, 6).Value
     tickerIndex = tickerIndex + 1

End If
```

Since our data is already sorted alphabetically, we can get away with this, since our tickers were put into the array alphabetically. Additionally, since we know our data is sorted by date, and we aren't looking at the entire data set once for each ticker, we can omit the checks to see if the ticker is correct beyond the code snippet above. However, we quickly run into a new issue: we can't simply use a single variable each to hold all the ticker data we need. Well... we could, in retrospect, but it's finicky and not ideal. We opted instead to create an array each for the volume, starting price and ending price of each ticker, and store them that way. By doing this, we can use a for loop to put each down at the same time:

```
For i = 0 To 11
    
     Worksheets("All Stocks Analysis").Activate
     Cells(i + 4, 1).Value = tickers(i)
     Cells(i + 4, 2).Value = tickerVolumes(i)
     Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

Next i
```

With that, besides some formatting shenanigans in order to make the output sheet legible, the code changes were completed. Now to see how more or less efficient it was.

[Results from 2017](https://github.com/SirNancyTheNegative/stocks-analysis/Resources/VBA_Challenge_2017.png)

That's 0.0898 seconds. Considering the previous iteration of the code finished analyzing the 2017 data in 0.44 seconds, that's almost 5 times as fast. In order to see if this is a fluke or not, let's run it again to analyze the 2018 data.

![Results from 2018](https://github.com/SirNancyTheNegative/stocks-analysis/Resources/VBA_Challenge_2018.png)

0.0859 seconds. Compared to the first iteration of the code when looking at the data for 2018, that's, again, 5 times as fast. Even if we assume some of the more minor portions of the code don't add any time, with the data set of the New York Stock Exchange, we'd, in theory, only have 700,000 iterations to account for, and the data combing could be done within roughly 20 seconds. Now, there's no way to know this for sure without testing, but theoretically, it could be that short a time. That's much more reasonable than the 13 hours it would take for doing a single year analysis with the previous code. Granted, we'd have to re-rework our code in order to support 2,800 different tickers, but the fact of the matter is that considerable time would be saved in such a theoretical scenario.

## Summary

When it comes to refactoring code, there are a number of advantages to trying to make it more time efficient. For one, for huge data sets (e.g. 100 items spanning 100,000 entries), the time one spends breaking down a series of nested for loops into a single for loop with other iterative submeasures taking place inside is time well spent, as, if there's an issue with the code that doesn't appear until late into runtime, the time it takes for it to go through it all is severely lessened. Additionally, even if a client who uses the code only needs it to deal with small data sets, having more time-efficient code follows the adage 'A stitch in time saves nine', especially if the client decides to use it for much larger data sets. It's not pleasant to try to explain to a client that something only works so fast with smaller data sets, and that using something 100 times as big would take multiple real life hours. 

However, there are multiple downsides to refactoring as well. For one, it's time spent trying to find a solution to a problem that might not even exist. If something is already as fast as it can be, there's not much you can do that'll remedy that problem without stripping down features. The second downside, which ties in rather neatly to the first, is that it often requires convoluted workarounds that might not work, which means that it might not be anywhere as clean as it could be. A minor downside worth mentioning is that sometimes, it might be more work than it's worth. As the inverse of the "Stitch in time" example in the advantages, maybe a client will never go outside of what they wanted to work with. All that time optimizing the runtime might never come into play, but the reason it's only minor is that even if it doesn't, it still gives practice with solving the problem of optimizing, so that, in a meta sense, the process of optimizing runtime for other projects might also become that more optimized.

In the case of this project specifically, there are a number of advantages and disadvantages that exist that are unique to those put above. For one, with longer and larger sets of data, optimizing the code cuts down on a longer wait the more tickers there are. For each ticker, with the unoptimized code, we run through the data an additional time, and if we look through the entire year for each ticker like we have in this project, the time spent combing through the data starts to add up tremendously. Additionally, the refactored code cuts down on the number of times the following line of code is written:

```
If Cells(j, 1).Value = ticker Then
```

Which does wonders both for readability and any time spent making comparisons. On the other hand, the disadvantages are rather apparent. As the number of tickers we need to keep track of increases, the number of arrays we make may take a toll on the memory allocated to run the macro in the first place. Whereas only one is needed in the original code, the need for four arrays in the refactored code means that in lieu of time, we're running the risk of running out of memory for keeping track of everything, especially if we have to make more arrays to hold more data. Additionally, if we don't have properly sorted data, we run the risk of getting a bad analysis, not just in the refactored code but also with the original, as the original would consistently return incorrect starting and ending prices, while the refactored would return incorrect values for all three. Beyond that, though, the refactored code also runs the risk of extending outside of the limits of the ticker arrays, thus producing an Array Index Out of Bounds error before we could get the bad results in the first place. In all, however, with appropriately sorted data, the refactored code stands to be more time efficient than the original, especially if larger data sets are brought into the picture.
