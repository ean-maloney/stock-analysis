# Stock Analysis
## Overview of Project
The purpose of this project was to create a VBA subroutine that would analyse an Excel worksheet containing daily trading information of a list of green energy stocks to extract various measures for each of these stocks based on the dataset. I tested two different code designs that successfully accomplished these tasks and compared their runtimes. 

## Results
Below I have attached the output of my code-generated analysis of a set of green energy stocks from 2017 and 2018.

<img width="240" alt="All Stocks (2017)" src="https://user-images.githubusercontent.com/80861610/116569836-c31fac80-a8d7-11eb-8f79-131a7c774eca.png"> <img width="240" alt="All Stocks (2018)" src="https://user-images.githubusercontent.com/80861610/116569841-c450d980-a8d7-11eb-8d21-a2615d61aca0.png">

As these graphics illustrate, 2017 tended to be a good year for the stocks I analyzed, while in 2018 most of these stocks lost value. Additionally, these stocks were highly volatile over this two year period, as is evidenced by the many double- or even triple-digit percentage changes in value over the course of a year. Only two of the stocks I analysed gained value in both years (ENPH and RUN), while only one stock lost value in both 2017 and 2018 (TERP).

The graphics were created using a code that created four arrays containing stock tickers, trading volumes, starting prices, and ending prices for each of the stocks analyzed. 

Two methods were tested for gathering this data. Both methods begin by initializing an array with the relevant stock tickers hard-coded into it. 

The first method then iterates over this array stock by stock and, for each ticker, searches a spreadsheet containing the daily trading data of the stocks over the course of the year. While doing this, it adds up daily trading volume for each ticker and finds its closing price at the first and last trading days of the year. The return percentage is then calculated by comparing these closing prices for each stock. 

The second (refactored) method works by initializing three additional arrays: one for total volumes, one for starting prices, and one for ending prices. The index each stock in the first array corresponds to the index of its data in each of the other arrays. The submodule then searches through the spreadsheet of trading data and adds up the trading volumes for each stock and records its starting and ending prices for the year. Whenever a new stock ticker is encountered on the spreadsheet, the value of a variable corresponding to the index of the array for each stock is changed.

Below are screenshots of the runtimes of the first method for 2017 and 2018.
<img width="226" alt="VBA_Challenge_2017_unref" src="https://user-images.githubusercontent.com/80861610/116588791-f79c6400-a8e9-11eb-81f6-356e9d6edca6.png"> <img width="230" alt="VBA_Challenge_2018_unref" src="https://user-images.githubusercontent.com/80861610/116588797-f9662780-a8e9-11eb-9ee1-78ea30405e3d.png">

And here are the (significantly faster) runtimes of the second method.
<img width="235" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/80861610/116588962-27e40280-a8ea-11eb-8485-ae0325da0233.png"> <img width="235" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/80861610/116588883-113dab80-a8ea-11eb-88bc-e5b87f82b7ec.png">

## Summary
### Advantages and Disadvantages of Refactoring
In general, refactoring a code is advantageous because it tries to make the code more efficient in terms of its usage of time and memory. The disadvantage of refactoring is that, in making the code more efficient, it may also become more difficult for a human reader to parse, especially if operations have been combined in ways that make what the code is doing less explicit to a reader.  

### Advantages and Disadvantages of the Methods Used
With respect to the code written for this project, the main advantage of the refactored code was that it ran much faster than the unrefactored code. Another advantage is that it is somewhat simpler for a human reader to parse, since it eliminates the need for nested loops.

One disadvantage of the refactored code is that, if we wanted to run it over a different number of stocks than the twelve analyzed, we would have to change the sizes of four different arrays in the code, whereas, with the unrefactored code, we would only have to do this for one. This problem could be solved by changing the code so that it determines the size of these arrays dynamically at runtime instead of through hard-coded values.
