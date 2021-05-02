# VBA_Challenge
Using VBA for Stock Analysis
## Introduction 
In this project we are using VBA to help us analyze the stocks we are given. There are 12 different tickers and over 3,000 rows of data. Our goal is to look at each ticker and see the total volume and return that they are getting in each year since we have two sheets from two different years. In module 2 we completed most of this work, we compiled a code that made it easy for Steve to look into the spread sheet, click a button and See the total Volume and total return but now our goal is to refractor the code so that it is more efficient but can also hold more tickers if necessary.
## Findings 
Year 2017 findings show that most of the tickers had a positive return and were safe investments. The Particular one that Steve's parents were looking into was DQ which had the greatest return of all. For the 2018 you can see that the returns were not nearly as high and the total volumes were down as well. I think some of the reason for this was because if you look at the very top of the data in the 2018 spread sheet you can see the prices went up a bit.

<img width="317" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/82982480/116800828-1e18f580-aaca-11eb-91de-18ba35862925.png">

<img width="303" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/82982480/116800831-23764000-aaca-11eb-9310-d24e1648ff3f.png">

## Refractoring the code for this challenge: 
we refactored the code so that the code would be more efficient and run faster. One of the main things we did in this refactored code that made it different was creating the three output arrays (tickerIndex and tickerVolumes) as arrays. Having the extra array allowed us to loop through the rows, increase our tickerVolume and add the volume for the current stockticker. From that point we are able to create our if-statements to select the ticker that we are looking for each specific cell and get our total volume and return.
## The original code:
The original codehas only initialized startingPrice and endingPrice and did loop through all of the values, essentially making the process longer. We only had one loop and in the refactored code we have an inner and outer loop.
## Conclusion of using refactored code:  
The refactored code had a much quicker run time because of the changes. This is a very positive outcome and means the code is much more efficient and most likely will run better if we have more tickers.
 <img width="1009" alt="Screen Shot 2021-05-01 at 9 54 33 PM" src="https://user-images.githubusercontent.com/82982480/116800850-46085900-aaca-11eb-84aa-d55d4ff4a1ad.png">
<img width="1033" alt="Screen Shot 2021-05-01 at 9 55 08 PM" src="https://user-images.githubusercontent.com/82982480/116800853-4a347680-aaca-11eb-939c-84a3f5e48bc1.png">
