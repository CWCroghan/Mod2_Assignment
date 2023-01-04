# Stock Evaluation for 2017 & 2018
## Project Overview
When advising clients about possible stock purchases, it is common to review the past performance of the stocks.  Although past performance of a stock does not guarantee the future performance, it can provide insight into the quality of the company in that companies that have shown to be successful in the past tend to be more likely to be successful in the future.

Initially, when approached about doing some stock performance calculations, the analyses were limited to only one year and only one stock ticker.  Being initially concerned about getting an answer with minimal programming effort, hardcoding was utilized to capture the specified features by year and stock ticker.  Then the client requested information concerning the full- range of available stocks over multiple years.  Instead of continuing with the original code, it was determined that we should refactor the original code with less hardcoding and more flexibility utilizing additional features of Visual Basic for Applications (VBA), for example arrays.

##Results:
  
![Table of Results for 2017 Stocks](https://github.com/CWCroghan/VBA_Challenge/blob/main/Results2017.png)
![Table of Results for 2018 Stocks](https://github.com/CWCroghan/VBA_Challenge/blob/main/Results2018.png)

The Annual Return of the stocks for 2017 were mostly positive with only one stock (TERP) having a negative return.  Whereas for 2018, the opposite was true in that only 2 stocks (ENPH and RUN) had positive returns and the rest were negative.  Between the two years, only 3 stocks (ENPH, TERP, and RUN) matched the direction of returns between the two years.  The Total Daily Volume of trades were consistent overall between the two years (2017 Mean – 1051K, STD - 1344K; 2018 Mean - 1097K, STD - 1248K).

One of the primary goals of the task was to refactor the code to achieve faster run times. There were three main modifications made to the code.  
1.	Removal of nested loop,
2.	Use of arrays to capture the information generated on each of the stock tickers, and
3.	Limiting the shifting between spreadsheets.
Originally, all the data were read and processed for each stock ticker.  Once  information for a stock ticker was complete,  the program output  results into a separate sheet.  This method of nested loops involved processing over 3,000 records 12 times with 5 comparisons each time totaling over 180K comparisons.  In addition, outputting each stock ticker individually necessitated switching spreadsheets 24 times. 
In comparison, the refactored code leveraged the organization of the data to reduce processing.   The data were sorted by stock ticker and date.  It could be read once and each stock ticker processed as it was encountered.
Since arrays were used to capture the calculated values for each of the stock ticker, there was no need to switch back and forth between data spreadsheets and output spreadsheets.  The data were written out with a single switch and a loop for each of the 12 stock tickers.  
With these coding changes, the refactored the code time to run was 0.1 seconds for each year as opposed to original code runtime of 6 seconds for each year.

![Image of Time for 2017 Stocks](https://github.com/CWCroghan/VBA_Challenge/blob/main/VBA_Challenge_2017.png)

#Summary: 
Some level of review should always be done in code that will be used on a regular basis.  Just like in writing, a first draft is rarely used without changes.  Reading through making sure that each element of the code is necessarily and not redundant can save processing time and make the program clearer.  The longer the program, the more likely that some level of redundancy has crept into the code.  Cleaning up code to improve readability by adding white space and comments is also an important part of review code.  This will aid in code updates and code reutilization.  
If during review, evidence is found that a more serious level of code editing is needed, then consideration of preforming refactoring should be made.  Refactoring code can lead to more streamlined, easier to understand, and more in compliance with other coding efforts. 
However, refactoring code should not be done in a vacuum.  There are of course costs of time and effort.  The bigger requirement is for the ability to think flexibly.  Finding a solution to a problem can lead the analyst to become blinded to other possible soluions.  For this reason, many companies have a second team work on the code-review and refactoring of code.  When that is not available, having a diverse team who are able to express diverse solutions can be an effective alternative.  When a team of one, it is critical to maintain a flexible thinking and from the start and explore alternative solutions to each problem.  This flexibility can help if one solution path becomes a dead-end or when refactoring problematic code.  It can be very difficult to self-edit, placing time between when the code is originally written and when you attempt to refactor can be beneficial, but it requires that time is available.  
When code will be utilized on a continuing basis, taking the time to refactor the code can see a great return in time savings.  If it is a one and done coding, then it is much harder to justify an extensive effort to refactor the code.
Scope creep was one of the primary reasons for needing to refactor this code.  Originally, we were asked to analysis the performance of a single stock for a single year.  The next ask was to analyze all the stocks for a couple of years.  That was done with some simple modifications to our original code written for a single stock.  When the idea was to use our simple code for unknown number of years and unknown number of stocks, it became clear that our simple code was too clumsy.  Refactoring was necessary to meet the new needs of the client with sleeker and leaner code.



