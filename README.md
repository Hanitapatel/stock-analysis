# Overview of Project
Purpose of this analysis to help Steve on his analysis to help his parents to analyze a **DAQO _New Energy Corp stok_**.
Steve also wants to analyze some more green energy stocks for his parents to diversify their investments. On this project we are helping steve on his 
on his analysis by using VBA. We can write code that will automate these analyses for us also we are refactiring our code to make our VBA script run faster.

# Results
## DQ Analysis
Based on the analysis we did for DQ it appears that it did perform very well in 2017 but not so much in 2018. in the DQ analysis work sheet we can see that this
particular stock has total volume of 107,873,900 and -63% return in 2018. 

DQ Analysis Excle Image

DQ Analysis VBA Image

## All Stocks Analysis
In the All Stocks Analysis worksheet, we can see total volume and Return for all the given tickers for the year of 2017 and 2018. 
-	In 2017 DQ has the lowest total daily volume and highest return of 199% and ticker TERP has the lowest Return of -7% . 
-	In 2018 DQ has lowest return of -63% among all the given tickers and ticker RUN has highest return of 84%. 

Based on this analysis we can see that DQ performed very good in 2017 but in 2018 its was the worst performer but we can also see that most of the green energy
stocks did not perform well in 2018.  Our analysis shows that there are some other good green energy stocks out there to invest in example of ticker ENPH which
did perform good in both years. Tickers RUN, SEDG, VSLR also needs to consider investing in.

All Stocks Analyisis Excel Image 2017

All Stocks Analysis Excel Image 2018

By using VBA we automated the entire analysis for Steve. Created button for Run the analysis and to clear the analysis also created InputBox  which  Steve can 
use to enter the year value to run the analysis for 2017 and 2018 separately. We also wrote the code to find the elapsed time for 2017 and 2018 and create 
message box that shows the time every time we run analysis for either of the year

All Stocks Analysis VBA image

In the challenge I refactor the code that we did for All Stocks Analysis to make it more efficient. 
        The execution time to run original script for 2017 is 0.53125 and for 2018 is 0.46875 and the execution time to run the refactor script for 2017 is 
        6.640625 and for 2018 is the same.
The refactor script is taking longer because it include all the codes for Inpubox ,  start and end time and formatting. 

# Advantages of refactoring code
- Make the code more efficient by improving the logic
- Taking fewer steps, create code easy to read for the users 
- can add more functionality

## Disadvantages of refactoring code
-	Time consuming
-	It can mess up the purpose of the original code

# Advantages of refactored VBA script
-	Automate excel data base
-	Help analyze data faster 
-	Save the user’s time

## Disadvantage of refactored VBA script
-Time consuming
-It can crash your original script







