# Stocks-Analysis Using VBA/Excel
## Purpose
To analyze stocks for  2017 and 2018 to determine what stocks performed the best for the last two years and thereby help the decision maker make an informed decision about the right stocks to invest in. 

As part of the analysis, since we are using VBA, we would need to determine of the code as created would run faster if refactored since refactoring is supposed to take fewer steps.
## Challenge Background
Steve loves the workbook we prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although this code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.
I attempted to improve the code by using better logic and may have added newer functionality on formating. 
## Results:
It looks like steve may have to take a look at stocks other than the DQ based on the findings. There were other better performing stocks that would be better investments for his dad.
## Code
The finalized code (refactored code) begins as followed:
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
 
 It looks at only the few stocks, but will can apply to many when used on a larger number of stocks with little to no tweak to the code. 
 
 I found that the time was about 10 preseconds fastor on the original code vs the refactored code. I believe it is because my original code was a bit leaner as it did not have much in the way of formatting. This finding was on both the 2017 and 2018 years. 
 <img width="522" alt="image" src="https://user-images.githubusercontent.com/36766602/149449981-bc0f4063-ef9c-4b76-96a3-b89f2b908e23.png">
 <img width="518" alt="image" src="https://user-images.githubusercontent.com/36766602/149450046-563c8765-7c73-4d69-9b46-5b0dfbd37944.png">
 
 Overall some of the main advantages of refactoring code are as follows:
 Why is it important? There are several reasons why regular code refactoring is important in software development:
- Simplified support and code updates. Clean code is much easier to update and improve. Developers can quickly make new functionality available for users, as well as save the support budget, as the maintenance will require less working time spent by the programmers involved.
- Saved time and money in the future. Code refactoring reduces the likelihood of errors in the future and simplifies the implementation of new software functionality. Instead of making sense with tangling code or fixing bugs, developers can start implementing the required functionality at once.
- Reduced complexity for easier understanding. If the team engages a new employee, or the entire development team changes altogether, it will be easier for new developers to comprehend the code and make the necessary alterations faster.
- Maintainability and scalability. At times, programmers simply avoid making alterations to some dirty code since they do not clearly understand what consequences these modifications will lead to. The same is true for scalability. Removing this obstacle is another benefit of code refactoring.

The main disadvantage of refactoring is that it takes too much time and can take away from the urgency of the project and thereby slow the project down.
 

