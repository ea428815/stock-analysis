# stock-analysis

## *Overview of Project*


#### In this project, Steve requested for help with his analysis of stock data for his clients. Steve’s clients want to invest in green energy and DQ stock. Steve requested for a thorough analysis of the data in Visual Basic for Application (VBA) to ensure the accuracy of his future clients’ investment plan. In VBA, we created a script to analyze stock data based on the year the client inputs. Then we refactored the VBA script to increase the efficiency of running the code. The purpose of the project is to understand efficiency in coding where Steve can make decisions about which green energy and DQ stock to recommend to his clients.


## *Results*


The first line of code established the total volume of the stock, showing the number of shares that correlate to the strength of the stock. Below is the line of code that corresponds to the aforementioned:
```
If Cells(j, 1).Value = tickerIndex Then
tickerVolumes = tickerVolumes + Cells(j, 8).Value
End If
```

Next, to calculate the starting price of the stock data, we added the code below to correspond to the beginning of the year in order. To do this, the script is made to look at the line of data above where the stock is listed to see if line is different from the stock being calculated. The starting price of the stock is later compared to the ending price in order to determine the success of individual stocks. The code needed to perform this function is:
```
If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
tickerStartingPrices = Cells(j, 6).Value
End If
```

Additionally, the ending price of the stock is determined by looking at the line of data below the stock to see if the value is different. If the stock of the following row is different, then this signals to the code that this is the last line of data for a particular stock. Below is an example of this code:
```
If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
tickerEndingPrices = Cells(j, 6).Value
End If
```

### Recommendations on Stock
Based on the stock data from the years 2017 and 2018 below, Steve’s clients should consider investing in either RUN or ENPH, since both showed growth within two years. In fact, ENPH reported a return of 129.5% in 2017 while 81.9% in 2018 whereas RUN reported a return of 5.5% in 2017, and 84.0% in 2018. Observe below the analyses for both years, along with the time it took for the code to complete the analysis.



![allStocks2017_Resource](https://user-images.githubusercontent.com/62758795/202239948-df882874-0164-4d00-bc8c-317e377fbe05.png)

![allStocks2018_Resource](https://user-images.githubusercontent.com/62758795/202240073-61e964ac-5491-4bd5-b3d6-a956746a033e.png)


Based on the elapsed time of the first macro, it took  0.8984375 seconds while the refactored took 0.3242188 seconds. Hence, the percentage decrease is about 64%.

![VBA_Challenge_2017_Resouce](https://user-images.githubusercontent.com/62758795/202240116-1fd784c2-56b5-4158-8705-c2381a80192c.png)

![VBA_Challenge_2018_Resource](https://user-images.githubusercontent.com/62758795/202240201-5454f75b-17e4-4128-9898-8e09d41f5d90.png)


## *Summary*
Refactoring code comes with its advantages and disadvantages. When VBA codes are Refactored, the logic and design of the code advanced, the code is easier to understand, the code saves time when running and helps to find bugs. Additionally, refactoring code increases efficiency. Refactored code can also take up less storage, which can be very valuable. Conversely, refactored code can be a time-consuming process as it can be difficult to figure how exactly to make the code perform better. It might also affect further use of the code as features present in the unfactored code might be needed in the future.

When creating the script for Steve client’s, we used refactoring to decrease the time it takes to analyze the data. Most importantly for Steve, he needs his analyses to run concisely and promptly. Observe that the refactored script runs without looping through unneeded data points. One disadvantage we experienced in refactoring this code was the time it took to make this code run correctly, although the original script produced the same results. 

