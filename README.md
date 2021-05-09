
# Stock Research Analysis

## Background
Steve, our client, has recently graduated with his finance degree and his parents have agreed to be his first clients.They are very passionate about green energy and the many different energy alternates available to replace fossil fuels. However, they have a narrow understanding of the different investible companies available that align with their investment philosophy and can help them with their investment goals. Steve has agreed to perform analyses on different green energy companies as he is concerned about having suitable diversification in the their portfolio. As such, a solution was prepared for Steve that allows him to perform these analyses on the desired green energy companies and provide feeback to his parents. At this time, Steve would like to revist this existing solution to include the entire stock market over a number of years.

## Purpose
Steve has been using a VBA coded solution that has allowed him to run analyses with many different stocks. This existing solution is reusable and allows him to reduce the chance of accidents and errors when performing reliable analyses.

In order to meet this new requirement we need to look at refactoring the VBA code to make it more robust and faster in order to execute the expanded dataset within a reasonable amount of time. We can also perform comparative testing on the old and any new refactored solution that is designed.

### Code Review and Analysis

After reviewing the existing VBA code the following changes were recommended for delivery:

- A new tickerIndex was created as an iterator variable to move down each row of the array tables for each ticker.
- Four new arrays were created with the goals of saving each ticker symbol, total Volume, starting prices and ending prices
  
During the collection part of the VBA code, beginning at (1a) - a ticker symbol is chosen, the tickerIndex is set to zero and the tickervolume array is initialized to zero. All rows pertaining to this particular ticker are iterated through by using the i loop variable. The volume is added, the starting price and ending price are captured and all values are stored into the appropriate array. Once the ticker changes, the tickerIndex increments by one and we continue down the dataset until the next ticker value changes and arry values are likewise captured again. By doing this, we are able to poplate each array with the proper values and only perform a single iteration of the entire dataset. Once all values are recorded we can then loop through the ticker symbol one at a time, perform calculations and output the data from the arrays to the worksheet then apply formatting as required.

In our particular tests for speed optimization we observe the following speed improvement:

_**Original Solution**_

![Original 2017 Results](Resources/VBA_Challenge_old_2017.png)

![Original 2018 Results](Resources/VBA_Challenge_old_2018.png)

_**Refactored Solution**_

![Refactored 2017 Results](Resources/VBA_Challenge_2017.png)

![Refactored 2018 Results](Resources/VBA_Challenge_2018.png)

## Results 
As can be seen in the results above, the refactored code ran approximately 453% faster than the original code. It appears that the refactored code has met the objective of being more robust and faster as per the original plan.


## Summary

Given the results obtained, we are confident that Steve will be impressed with the performance improvement and his desire to run a much larger dataset using the new VBA solution for his stock analyses. From Steve's perspective, the results are exactly the same only faster so it meets his goal to be able to analyse larger datasets in a shorter timeframe.


- There is a detailed statement on the advantages and disadvantages of refactoring code in general?
As can be seen in this solution refactoring can make the software easier to understand and read. We were also able to improve the design of the VBA code and improve its operational structure while keeping the results exactly the same. Alternatively, refactoring code takes time and the client already has an expectation of how the code should operate. If we need to change this for any reason then we then need to add training time to the deliverable.

- There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script

By using arrays we were able to use more of the faster computer memory to store and retrieve values. In the origoinal design we used a nested loop that iterated through the entire dataset for each ticker symbol. In the new solution we iterated through the dataset exactly one time while capturing all the data for each ticker symbol.

