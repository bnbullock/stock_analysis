
# Stock Research Analysis

## Background
Steve, our client, has recently graduated with his finance degree and his parents have agreed to be his first clients. They are very passionate about green energy and the many different energy alternates available to replace fossil fuels. However, they have a narrow understanding of the different investible companies available that align with their investment philosophy and that can help them meet their investment goals. Steve has agreed to perform analyses on different green energy companies as he is concerned about having suitable diversification in their "green energy" portfolio. As such, a solution was prepared for Steve that allows him to perform these analyses on the desired green energy companies and provide feeback to his parents. At this time, Steve would like to revisit this existing solution to include the entire stock market over a much longer duration.

## Purpose
Steve has been using a VBA coded solution that has allowed him to run analyses with many different alternative energy stocks. This existing solution is reusable and allows him to reduce the chance of accidents and errors when performing reliable analyses for his parents.

In order to meet the new requirement focusing on all investible stocks we need to look at refactoring the VBA code to make it more robust and faster in order to execute the expanded dataset. We can also perform comparative testing on the old and any new refactored solutions. The goal is to determine if the new refactored design is more or less efficient and to provide evidence in either case.

### Code Review and Analysis

After reviewing the existing VBA code the following changes were recommended for delivery:

- A new tickerIndex was created as an iterator variable to move down each row of the array tables for each ticker symbol.
- Four new arrays were created with the goals of saving each discrete ticker symbol, total Volume, starting prices and ending prices.
  
During the row iteration part of the VBA code, beginning at comment marker (1a) - a ticker symbol is chosen, the tickerIndex is set to zero and the tickervolume array is initialized to zero. All rows pertaining to this particular ticker are iterated through by using the i loop variable. The volume is added, the starting price and ending price are captured and all values are stored into the appropriate array. Once the ticker changes, the tickerIndex increments by one and we continue down the dataset until the next ticker value changes and arry values are likewise captured again. By continuing this path, we are able to poplate each array with the proper values for each ticker and only perform a single iteration of the entire dataset. Once all values are recorded we can then loop through the ticker symbol array one at a time, output data to the worksheet and perform calculations as may be needed. Finally, we can apply formatting to the worksheet as required.

In our particular tests for optimization we observe the following speed improvements:

_**Original VBA Solution**_

![Original 2017 Results](Resources/VBA_Challenge_old_2017.png)

![Original 2018 Results](Resources/VBA_Challenge_old_2018.png)

_**Refactored VBA Solution**_

![Refactored 2017 Results](Resources/VBA_Challenge_2017.png)

![Refactored 2018 Results](Resources/VBA_Challenge_2018.png)

## Results 
As can be seen in the results above, the refactored code ran approximately 78% faster than the original code. It appears that the refactored code has met the objective of being more robust and faster as per the original expectation.

## Summary
Given the results obtained, we are confident that Steve will be impressed with the performance improvement and his desire to run a much larger dataset using the new VBA solution for his all stock analyses. From Steve's perspective, the results are exactly the same with a much better execution speed so it meets his goal to be able to analyse larger datasets without long delays. Steve will now be able to provide his parents with much better information and will allow them to make informed philosophical investment decisions to meet their investment goals.

### Advantages/Disadvantages of Refactored Code 
As can be seen in this solution, refactoring can make the software easier to understand and read. We were also able to improve the design of the VBA code and improve its operational structure while keeping the results exactly the same. Alternatively, refactoring our VBA code took time to design, implement and test. We also have our clients expectation of how the code should operate and what the results should look like. If we need to change the execution process for any reason then we then need to add training time to the deliverable. We are also continuing to run the VBA code in Excel, so any issues with the overall architecture or interface will continue to be present in the refactored VBA code as well.

### Advantages/Disadvantages of the Original vs Refactored VBA script
By using arrays we were able to use more of the faster computer memory to store and retrieve values. In the original design we used a nested loop that iterated through the entire dataset for each ticker symbol and we also switched back and forth from 2 different worsheets to perfrom the analysis. In the new solution we iterated through the dataset exactly one time while capturing all the data for each ticker symbol. There was also a great advantage in using arrays to store values in lists in the computer memory it made it easier to retrieve and make calculations for output.