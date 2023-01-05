# stock-analysis

## Project Overview
* In this project we will refactor the code generated for our stock analysis. This will allow us to expand our stock market data to include other years and companies without compromising the speed at which the code is run.
* Specifically we will be analyzing the improvements to the code, comparing the execution times of the original and refactored script, and lastly looking at stock performances for 2017 versus 2018.

## Results

### Original Code Analysis

When analyzing the differences between the original and refactored script we can see a lot of similarities, but there are also a few differences worth noting which can be seen in the code below:

![VBA_Challenge_Original_Code](https://user-images.githubusercontent.com/120291854/210708137-05108563-0162-4f20-b63e-f83efa875d58.png)

Here in the original script, we looped through all of the tickers and then looped through all of the rows in a nested For loop. In the outer loop, we told VBA to loop through each ticker and we set our “totalVolume” variable equal to 0 so that it would be initialized prior to calculating our Total daily Volume for each ticker. In the inner loop we told VBA to go through each row to find the total volume and repeat this for starting price and ending price. For starting price and ending price, our script uses And conditionals that indicate when these values should be recorded. 

### Refactored Code Analysis

![VBA_Challenge_Refactored_Code](https://user-images.githubusercontent.com/120291854/210708372-1a7758e4-26be-4e37-860c-f785f7f572ef.png)

For the refactored script we wanted to improve the codes efficiency by only going through the data once. To achieve this, we started by creating a ticker index as well as three additional arrays. These new variables were the driving force for refactoring the code. The “tickerIndex” variable allows us to access the correct index across our four arrays when we start our second For loop. In the next step we created a For loop for “tickerVolumes”, which we set equal to 0. This For loop allows the total daily volume for each ticker to start at 0. Our second For Loop has us increase the volume for our current ticker (to find total daily volume for that ticker), check if the current row is the first row (to obtain the starting price for that ticker) and finally check if the current row is the last row with the selected ticker and increase the tickerIndex if the values don’t match (which gives us our ending price). The final step simply outputs the code in a table and is nearly identical to the original script. 

### Runtime Analysis

When looking at the execution times we can see a massive difference between the elapsed run time for our original script compared to the refactored script. Based on the images below we can see that the our refactored code runs in about 4/100ths of a second while our original script takes over half a second to run. This means our refactored code is running about 10x faster!

#### Original Code Run Times
![Green_Stocks_2017](https://user-images.githubusercontent.com/120291854/210708832-b3492bfa-4313-45c5-8a6e-8b060e3c0e32.png)
![Green_Sotcks_2018](https://user-images.githubusercontent.com/120291854/210708841-57f68360-98e5-4c3c-aeed-8df747cd35e1.png)

#### Refactored Code Run Times
![VBA_Challenge_2017](https://user-images.githubusercontent.com/120291854/210708797-cff246cf-49d5-4eac-91c7-2bb5838a62fa.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/120291854/210708807-cad5c717-9a9e-4e1b-98fe-f48151f03166.png)

### Stock Performance Analysis

The code we have created and refactored allows us to analyze stock performances for both 2017 and 2018. In the tables below we can clearly see that 2017 was a much better year for nearly all these companies. Almost every company’s stock yielded positive returns in 2017, while yielding negative returns in 2018. However, the amount of trading for each stock (Total Daily Volume) tends to vary, with no clear trend between the two years. 

![2017 Stock Performance](https://user-images.githubusercontent.com/120291854/210709023-38ac0155-d623-47fc-8e65-c80eae49fa4b.png)
![2018 Stock Performance](https://user-images.githubusercontent.com/120291854/210709045-6af5104a-7e21-43eb-b0a2-3d0e567bf16b.png)

## Summary Statements

* In conclusion, refactored code has clear advantages, such as increasing the efficiency of the code (allowing it to run faster), making the code easier to comprehend and review (by detecting code smells and anti-patterns), and even finding potential bugs. However, does pose potential drawbacks like having to take time to refactor the code and potentially introducing new bugs into the code. 
* For our code we had to spend additional time refactoring and this process was not without its fair share of errors that needed to be debugged. Nevertheless, taking the time out to refactor our Stock Analysis code resulted in a 10-fold increase in run speed as well as making the code easier to follow by eliminating the nested For loop as well as the And conditionals.
