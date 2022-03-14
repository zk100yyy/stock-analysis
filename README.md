# Stock Analysis - VBA Challenge
## Project Overview

The workbook we had created that goes through the data and outputs stock analysis for each ticker in the code. The analysis shows the total volume traded as well as the annual return. The information is presented in a formatted table. Now we're trying to see if we can make the code run faster and more efficient. 


## Results of the Analysis
#### After we run the code for 2017 and 20018 we get the below data presented in a formatted table:
### 2017 Stock Analysis:
![Stock Analysis 2017](https://github.com/zk100yyy/stock-analysis/blob/main/Resources/Stock_Analysis_2017.png?raw=true)
### 2018 Stock Analysis:
![Stock Analysis 2018](https://github.com/zk100yyy/stock-analysis/blob/main/Resources/Stock_Analysis_2018.png?raw=true)

## Runtime comparison Original vs Refactored
### 2017 Original Runtime:
![2017 Orginal Runtime](/Resources/VBA_Challenge_2017_original.png)
### 2017 Refactored Runtime:
![2017 Refactored Runtime](/Resources/VBA_Challenge_2017.png)
### 2018 Original Runtime:
![2018 Orginal Runtime](/Resources/VBA_Challenge_2018_original.png)
### 2018 Refactored Runtime:
![2018 Refactored Runtime](/Resources/VBA_Challenge_2018.png)
<br>
#### We can clearly see a significant improvement in runtime for both 2017 and 2018 Analysist. For 2017 runtime improved by more than 3 folds and for 2018 runtime improved by more than 4 folds. 

#### The main reason for this improved efficiency and faster runtime is the fact that our code runs through the data only once. In the original code we had a nested for loop where in the first for loop (in green below) loops through the 12 tickers, and then for each ticker we have a second for loop (in red below) that loops through our data from beginning to end. In essence we're looping through the data 12 times!

### Nested for loop in Original Code:
![Original Code](/Resources/Original_Code.png)

#### In contrast, our refactored code only has to loop through the data once and while doing so moves from one ticker to the next, collecting data or each ticker and storing it in the corresponding array indices. See loop in red below and note that there isn't a nested for loop inside of it:

### Singlefor loop in Refactored Code:
![Refactored Code](/Resources/Refactored_Code.png)

##Summary

###Advantages of refactoring code in general:
* Code runs faster and more efficient
* Can extend the code to run more than the originally intended set of analysis
* Code is cleaner

<br>

###Disadvantages of refactoring code in general:
* Requires more complexity and creativity
* Might take more storage space
* Might be harder to debug if it's too complex

<br>

###Advantages of refactoring code in VBA_Challenge:
* Code ran a lot faster: 3-4 times
* We can now use it to analyze more than 12 stocks
* Code is cleaner

<br>

###Disadvantages of refactoring code in  VBA_Challenge:
* Requires more data structures like arrays
* Had to hold more data in storage for the time the code was running 
* A bit more complex than original code
