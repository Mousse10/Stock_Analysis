# Stock Analysis with Excel VBA

## Overview of Project

### Purpose
We previously helped Steve by creating a workbook that enabled him to easily do yearly stock analysis by the click of a button. He is now able analyze an entire dataset in a matter of seconds. He now wants to be able to perform the same analysis but on a larger scale. Although the code previously created works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.
This said, we need to edit, or refactor the solution code previously created to loop through all the data one time in order to collect the same information. This would make our code more efficient by taking fewer steps, using less memory, or even improving the logic of the code to make it easier for future users to read.

## Results
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Upon further review and following the provided instructions, the below changes were made to our previous script 

![](Resources/All%20Stocks%20Analysis%20Refactored.PNG)

![](Resources/All%20Stocks%20Analysis%20Refactored2.PNG)

Once those adjustments were made, I was able to rerun the macro and got the following results for the years 2017 and 2018 which matched our previous results from the initial script. 

![](Resources/2017%20-%20Results.PNG) ![](Resources/2018%20-%20Results.PNG)

When I run the initial script, the timer results are longer in comparison to the timer when the new script is run. 



## Summary 
Some advantages of refactoring code include extensible code and maintainability. Code Refactoring makes the code more extensible for adding many other functions to it. It also helps in increasing both the flexibility and capability of code. After refactoring, the code is fresher, easier to understand or read, less complex and easier to maintain. Some disadvantages of refactoring code include the facts that it can be time consuming while also increasing the chances of mistakes. You may have no idea how much time it may take to complete the process. It may also land you into a situation where you have no idea where to go. In case it went wrong, you would have to waste much more time in solving the problem and there are probable chances that it may go wrong due to complexity of the code. Code Refactoring is a wise method of extending the code and making it more capable and in the same way it has some disadvantages. Refactoring the original VBA script was indeed time consuming due to the adjustments that were being made to the code.  The code then had to be tested in order to figure out what the problems or issues were in the end, refactoring allowed us to improve our code in regards to execution time and supporting a higher number of inputs and outputs.
