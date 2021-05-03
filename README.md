# Written Analysis of the Results
## Overview of Project
The VBA Challenge is ultimately about refactoring previously written code inside the Visual Basic application in Microsoft Excel. Refactoring runs through the data at one time without adding anything new but rather makes it more efficient by improving the logic within the code and using less memory. The code collected data from two datasets separated by year containing information of thousands of stocks and narrowed it down to twelve individual stocks with their total daily volume and returns. 

### Refactoring the Code
The refactoring of the code was accomplished a few different ways, first by using the ReDim function as opposed to Dim in order to resize the array. The tickerIndex was also resized in order to match each row’s ticker. The refactored code also looped through each row starting from zero in order to find the volume, starting prices, and ending prices of each year. Overall these changes came from the nesting order of the forloops using the tickerIndex varible.  
  
## Results

### Analysis 
The first and second PNG’s shows the original code running stock information for 2017 and 2018 as well as how fast the code ran. It marks the difference between the two years in the returns row with a positive return being highlighted in green and negative returns being highlighted in red. The original Code for 2017 ran in 0.6015625 seconds while the original code in 2018 ran for 0.6796875 seconds. The third and fourth PNG’s shows how fast the code ran for the refactored code with a notable increase in efficiency for both. The refactored code for 2017 ran in 0.0625 seconds while the refactored code for 2018 ran in 0.0625 seconds as well. 

![2017Original](https://user-images.githubusercontent.com/82983000/116842030-1714e480-aba9-11eb-8adc-3a633e303a63.png)
![2018Original](https://user-images.githubusercontent.com/82983000/116842031-1a0fd500-aba9-11eb-803d-69f92db61325.png)
![2017Refractured](https://user-images.githubusercontent.com/82983000/116842035-1c722f00-aba9-11eb-97f0-e9b977f82252.png)
![2018Refractured](https://user-images.githubusercontent.com/82983000/116842037-1da35c00-aba9-11eb-84c0-5f1ad537abbf.png)



  
  
## Summary

### Advantages and Disadvantages of refactoring code



### Advantages and Disadvantages of the original and refactored VBA script


