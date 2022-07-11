# Module 2 Challenge 
## Refactor VBA Code and Measure Performance 
Overview: 
For this analysis, a collection of twelve stocks was considered in order to show understanding of VBA. The data was formatted to show which stocks performed well and which did not using a flexible macro. The original code that involved nested loops was refactored to include an array for both the 2017 and the 2018 sheets in the workbook.

## Results 
The performance of stocks in 2017 originally ran in .90 seconds and when it was refactored, it ran in .77 seconds. See image (scroll to right on screenshot, sorry!)  https://github.com/Lisa-Floading/stock-analysis-/blob/main/VBA_2017_Challenge.png.png 

The performance of stocks in 2018 originally ran in .91 seconds and when it was refactored, it ran in .79 seconds. See image https://github.com/Lisa-Floading/stock-analysis-/blob/main/VBA_Challenge_2018.png.png


## Summary 
The advantages of refactoring code are considerable when there is a massive amount of data. The developer must ask if refactoring code will make the end user experience more efficient because it runs faster. The disadvantage can be summed up with the old adage, "perfect is the enemy of good." If making the code more efficient also makes it more complicated, this may create problems down the road. However, if refactoring the code provides a solution that can be used in different circumstances than the one it was originally created for, then the code is more valuable. 

This relates to my experience refactoring the original VBA script in that this is a relatively small dataset and refactoring the code didn't create that much of an impact. The important element to consider is how this refactored code is ultimately more flexible than the original, making this a valuable experience in learning to code. 

