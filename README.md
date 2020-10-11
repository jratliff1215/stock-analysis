# Using VBA to Analyze Stock Performance
## Overview of Project 

Steve requested us to refactor our original code to allow for data expansion. The current code is limited to a specific year contained within a worksheet. Further, the original code presented with 12 iterations, with each ticker requiring a pass through of the data. The final refactors reduced this to a single pass through of all the data, utilizing output arrays to decrease total run time. 
## Results

The refactor of the code reduced the run time considerably for the VBA script. Both 2017 and 2018 required 0.125 seconds to complete with the refactored script, an approximate reduction of 80% from the original times for each. 

The below figures show to change in run time for the 2017 Stock Data, with the original run time followed by the refactored run time 
<img src="https://github.com/jratliff1215/stock-analysis/blob/main/Resources/2017%20Original.PNG" width="500" height="300">
<img src="https://github.com/jratliff1215/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG" width="500" height="300">


The below figures show to change in run time for the 2018 Stock Data, with the original run time followed by the refactored run time 
<img src="https://github.com/jratliff1215/stock-analysis/blob/main/Resources/2018%20Original.PNG" width="500" height="300">
<img src="https://github.com/jratliff1215/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG" width="500" height="300">


## Summary

*What are advantages or disadvantages to refactoring the code?*

Most scripts that are written in VBA are derivatives or enhancements of prior used code. However, even if a code has been used by many, it may not be the most efficient or extensive code. Each section of code may be improved upon to enhance the overall functionality of the code. Additionally, the continual refactoring of code increases the toolkit and knowledge of the developer. Refactoring allows the developer to create the best code for that application. 

*How do these pros and cons apply to refactoring the original VBA script?*

The refactored code is stronger and more efficient than the original code. However, this script requires the data to be in alphabetical and chronological order. This requires to original data to be in specific order for the data to correctly compile. An addition to the script may include a sort function, which sorts the data correctly prior to running the script. 
