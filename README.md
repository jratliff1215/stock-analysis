# Using VBA to Analyze Stock Performance
## Overview of Project 
Steve requested us to refactor our original code to allow for data expansion. The current code is limited to a specific year contained within a worksheet. Further, the original code presented with 12 iterations, with each ticker requiring a pass through of the data. The final refactors reduced this to a single pass through of all the data, utilizing output arrays to decrease total run time. 
## Results
The refactor of the code reduced the run time considerably for the VBA script. Both 2017 and 2018 required 0.125 seconds to complete with the refactored script, an approximate reduction of 80% from the original times for each. 
