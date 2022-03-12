# VBA Challenge
##Overview
Steve approached us about developing a macro for him to use to analyze stocks that he thought his parents should invest in.  In order to manage the necessary calculations, we decided to code this project in VBA to utilize Excel's spreadsheet format to help analyze the data and display our results.  After a review of the code, we were able to refactor it for more optimal performance.  

##Results
One reason for refactoring the code was to make it run faster in real time.  I was able to accomplish this by running through the data once per cycle instead of the code running through the data multiple times by way of nested loops.  
###2017
In the orignal code, when we ran the macro for 2017, the runtime was 0.289 seconds.  While this wasn't too long for this number of stocks, it would become time consuming if we tried to analyze the entire market.  By refactoring the code, I was able to reduce that runtime significantly.  The refactored code runs in about a third of the time.
<img width="257" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/99457275/157998342-a268a1a3-238c-4af3-b43c-a6c45308deb1.png">
###2018
For 2018, the time for running the original code was, similarly to 2017, 0.289 seconds.  The refactored code also saw a drop in runtime.  
<img width="265" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/99457275/157998521-94b4cc9c-5f74-456d-af95-8388494ad791.png">
Once again, the refactored code runs in about a third of the time.  It is also important to note that the original code didn't include the formatting that the refactored code has, since it had been separated out into a separate subroutine.  This time savings can be extrapolated to larger datasets, as well.

##Summary
Refactoring code can be very beneficial.  The results show a clear time advantage for the refactored code.  Other benefits can include reduced storage requirements and easier navigation for future coders who will use or view your code.  The downside is the labor investment required upfront to "fix" something that isn't broken.  However, most of the time, the pros outweigh the cons.  The timer test in the refactored code allows us to see the pros very clearly, as this cleaner code is clearly more efficient.  The con of spending more time on this code is doesn't necessarily apply as it didn't involve any business resources, just my time.  In summation, refactoring code is a good thing, since cleaner code can allow it to become useful for a wider variety of uses.
