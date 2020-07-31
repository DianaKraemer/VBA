VBA code I wrote to make my life easier. 

In public safety, there are many radio channels and options. Please take a look at zone.PNG for an example of part of a document that we provide to agencies that showcases their chosen channel options. 
Making this document is a tedious process. After making a number of unnecessary mistakes, I wanted to see if I could automate the process. I wanted to save time and eliminate as many errors as possible. (Please note that there are some things that may not be able to be automated, namely the individual agency's channel names because those 
depend on what the agencies naming conventions are and any customizations that may be needed. This is something I'm hoping to address in the future.)
          
This document automates a numbering system, coloring/formatting of the zone blocks and coloring/formatting of set channels, such as federal interoperability channels and other known constants for my area. 

#Instructions
Download the most recent Excel file. There is a column of usable data already ready to go. **Make sure to enable macros or the document won't work!** 
Push the button and watch the magic!
If you'd like to see further funtionality, the Exel file "Selections" has a few more options. This allows you to see what I mean when I say that there are some cells that I may not be able to automate. **Must have applicable data in A1, B1, or C1. Currently data placed into D1 does not work correctly**


**Update uploaded 7/27/2020:** Picking the project up again after a few months. The projects is moving towards the goal of handling more specific cases, such as the ones claimed above to not be automated due to individual agences naming conventions. After looking through previously created layouts from my coworkers, I've found there aren't many naming standards. This allows me to set my own standards, furthuring the functionality of my application. 
The difficulty in implementing these is that I am trying to eliminate pages of If/ElseIf checks/handling conditions. 



**Update uploaded 7/31/2020:** 
Added function to execute the Zone name with any length name (but not formats that aren't "Zone + letter + name"), IFERNs, VTACs. 
Added a "cellChecker" function that correctly executes the code if the information is placed in A1, B1 or C1. D1 coming. 
Finally got the "Marc #" working correctly. A lot of difficulty here. I realized that previous documents I'm working from have these cells as "text" format which, once the code started to execute, was changed to "custom" and not catching in further checks. Fixed this by altering my reformatting function and placeing it first in the execution hierarchy. 
