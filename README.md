# VBA_Helper_Methods
A collection of useful subs and functions for Excel VBA

## Log
Push text logs to the VBA Editor's Immediate Window with an option to clear the previous entries
wip...

## OptimizeVBA
* This subroutine allows you to enable an optimized VBA calculation mode with a single call before executing your VBA code and disabling it all with one call after that
* It will also disable the page breaks on all sheets. To disable this functionality, delete or comment the following lines:
```vba
  For Each oWS In ActiveWorkbook.Worksheets
    oWS.DisplayPageBreaks = False
  Next oWS
```
* An explanation on why the page breaks should be disabled:

... *"One bad thing about page breaks when you have VBA code running is they want to recalculate the "breaks" whenever a change is made to the spreadsheet.  You could image the time consumption that might take place if you are running VBA code that is deleting or adding thousands of rows to a spreadsheet.  Because Page Breaks need to constantly recalculate, it is good to play it safe and shut them off while your code is being executed"* Chris Newman :: TheSpreadsheetGuru


## CountFilesInFolder
todo...
