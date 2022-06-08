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


## CountFilesInFolder
todo...
