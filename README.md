# SigDigits

Macro to format cells with given number of significant decimal digits.  
It sets only the format code and doesn't change the cell value.  
It respects and maintains the percent format, the € and the $ sign.  
It changes only the appearance of the fractional digits, it doesn't round anything before the decimal separator.  
It shows trailing zeroes only if they are significant / reliable and only up to the amount where they are result from rounding. Thus they are not shown for integer numbers for example.  
It works with split selection of cells.  
It has a safety timeout of 5 seconds. If it it gets reached, and not all Your values have been calcualted, make a smaller selection. Happens most of the time, if You marked a whole column. In this case, most of the time it will have calulated all Your occuopied fields in this column, so no need to make a smaller selection an recalculate.

*improved from http://www.spreadsheet-validierung.de/excel-signifikante-stellen/*

## Usage
- install in MS Excel or LibreOffice Calc as macro, for instance user-wide inside PERSONAL.XLSB (record empty macro to produce PERSONAL.XLSB)
- [optional] set up keyboard shortcut
- mark cells
- execute macro
