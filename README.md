# ExcelSigDigits

VBA macro to format cells with given number of significant decimal digits.  
It sets only the format code and doesn't change the cell value.  
It respects and maintains the percent format.  
It changes only the appearance of the fractional digits, it doesn't round anything before the decimal separator.  
It shows trailing zeroes only if they are significant / reliable. They are removed for integer numbers.  

*improved from http://www.spreadsheet-validierung.de/excel-signifikante-stellen*

## Usage
- install in excel as macro
- [optional] set up keyboard shortcut
- mark cells
- execute macro
