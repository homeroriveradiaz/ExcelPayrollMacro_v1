# ExcelPayrollMacro_v1

The purpose of this code is to demonstrate features of Excel's VBA, and it was created for a video class; The code is not efficient and it was never meant to be.

It can be re-used however, and could help as a reference or base for another project.

For reference, the data used in the sample Excel files is available at https://www.thespreadsheetguru.com/sample-data/, however I did a few adjustments to that file, thus using the file from https://www.thespreadsheetguru.com/sample-data/ won't work with this Macro; You'll have to use the files in this repo.

##Pre-requisites:
-MS Excel, 2007 and later most recommended, but could work with earlier versions.

##Instructions to use:
1. Create folder C:\temp\Excel examples\ and save all files in this repo to that folder.
You can use another folder by changing lines of code 59 and 68 (basically, changing any references to "C:\temp\Excel examples\").
2. Open the file "Employee Sample Data.xlsx", create a Macro, edit it, and paste MacrosCode.vb in its place.
3. Make sure no other Excel file is open
4. Execute macro "FormatPayrollData".
