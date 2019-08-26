*************************
Table_pop program
current version 0.1.6
*************************
Use excel output to fill in table shells stored in a word document
Produce a new word document with filled in table shells
Formatting is retained from the table shell, so set all formatting in table shell
Need an excel file for each table shell in the word document


*************************
Packages
*************************
argparse
datetime
docx (Document)
os
pandas 


*************************
User provided arguments
(in this order)
*************************
Word table shell location
Word table shell filename
The row containing the first cell to populate in the word table shells
The column containing the first cell to populate in the word table shells
Excel file(s) location
Comma separated list of excel files (no white space between files)
The row containing the first cell to pull from the xlsx table
The column containing the first cell to pull from the xlsx table
Filled shells save location
Filled shells save filename


*************************
Noteworthy drawbacks
*************************
Word documents with multiple shells must all begin at the same cell
Same goes for multiple xlsx documents used for multiple shells
Will not error out if xlsx table is dimenionally smaller than word table
will error out if more word tables than excel docs provided
It is user responsibility to ensure data in excel is in the same order as the word document it is intended to fill


*************************
Future enhancements:
*************************
make enforce row order integrity argument live
comprehensive log(s)
Error handling
Create a package for distribution including a virtual environment
work inside of a word document with full text, images, and other tables
gui
update approach to use templating with jinja (avoids potential for excel/table shell misalignment by rendering the table directly from excel)
