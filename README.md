# vba-scripts
Various VBA scripts created to save time in the process of data analysis. They do the job they are designed for but I make no guarantees that they are structured in the most efficient or beautiful manner. Tell me if you encounter any errors but consider them Persian flaws. Free to use without asking non-commerical.
For questions email this address, after shifting each letter up one: zkdldkk ZS pbb CNS btmx CNS dct

## TabCreatorbyCategory

This script employs Autofilter on a user-selected column with multiple nominal categories, with each tab of the worksheet a particular category. For instance, if your dataset has a column called color, with categories red, green, and blue, this script will create three new tabs for each color. This will leave you with four tabs: the original sheet plus three called "red", "green", and "blue". 

You can run the script multiple times to quickly categorize a worksheet.

## NSC_Formatting

This script will format CO, DA, or SE data queries for upload to National Student Clearinghouse's StudentTracker service. To avoid repetitive dialogue windows, you will have to adjust the first few lines of code to your particular institution. Original file must have first name, middle name, last name, YYYYMMDD birth date and student ID (Requester Return Field) in Columns A to E IN THAT ORDER, with headings.
