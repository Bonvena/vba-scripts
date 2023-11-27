# vba-scripts
Various VBA scripts created to save time in the process of data analysis. They do the job they are designed for but no guarantees are made that they are structured in the most efficient or beautiful manner. Please inform me if you encounter any errors. All rights reserved but free to use non-commerical.

## TabCreatorbyCategory

This script employs Autofilter on a user-selected column with multiple nominal categories, with each tab of the worksheet a particular category. For instance, if your dataset has a column called color, with categories red, green, and blue, this script will create three new tabs for each color. This will leave you with four tabs: the original sheet plus three called "red", "green", and "blue". 

You can run the script multiple times to quickly categorize a worksheet, such as by department, region, or company.

## NSC_Formatting

This script will format CO, DA, or SE data queries for upload to National Student Clearinghouse's StudentTracker service. To avoid repetitive dialogue windows, you will have to adjust the first few lines of code to your particular institution. Original file must have first name, middle name, last name, YYYYMMDD birth date and student ID (Requester Return Field) in Columns A to E IN THAT ORDER, with institution-specific variable names/headings in the first row. (The code assumes the institution saves suffixes to the end of the last name field, not as its own separate field. If not, the suffix column would have to be merged afterwards).

## Perkins_CTEA_CLNA_2P1_3P1_7A_7B_Table_Generator

This script will calculate and format certain tables required by the New York State Education Department for Perkins CLNA (Comprehensive Local Needs Assessment) for 2023-2024, namely tables 2p1 (Earned Recognized Postsecondary Credential), 3p1 (Non-Traditional Program Concentration), 7a (Enrollment), and 7b (Completion). It will also require the 2020 Nontraditional Occupations Crosswalk in Excel format, found here: https://cte.ed.gov/accountability/linking-data 
