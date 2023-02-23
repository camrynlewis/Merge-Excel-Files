# Merge multiple Excel files into one workbook with separate sheets

This repository may be useful for data analysis on large datasets in Excel.

Consider a case where you have multiple xlsx files in a folder & you want to merge them into a single file. Here, I used read_excel() and mutate() to store the file paths in a column. From there, I isolated the file names and then split the data into individual lists. lapply() was used to add each sheet to the final merged file.
