# Excel Related Tools
Here are some useful VBA macros and formulas that are handy for ad-hoc analysis

[Formatting Macros] (./format_macros)

A collection of macros to format excel files.

Instructions:

1. Copy this file to the Excel xlstart directory [(see related documentation here)] (https://support.office.com/en-in/article/Customize-how-Excel-starts-6509b9af-2cc8-4fb6-9ef5-cf5f1d292c19)
2. When Excel first launches, you may require to enable macros, depending on your security setting
3. The worksheet that now automatically loads maybe edited / printed as reference.
4. I like to hide the Excel worksheet. Now Excel will start to a blank page. Press "Ctrl+N" to start a new workbook. Unhide the worksheet and it will come back.
5. Feel free to edit the macro as necessary, remember to also update the worksheet.
6. I never got turning off autocalculate to work.

[Compare Sheet Macro] (./Compare Macro)

A macro to comapre the differences between two excel sheets (origionally to scan for changes in a massive HR csv file). 

Instructions:

1. Move the two sheets you want to compare into a single workbook
2. Copy the macro into the same workbook
3. Both sheets need to have a primarily index column that is common between the two and common column headings. In this case, it was employee ID.
4. Run the macro 
5. The macro will ask for the name of the more recent sheet and the older, prior sheet
6. The macro will ask for the column NUMBER (not letter) that is the primary key
7. The macro should iterate through each of the rows and columns and highlight on the more recent sheet what cells have changed.
8. If a new column or row has been inserted, that entire column or row will be highlighted. 

It would have been implemented better. For up to 400+ employees, run time was under 5 seconds. Compared to manually scanning the worksheets or using bunch of formulas, this was sufficiently good.
