ExcelReportBuilder

A Python script that takes student grade data and organizes it into a clean, formatted Excel workbook, complete with a raw data sheet, a summary sheet with averages and pass/fail status, and an error log. Built with OOP principles using openpyxl.

The script produces a single .xlsx file with three sheets:

Sheet 1 - Raw Data. Stores each student's full record

Sheet 2 - Summary. Automatically calculates the average grade per student and assigns a pass/fail status

Sheet 3 - Error log. Any row that fails to process (e.g. wrong date format, missing data) gets logged here automatically with the row content and the exact error message. This means the script never silently fails, you always know what went wrong and why

Key Features

Object-Oriented Design: The entire logic is wrapped in a clean ExcelReportBuilder class, making it easy to reuse, extend, or import into other projects
Error handling on every row: If one student's data is malformed, the rest still process normally. Nothing crashes the whole script
Automatic error logging: Bad rows are written to a dedicated error sheet instead of just printing and disappearing
Flexible and configurable: Sheet names, output filename, and data are all passed in at creation, so nothing is hardcoded
Date parsing: Dates are stored as proper Excel date objects (not plain strings), so Excel can sort and filter them correctly
Grade filtering: The summary sheet only averages actual numeric grades, safely ignoring any non-numeric values
Single entry point: main() runs everything in the correct order cleanly

How to run it 
1.Install the dependency
pip install openpyxl
2. Clone the repo and run the script
python newwe.py
Note: I really don't know why I picked that name
3. The output file will be saved as your_title.xlsx in the same directory.

How to use it with your own data: 

Replace the list_of_students list with your own data. Each row should follow this format:
[Name', 'YYYY-MM-DD', grade1, grade2, grade3, ...]
You can also customize the sheet names and output filename:
pythonreport = ExcelReportBuilder(
    worksheet="Raw Data",
    second_name="Summary",
    error_worksheet="Errors",
    data=your_data,
    title="MyReport"  # output file will be MyReport.xlsx
)
report.main()


Dependencies

openpyxl
Python 3.10+


Notes
My first time sharing code online. Right now, diving into openpyxl along with object-oriented programming using Python. If something feels off or could be better, feel free to point it out. Growth is the goal. I'm thinking of making a program which grabs info from websites, then drops it neatly into an Excel file soon. Maybe someone out there will get some use from it.
