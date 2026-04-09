ExcelReportBuilder
This Python based app that will effectively structure student grade information inside an excel workbook that is professionally formatted. It provides 1 raw data worksheet and 1 summary worksheet containing summary grade averages from the student groups and a dedicated worksheet for logging errors in the process. Developed using modern Object-Oriented Programming principles and `openpyxl`.

This script will produce 1 `.xlsx` file containing 3 separate worksheets / tabs.

**Sheet 1 – Main Report**: Contains every student and his/her complete record sorted in the correct data types.

**Sheet 2 – Student Summary**: Automatically calculates the average grade per student and assigns a PASS / FAIL status (based on a configurable average threshold) to each student.

**Sheet 3 – Error Log**: All rows that could not be processed (i.e. bad date format or empty cell) will be logged along with the row content and technical details of the error. The script will never crash; it will just log the error and move on.

**Key features**:

-Fluent Interface (Method Chaining): The class supports a fluent interface (method chaining) enabling initialization, processing and saving of the report in one coherent block of code.

-Encapsulation: All workbook logic and data containers/arrays are private protecting the "engine" from outside tampering.

-Centralized Row Processing: The use of a helper method for handling names, dates & grades ensures uniformity when processing data on each of the three worksheets.

-Configurable Pass Threshold: You can set your passing average directly inside the `ExcelReportBuilder` upon instance initialization. (i.e. Pass=60, Pass=70, etc.). into openpyxl along with object-oriented programming using Python. If something feels off or could be better, feel free to point it out. Growth is the goal. I'm thinking of making a program which grabs info from websites, then drops it neatly into an Excel file soon. Maybe someone out there will get some use from it.

Install the dependency
pip install openpyxl

Clone the repo and run the script
python your_script_name.py

The output file will be saved in the same directory.

How to use it with your own data
You can now use the new fluent syntax to build your report:

Python
list_of_students = [
    ['Jane', '2024-06-20', 40, 70, 100, 56],
    ['Bob', '2026-06-20', 93, 56, 74, 89],
]

# Configure, process, and save in one go
(ExcelReportBuilder(pass_threshold=60)
    .add_data(list_of_students, headers=['Name', 'Date', 'T1', 'T2', 'T3', 'Exam'])
    .create_main_report("Class Records")
    .create_data_records()
    .create_summary_sheet("Final Grades")
    .create_summary_report()
    .make_error_report()
    .save_file("student_summaries.xlsx"))
Dependencies
openpyxl

Python 3.11+ (uses typing.Self)

Notes
My first time sharing code online. Right now, diving into openpyxl along with object-oriented programming using Python. I've updated the script to be more modular and flexible based on feedback; moving towards a more "fluent" way of writing the code. If something feels off or could be better, feel free to point it out. Growth is the goal. I'm thinking of making a program which grabs info from websites, then drops it neatly into an Excel file soon. Maybe someone out there will get some use from it.
