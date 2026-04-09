from openpyxl import Workbook
from datetime import datetime
from typing import Self


class ExcelReportBuilder: 
    def __init__(self, pass_threshold: int = 70) -> None:
        """Initialize objects we will both use and potentially not use later on"""
        self.__workbook = Workbook()
        
        self.__main_sheet = None
        self.__summary_sheet = None
        self.__error_sheet = None
        
        
        self.__data = None           
        self.__error_log = []   
        self.__headers = None  

        self.__pass_threshold = pass_threshold #For creating summary report later
        
        #Remove active sheet for cleaner way of adding the main report later
        default_sheet = self.__workbook.active
        self.__workbook.remove(default_sheet)
    
    def add_data(self, data:list, headers: list | None = None) -> Self: 
        """Store row data and optional headers for later processing."""
        self.__data = data #General data such as names, dates, grades, etc...
        self.__headers = headers # The headers, which include [Names, Dates, Grade 1, ...]
        return self #Returns the instance itself to allow for method chaining and thus cleaner and more flexible code
        

    def _parse_row(self, row: list[str | int | float]) -> dict | None:
        """Helper function to initialize names, dates and grades, and combat errors cleanly"""
        try:
            name: str = row[0]
            date = datetime.strptime(row[1], '%Y-%m-%d')
            grades: list = [g for g in row[2:] if isinstance(g, (int, float))]
            return {'name': name, 'date': date, 'grades': grades} #Give back a dict which we can use to easily access the values when appending them
        # Error handling
        except ValueError as e:
            print(f"ValueError at row {row}. Please insert the correct datetime format.")
            self.__error_log.append(["ValueError", row, str(e)])
            return None
        except IndexError as e:
            print(f"IndexError at row {row}. Please insert the data at the correct slots.")
            self.__error_log.append(["IndexError", row, str(e)])
            return None
        except Exception as e:
            print(f"Unknown error at row {row}.\n Error code: {e}")
            self.__error_log.append(["UNKNOWN", row, str(e)])
            return None
        
        
    
    def create_main_report(self, title: str = "Main Report") -> Self:
        """Create the main report sheet"""
        self.__main_sheet = self.__workbook.create_sheet(title)
        return self
        
        

    def create_data_records(self) -> Self:
        """Creates the main report with all the provided data from the previous functions"""
        counter: int = 0 # Initialize a counter, this will store the amount of times the program has successfully appended a line
        if self.__data is None: #In case user hasn't called add_data() and self.__data is empty
            raise ValueError("No data found. Please call add_data() before processing.")
        
        if self.__main_sheet is None: #If user hasn't called create_main_report
            raise AttributeError("Main report not initialized. Call create_main_report() first")
        
        if self.__headers: #If the user has provided headers
            self.__main_sheet.append(self.__headers)
        else: #If the user hasn't provided headers
            if not self.__data:
                raise ValueError("Data is empty.")
            width = max(len(row) for row in self.__data)
            self.__main_sheet.append(['Names', 'Dates'] + [f'Grade {i}' for i in range(1, width - 1)])
        
        for row in self.__data:
            if not row or len(row) < 3: #If the row a) doesn't have anything or b) just has name and dates but not grades
                continue
            parsed = self._parse_row(row) #Call the helper func
            if parsed: 
                self.__main_sheet.append([parsed['name'], parsed['date'], *parsed['grades']]) #Using the dict the helper func returned
                counter += 1 
        
        print(f"Stored {counter} students in the main sheet! Go check it out!")
        return self
      

    def create_summary_sheet(self, title: str = "Summary") -> Self:
        """Initializes the summary sheet"""
        self.__summary_sheet = self.__workbook.create_sheet(title)
        return self
        
    
    def create_summary_report(self) -> Self:
        """Process loaded data into the summary sheet, calculating each student's average grade and pass/fail status."""

        if self.__data is None: #In case __data isn't initialized
            raise ValueError("No data found. Please call add_data() before processing.")
        if self.__summary_sheet is None: #In case __summary_sheet isn't initialized
            raise AttributeError("Summary sheet not initialized. Call create_summary_sheet() first")
        
        counter: int = 0 
        self.__summary_sheet.append(['Names', 'Dates', 'Average Grades', 'Status']) # Headers. For the main report we put the headers as args at add_data()
        
        for row in self.__data: 
            if not row or len(row) < 3:
                continue
            
            parsed = self._parse_row(row) #Call helper func
            if not parsed or not parsed['grades']: #It will skip this row if there are no grades
                continue

            average_grade = round(sum(parsed['grades']) / len(parsed['grades']), 1) #Finding the average grade

            status: str = "PASS" if average_grade >= self.__pass_threshold else "FAIL"

            self.__summary_sheet.append([parsed['name'], parsed['date'], average_grade, status])
            
            counter += 1
        
        print(f"Summarized {counter} students at {self.__summary_sheet.title}. Go check it out!")
        return self
    
    def make_error_report(self, error_sheet_name: str = "Error Sheet") -> Self:
        """Creates an error report. It logs every error that occurs in the program"""
        if not self.__error_log: 
            print("No errors found during processing. Skipping error sheet.")
            return self

        self.__error_sheet = self.__workbook.create_sheet(error_sheet_name)

        self.__error_sheet.append(["Error Type", "Row", "Technical Details"])

        for error_entry in self.__error_log: 
            self.__error_sheet.append(error_entry)
        
        print(f"Error report generated with {len(self.__error_log)} entries.")

        return self
       
    
    def save_file(self, title: str = "Student Report") -> Self:
        """Save the file"""
        if not title.endswith('.xlsx'):  #Checks if the provided title has a suffix or not. 
            title += '.xlsx'
        
        try:
            self.__workbook.save(title)
            print(f"Successfully saved to {title}")
        except PermissionError:
            print(f"Error: Could not save {title}. Please close the file in Excel and try again.")
        
        return self
        

if __name__ == "__main__":
    # Example, you can modify it :)
    list_of_students = [
        ['Jane', '2024-06-20', 40, 70, 100, 56],
        ['Bob', '2026-06-20', 93, 56, 74, 89],
    ]

    # Calls the class and its methods, making a nicely formatted Excel workbook with multiple sheets
    (ExcelReportBuilder(60)
        .add_data(list_of_students)
        .create_main_report()
        .create_data_records()
        .create_summary_sheet("Student Summary")
        .create_summary_report()
        .make_error_report("Error Log")
        .save_file("student_summaries.xlsx"))

"""
Calling the methods from the class should follow one of two structures (or both): 
1) add_data() -> create_main_report() -> create_data_records()
2) add_data() -> create_summary_sheet() -> create_summary_report()
Both followed by: make_error_report() (optional) -> save_file()
"""



