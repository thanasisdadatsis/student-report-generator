from openpyxl import Workbook
from datetime import datetime


class ExcelReportBuilder: 
    def __init__(self, worksheet: str, second_name: str, error_worksheet: str, data: list, title:str = 'Summaries') -> None:
        """Initialize the workbook, its sheets, the data and the title"""
        self._workbook = Workbook()
        self.sheet = self._workbook.active
        self.sheet.title = worksheet
        self.second_sheet = self._workbook.create_sheet(second_name)
        self.error_sheet = self._workbook.create_sheet(error_worksheet)
        self.data = data
        self.title = title

        #first row of the error sheet 
        self.error_sheet.append(["Failed Row", "Error call"])
    
    
    def use_data(self) -> None:
        """This function uses the data provided from __init__ and forms a readable excel sheet """
        counter: int = 0
        self.sheet.append(["Name", "Dates", "Term 1", "Term 2", "Term 3", "Exams", "Grade-if-retaken"])
        #Loop over each row and append the information in a proper way
        for row in self.data:
            if not row or len(row) < 3:
                continue
            try:
                self.sheet.append([row[0], datetime.strptime(row[1], '%Y-%m-%d'), *row[2:]])
                #keep a counter for successful appends
                counter += 1 
            except Exception as e:
                print(f"Found exception at row {row}: \n{e}")
                self.error_sheet.append([str(row), str(e)])
        print(f"Stored {counter} number of students in 1st sheet!")
        
    
    def second_worksheet(self) -> None: 
        """Summarizes the student data from self.data into the second sheet with average grades and pass/fail status """ 
        counter: int = 0
        #iterating over the data 
        self.second_sheet.append(["Name", "Date", "Average Grades", "Status"])
        for row in self.data: 
            try:
                name: str = row[0]
                date = datetime.strptime(row[1], '%Y-%m-%d')
                grades: list = [g for g in row[2:] if isinstance(g, (int, float))] #In order to take only ints or floats, not bools or strs

                if not grades: #in case there aren't any grades at all 
                    continue
                
                average_grade = round(sum(grades) / len(grades), 1)
                status = "PASS" if average_grade >= 70 else "FAIL"

                self.second_sheet.append([name, date, average_grade, status])

                counter += 1 
            except Exception as e: 
                print(f"Found exception at row {row}: \n{e}")
                self.error_sheet.append([str(row), str(e)])
                
        print(f"Summarized {counter} students at sheet 2. Go check it out!")
    
    def save_file(self) -> None:
        #What to do if there are no errors recorded
        if self.error_sheet.max_row == 1:
                self.error_sheet.append(["None", "None"])
        self._workbook.save(self.title + '.xlsx') #save the file
        print("File saved! Program over")
    
    #A function to run all methods
    def main(self):
        self.use_data()
        self.second_worksheet()
        self.save_file()


#test - you can alter it if you like :)
list_of_students = [
    ['Jane', '2024-06-20', 40, 70, 100, 45, 80],
    ['Bob', '2026-06-20', 93, 56, 74, 80, 56],
    ['Jane', '2026-06-20', 37, 68, 9, 40, 89],
    ['Hillary', '2026-09-11', 12, 34, 23, 70, 99],
    ['Lily', '2026-06-20', 67, 70, 59, 80, 91]
]

#run everything
if __name__ == "__main__":
    report = ExcelReportBuilder("Sheet 1", "Summary", "Error Sheet", list_of_students, 'Student_Summary')
    #Apply the main() function cleanly
    report.main()

  
            

