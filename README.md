# ExcelParser

Schedule Generator
This C# program generates a schedule for faculties based on input data from Excel files. It uses the Newtonsoft.Json library for JSON serialization and OfficeOpenXml for working with Excel files. The generated schedule is saved in a JSON file.

Newtonsoft.Json library is used for JSON serialization. Make sure to install it before running the program.
Usage
Input Data: Provide the input data in Excel format. The program reads two Excel files: fen.xlsx and fi.xlsx.

Excel Format:

Each row in the Excel file represents a schedule entry.
The columns contain the following information:
Day of the week
Time slot
Subject name (may include specialization information in parentheses)
Group name
Classroom
Weeks when the class is scheduled
Generating Schedule:

The program processes the input data from Excel files to generate a schedule for two faculties: "Факультет економічних наук" and "Факультет інформатики".
Specializations and subjects are extracted from the input data.
The generated schedule includes details such as day of the week, time slot, subject, group, classroom, and weeks.
Output: The generated schedule is saved in a JSON file named schedule.json.

How to Run
Clone the repository or download the source code and navigate to the project directory.
Open the project in Visual Studio or any C# compatible IDE.
Make sure to install the Newtonsoft.Json library if you haven't already.
Compile and run the Program.cs file.
Generated Schedule Structure
The generated schedule follows the structure below:

{
  "Faculties": {
    "Назва факультета": {
      "Specializations": {
        "Спеціальність": {
          "Subjects": {
            "Назва дисципліни": {
              "Groups": [
                {
                  "Name": "номер групи",
                  "Time": "час пари",
                  "Weeks": "номери тижнів, коли йде цей предмет",
                  "Classroom": "аудиторія",
                  "DayOfWeek": "день тижня"
                }
              ]
            },
          }
        }
      }
    }
  }
}
