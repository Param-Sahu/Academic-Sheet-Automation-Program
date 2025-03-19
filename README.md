# UTD.py Documentation

## Overview
The `UTD.py` script automates the process of extracting student and semester data from Excel files and populating a CSV file (`UTD.csv`) with this information. The script reads data from two Excel files (`SEM-II_master sheet.xls` and `student_detail_new.xls`), processes the data, and writes it into a new CSV file named based on the session and current semester.

## Dependencies
- pandas
- roman
- warnings

## Steps Performed by the Script

### Step 1: Read the Excel Files
The script reads two Excel files:
- `SEM-II_master sheet.xls`: Contains semester details.
- `student_detail_new.xls`: Contains student details.

The data is loaded into pandas DataFrames without headers to allow access by row and column indices.

### Step 2: Extract Data from Excel Files
- **Batch**: Extracted from the `student_detail_new.xls` file.
- **Parents' Names**: Extracted from the `student_detail_new.xls` file.
- **Roll Numbers**: Extracted from the `SEM-II_master sheet.xls` file.
- **Enrollment Numbers**: Extracted from the `SEM-II_master sheet.xls` file.
- **Student Names**: Extracted from the `SEM-II_master sheet.xls` file.
- **Semester Number**: Extracted from the roll numbers.
- **Attempts**: Extracted from the `student_detail_new.xls` file.
- **Credits, SGPA, and Results**: Extracted from the `SEM-II_master sheet.xls` file.
- **Subjects**: Extracted from the `SEM-II_master sheet.xls` file.
- **Grades and Grade Points**: Calculated from the subjects and credits.

### Step 3: Read the CSV File
The script reads the `UTD.csv` file into a pandas DataFrame without headers.

### Step 4: Ensure UTD.csv Has Enough Rows and Columns
The script calculates the required number of rows based on the length of the roll and enrollment numbers. If the CSV file does not have enough rows, additional rows are added.

### Step 5: Fill Constant Data for All Students
The script fills columns A-F with constant values for all students:
- **Column A**: University Name
- **Column B**: College Name
- **Column C**: Course Name in Short
- **Column D**: Full Course Name
- **Column E**: Full Course Name in Detail
- **Column F**: Stream
- **Column H**: Session for Batch
- **Column S**: Year for Batch
- **Column T**: Month for Batch
- **Column X**: Semester Number
- **Column Y**: Semester Type
- **Column AB**: Total Credits

### Step 6: Write the Data into UTD.csv
The script writes the extracted data into specific columns of the `UTD.csv` file:
- **Column I**: Enrollment Numbers
- **Column J**: Roll Numbers
- **Column K**: Student Names
- **Column N**: Fathers' Names
- **Column O**: Mothers' Names
- **Column Q**: Marksheet Status
- **Column R**: Results
- **Column U**: Division
- **Column AC**: Credits Earned
- **Column AF**: Credits Earned
- **Column AG**: CGPA
- **Column AI**: SGPA
- **Column AM**: Student Names

The script also fills the constant data for subjects, credits, and grades for each student.

### Step 7: Save the Updated CSV File
The script saves the updated CSV file with a new name based on the session and current semester.

## Output
The script generates a new CSV file named `UTD_<batch>_<current_semester>.csv` containing the processed student and semester data.

## Usage
To run the script, ensure you have the required Excel files (`SEM-II_master sheet.xls` and `student_detail_new.xls`) and the `UTD.csv` file in the same directory as the script. Execute the script using Python:

```bash
python [UTD.py](http://_vscodecontentref_/1)