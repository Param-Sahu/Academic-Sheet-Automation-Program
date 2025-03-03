# UTD.py Documentation

## Overview
The `UTD.py` script is designed to automate the process of extracting student and semester data from Excel files and populating a CSV file (`UTD.csv`) with this information. The script reads data from two Excel files (`student_detail.xls` and `SEM-I_master sheet.xls`), processes the data, and writes it into a new CSV file named based on the session and current semester.

## Dependencies
- pandas

## Steps Performed by the Script

### Step 1: Read the Excel Files
The script reads two Excel files:
- `student_detail.xls`: Contains student details.
- `SEM-I_master sheet.xls`: Contains semester details.

The data is loaded into pandas DataFrames without headers to allow access by row and column indices.

### Step 2: Extract Data from Excel Files
- **Session**: Extracted from the `student_detail.xls` file.
- **Parents' Names**: Extracted from the `student_detail.xls` file.
- **Roll Numbers**: Extracted from the `SEM-I_master sheet.xls` file.
- **Enrollment Numbers**: Extracted from the `SEM-I_master sheet.xls` file.
- **Student Names**: Extracted from the `SEM-I_master sheet.xls` file.

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

### Step 6: Write the Data into UTD.csv
The script writes the extracted data into specific columns of the `UTD.csv` file:
- **Column I**: Enrollment Numbers
- **Column J**: Roll Numbers
- **Column K**: Student Names
- **Column N**: Fathers' Names
- **Column O**: Mothers' Names

### Step 7: Save the Updated CSV File
The script saves the updated CSV file with a new name based on the session and current semester.

## Output
The script generates a new CSV file named `UTD_<session>_<current_semester>.csv` containing the processed student and semester data.

## Usage
To run the script, ensure you have the required Excel files (`student_detail.xls` and `SEM-I_master sheet.xls`) and the `UTD.csv` file in the same directory as the script. Execute the script using Python:

```bash
python UTD.py
```