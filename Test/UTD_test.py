import pandas as pd
from roman import toRoman
import warnings
warnings.filterwarnings("ignore", category=FutureWarning)

sem_sheet = "SEM-I_master sheet_test.xls"
all_sem_file = "student_detail_new.xls"
UTD_file = "UTD_test.csv"
abc_file = "ABC ID's_2022_test.xlsx"

try:
    # -----------------------------
    # Step 1: Read the Excel file
    # -----------------------------
    # Load the Excel file without headers so that rows and columns can be accessed by their index.
    df_excel = pd.read_excel(sem_sheet, header=None)

    df_all_sem = pd.read_excel(all_sem_file, header=3)

    # Extract Batch Information for all students (Row index 2, Column index 1 assuming zero-based index)
    batch =  (pd.read_excel(all_sem_file, header=None)).iloc[2, 1]  # Row 2, Column B in Excel (0-based index)
except Exception as e:
    print("Error occurred while reading the Excel files: ", e)
    print("Please check if the file exists and the path is correct.")

try:
    # Extract Parents Names for all Students 
    fathers_name = df_all_sem.loc[:,"Father’s Name"].reset_index(drop = True)
    fathers_name = fathers_name.dropna()

    mothers_name = df_all_sem.loc[:,"Mother’s Name"].reset_index(drop = True)
    mothers_name = mothers_name.dropna()

    # Extract roll numbers from column C (index 2) starting from row 4 (index 3)
    roll_numbers = df_excel.iloc[3:, 2].reset_index(drop=True)

    # Extracting Semester for Batch
    sem_number = int(roll_numbers[0][5:-3]) # Extracting Semester Number from roll number.
    current_semester = "SEM-" + str(toRoman(sem_number))
    semester_attempt = "Attempt_"+str(sem_number)

    # Extract Attempts for all Students
    attempts = df_all_sem.loc[:,semester_attempt]
    attempts = attempts.dropna()
    attempts = attempts.astype(int)
    marksheet_status = attempts.apply(lambda x: "O" if x ==1 else "M") # O for Original and M for Modified Marksheet (represents ATKT)

    # Extract enrollment numbers from column D (index 3) starting from row 4 (index 3)
    enrollment_numbers = df_excel.iloc[3:, 3].reset_index(drop=True)

    # Extract Names of Student from column 'B' (index 1) starting from row 4 (index 3)
    student_names = df_excel.iloc[3:, 1].reset_index(drop=True)
    # Extract Credits, SGPA and Results of all Students
    subject_codes = df_excel.iloc[0, 4:].reset_index(drop=True)
    subject_codes = subject_codes.dropna().reset_index(drop=True).apply(lambda x:x.strip()) # Drop columns where all values are NaN
    total_subjects = len(subject_codes)
    credits = df_excel.iloc[2, 4:4+total_subjects].reset_index(drop=True)
    total_credits = sum(credits)
    sgpa = df_excel.iloc[3:, 4 + total_subjects ].reset_index(drop=True)
    results = df_excel.iloc[3:, 5 + total_subjects].reset_index(drop=True)

    # Extract Subjects for all students
    subjects = df_excel.iloc[:,4:4 + total_subjects].reset_index(drop=True)
    subjects.columns = range(total_subjects)
    subjects_names = subjects.iloc[1,:]

    # Extracting Grades and calculating grade points for all students
    grades = subjects.iloc[3:, :].reset_index(drop=True)
    grade_points = {'O':10, 'A+':9, 'A':8, 'B+':7, 'B':6, 'C':5, 'P':4, 'F':0, 'Ab':0} # Dictionary for Grade points with respective grade.

    # Replace grades with corresponding grade points
    grade_points_df = grades.replace(grade_points)

    # Multiply each column by its corresponding credit value and sum row-wise
    credits_earn = grade_points_df.mul(credits, axis=1).sum(axis=1)


    division = sgpa.apply(lambda x: 'FIRST WITH DISTINCTION' if x >= 8.0 
                        else ('FIRST' if x >= 6.5
                        else ('SECOND' if x >= 5.0 
                                else ('PASS' if x >= 4.0 else 'FAIL'))))

    # Calculating CGPA according to Semesters (Total SGPA till current Semester)/semester_number
    total_sgpa = df_all_sem.loc[:len(roll_numbers)-1 ,'Sem1':'Sem'+str(sem_number)].sum(axis=1)
    cgpa = round(total_sgpa/sem_number,3)

except Exception as e:
    print("Error occurred while reading Student Details: ", e)
    print("Please check File contains correct Student Details at correct Index.")

try:
    # Deciding Month and Year of Semester Exam according to semesters.
    if sem_number ==1 :
        YEAR = int(batch.split('-')[0])+1
        if YEAR == 2023:
            MONTH = "MARCH"
        else:
            MONTH = "JANUARY"

    elif sem_number ==2 :
        YEAR = int(batch.split('-')[0])+1
        if YEAR == 2023:
            MONTH = "JULY"
        else:
            MONTH = "JUNE"

    elif sem_number == 3:
        YEAR = int(batch.split('-')[0])+1
        MONTH = "DECEMBER"
    elif sem_number == 4:
        YEAR = int(batch.split('-')[0]) + 2
        MONTH = "MAY"
    elif sem_number == 5:
        YEAR = int(batch.split('-')[0]) + 2
        MONTH = "DECEMBER"
    elif sem_number == 6:
        YEAR = int(batch.split('-')[0]) + 3
        MONTH = "MAY"
    elif sem_number == 7:
        YEAR = int(batch.split('-')[0]) + 3
        MONTH = "DECEMBER"
    elif sem_number == 8:
        YEAR = int(batch.split('-')[0]) + 4
        MONTH = "MAY"
    elif sem_number == 9:
        YEAR = int(batch.split('-')[0]) + 4
        MONTH = "DECEMBER"
    elif sem_number == 10:
        YEAR = int(batch.split('-')[1])
        MONTH = "MAY"

    provided_date = str(df_excel.iloc[0, 1]).strip()
    if provided_date != "" and provided_date.lower() != 'nan':
        _ , exam_month, exam_year = provided_date.split()
        MONTH = exam_month.upper()
        YEAR = int(exam_year)

except Exception as e:
    print("Error occurred while Configuring Month and Year: ", e)
    print("Please check that Batch must be in correct format, Ex:'2000-2005'")

# -----------------------------
# Step 2: Read the CSV file
# -----------------------------
# Load UTD.csv; here it's assumed that the CSV does not have a header.
try:
    df_utd = pd.read_csv(UTD_file, header=None)
except Exception as e:
    print("Error occured while reading output 'CSV' file.")
    print("Please check if the file exists and in CSV format and the path is correct.")

try:
    # -----------------------------
    # Step 3: Ensure UTD.csv has enough rows and columns
    # -----------------------------
    # Calculate required number of rows (roll/enrollment numbers etc. start at row 3)
    required_rows = max(len(roll_numbers), len(enrollment_numbers)) + 2  # Account for first 2 rows
    if len(df_utd) < required_rows:
        extra_rows = pd.DataFrame([[""] * df_utd.shape[1]] * (required_rows - len(df_utd)))
        df_utd = pd.concat([df_utd, extra_rows], ignore_index=True)
except Exception as e:
    print("Error Occured while adding the rows in UTD file.")
    
try:
    # Extract Course Information for all students (Row index 0, Column index 1 assuming zero-based index)
    course_name =  (pd.read_excel(all_sem_file, header=None)).iloc[0, 1]  # Row 2, Column B in Excel (0-based index)
    # Extract stream Information for all students (Row index 1, Column index 1 assuming zero-based index)
    stream =  (pd.read_excel(all_sem_file, header=None)).iloc[1, 1]  # Row 1, Column 1 in Excel (0-based index)
    # Step 4: Fill constant data for all students
    # -----------------------------
    # For each student (starting at row 3 i.e. index 2), fill in columns A-F with constant values.
    if "five" in course_name.lower() or "dual" in course_name.lower():
        course_full_name = "B.TECH.+M.TECH. (INTERNET OF THINGS) DUAL DEGREE 5 YRS."
    else:
        course_full_name = f"{course_name} {stream}"

    df_utd.iloc[2:required_rows, 0] = "DEVI AHILYA VISHWAVIDYALAYA INDORE"   # University Name (Column A)
    df_utd.iloc[2:required_rows, 1] = "SCHOOL OF INSTRUMENTATION"     # College Name (Column B)
    df_utd.iloc[2:required_rows, 2] = course_name              # Course Name in Short (Column C)
    df_utd.iloc[2:required_rows, 3] = course_full_name      # Full Course Name (Column D)
    df_utd.iloc[2:required_rows, 4] = course_full_name         # Full Course Name in Detail (Column E)
    df_utd.iloc[2:required_rows, 5] = stream                # Stream (Column F)
    df_utd.iloc[2:required_rows, 7] = batch                                  # Session for Batch  
    df_utd.iloc[2:required_rows, 18] = YEAR                                  # Year for Batch  
    df_utd.iloc[2:required_rows, 19] = MONTH                                 # MONTH for Batch  
    df_utd.iloc[2:required_rows, 23] = current_semester.split('-')[1]        # Semester for Batch , Extracting Semester Number. By SEM-I , spliting it by '-' and taking 2nd part of it.
    df_utd.iloc[2:required_rows, 24] = current_semester.split('-')[0]        # Semester for Batch , Extracting Sem or Year Exam Type. By SEM-I , spliting it by '-' and taking 1st part of it.
    df_utd.iloc[2:required_rows, 27] = total_credits                         # Total Credits for Batch
    df_utd.iloc[2:required_rows, 339] = batch.split('-')[0]                  # Admission Year 

except Exception as e:
    print("Error occurred while filling constant data for all students: ", e)

try:
    # Extracting ABC_ID from file but only till required rows.(As per number of students)
    abc_id = pd.read_excel(abc_file,header=0).loc[:required_rows-3,"ABC ID"].astype(str) # required rows-3 indicated included rows 
except Exception as e:
    print("Error occurred while reading ABC ID from Excel: ", e)

try:
    # -----------------------------
    # Step 5: Write the data into UTD.csv
    # -----------------------------
    # Dictionary of data to be written into UTD.csv in the format {column: data}
    variable_data = {8: enrollment_numbers,  # Column : Data format
                    9: roll_numbers,
                    10:student_names,
                    13:fathers_name,
                    14:mothers_name,
                    16:marksheet_status,
                    17:results,
                    20: division,
                    28: credits_earn,
                    31: credits_earn,
                    32: cgpa,
                    34:sgpa,
                    35:abc_id,
                    38:student_names
                    }  # Dictionary of data to be written into UTD.csv

    # Write data into UTD.csv
    for column, values in variable_data.items():
        for i, value in enumerate(values):
            df_utd.iat[i + 2, column] = value

except Exception as e:
    print("Error occurred while writing variable data for all students: ", e)

try:
    # Filling the constant data Subject Name, Credit and Code for every student.
    j=0
    for i in range(total_subjects):
        df_utd.iloc[2:required_rows, 39 + j] = subjects_names[i]
        df_utd.iloc[2:required_rows, 40 + j] = subject_codes[i]
        df_utd.iloc[2:required_rows, 51 + j] = credits[i]
        j+=15

    # Filling the Grades row-wise for each student in the interval of 15 columns for each subject.
    k=0
    for i in range(total_subjects):
        for j,grade in enumerate(grades.iloc[:,i]):
            
            if grade=='F' or grade=='Ab':
                df_utd.iat[2+j, 48 + k] = 'FAIL'
            else:
                df_utd.iat[2+j, 48 + k] = 'PASS'

            df_utd.iat[2+j, 49 + k] = grade
            df_utd.iat[2+j, 50 + k] = grade_points[grade]
            df_utd.iat[2+j, 52 + k] = grade_points[grade]*credits[i]
        k+=15

except Exception as e:
    print("Error occurred while filling Grades and calculating Total Grade Point and Credits: ", e)
    print("Please check the integrity of the Grades data.")

# -----------------------------
# Step 5: Save the updated CSV file
# -----------------------------
UTD_output_file = "UTD_" + batch + '_' + current_semester + '.csv' # Creating a seperate CSV file for different semesters data.

try:
    df_utd.to_csv(UTD_output_file, index=False, header=False)
    print("Student Data and Semester Data has been Successfully added to UTD Output file. ")
except Exception as e:
    print("Error occurred while writing data to UTD file: ", e)
    print("Please check if the file is open in another program and close it before running the script again.")