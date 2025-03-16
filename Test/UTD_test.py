import pandas as pd
from roman import fromRoman
student_sheet = "student_detail_test.xls"
sem_sheet = "SEM-I_master sheet_test.xls"
all_sem_file = "student_detail_new.xls"
UTD_file = "UTD_test.csv"
current_semester = sem_sheet.split('_')[0] # Extracting Current Semester from file name.
sem_number = fromRoman(current_semester.split('-')[1]) # Extracting Semester Number from file name.
semester_attempt = "Attempt_"+str(sem_number)
# -----------------------------
# Step 1: Read the Excel file
# -----------------------------
# Load the Excel file without headers so that rows and columns can be accessed by their index.
df_excel = pd.read_excel(sem_sheet, header=None)
df_student = pd.read_excel(student_sheet,header=None)
df_all_sem = pd.read_excel(all_sem_file, header=3)

# Extract Attempts for all Students
attempts = df_all_sem.loc[:,semester_attempt]
attempts = attempts.dropna()
attempts = attempts.astype(int)
marksheet_status = attempts.apply(lambda x: "O" if x ==1 else "M") # O for Original and M for Modified Marksheet (represents ATKT)

# Extract Session for particular Batch
session = df_student.iloc[2,1]

# Extract Parents Names for all Students 
fathers_name = df_student.iloc[4:,4].reset_index(drop = True)
mothers_name = df_student.iloc[4:,5].reset_index(drop = True)

# Extract roll numbers from column C (index 2) starting from row 4 (index 3)
roll_numbers = df_excel.iloc[3:, 2].reset_index(drop=True)

# Extract enrollment numbers from column D (index 3) starting from row 4 (index 3)
enrollment_numbers = df_excel.iloc[3:, 3].reset_index(drop=True)

# Extract Names of Student from column 'B' (index 1) starting from row 4 (index 3)
student_names = df_excel.iloc[3:, 1].reset_index(drop=True)
# Extract Credits, SGPA and Results of all Students
credits = df_excel.iloc[2, 4:].reset_index(drop=True)
credits = credits.dropna()
total_subjects, total_credits = len(credits), sum(credits)
sgpa = df_excel.iloc[3:, 4 + total_subjects ].reset_index(drop=True)
results = df_excel.iloc[3:, 5 + total_subjects].reset_index(drop=True)

# Extract Subjects for all students
subjects = df_excel.iloc[:,4:4 + total_subjects].reset_index(drop=True)
subjects.columns = range(total_subjects)
subjects_id = subjects.iloc[0,:] # Extracting Subject IDs from first row of subjects
subjects_names = subjects.iloc[1,:]
grades = subjects.iloc[3:, :].reset_index(drop=True)
grade_points = {'O':10, 'A+':9, 'A':8, 'B+':7, 'B':6, 'C':5, 'P':4, 'F':0, 'Ab':0} # Dictionary for Grade points with respective grade.


division = sgpa.apply(lambda x: 'FIRST WITH DISTINCTION' if x >= 8.0 
                      else ('FIRST' if x >= 6.5
                      else ('SECOND' if x >= 5.0 
                            else ('PASS' if x >= 4.0 else 'FAIL'))))

if current_semester == "SEM-I":
    cgpa = sgpa.copy()
    YEAR = int(session.split('-')[0])+1
    if YEAR == 2023:
        MONTH = "MARCH"
    else:
        MONTH = "JANUARY"

# -----------------------------
# Step 2: Read the CSV file
# -----------------------------
# Load UTD.csv; here it's assumed that the CSV does not have a header.
df_utd = pd.read_csv(UTD_file, header=None)

# -----------------------------
# Step 3: Ensure UTD.csv has enough rows and columns
# -----------------------------
# Calculate required number of rows (roll/enrollment numbers etc. start at row 3)
required_rows = max(len(roll_numbers), len(enrollment_numbers)) + 2  # Account for first 2 rows
if len(df_utd) < required_rows:
    extra_rows = pd.DataFrame([[""] * df_utd.shape[1]] * (required_rows - len(df_utd)))
    df_utd = pd.concat([df_utd, extra_rows], ignore_index=True)

# Step 4: Fill constant data for all students
# -----------------------------
# For each student (starting at row 3 i.e. index 2), fill in columns A-F with constant values.

df_utd.iloc[2:required_rows, 0] = "DEVI AHILYA VISHWAVIDYALAYA INDORE"   # University Name (Column A)
df_utd.iloc[2:required_rows, 1] = "SCHOOL OF INSTRUMENTATION INDORE"     # College Name (Column B)
df_utd.iloc[2:required_rows, 2] = "INTEGRATED M.TECH (IOT)"              # Course Name in Short (Column C)
df_utd.iloc[2:required_rows, 3] = "INTEGRATED MASTER OF TECHNOLOGY"      # Full Course Name (Column D)
df_utd.iloc[2:required_rows, 4] = "INTEGRATED MASTER OF TECHNOLOGY"      # Full Course Name in Detail (Column E)
df_utd.iloc[2:required_rows, 5] = "INTERNET OF THINGS"                   # Stream (Column F)
df_utd.iloc[2:required_rows, 7] = session                                # Session for Batch  
df_utd.iloc[2:required_rows, 18] = YEAR                                  # Year for Batch  
df_utd.iloc[2:required_rows, 19] = MONTH                                 # MONTH for Batch  
df_utd.iloc[2:required_rows, 23] = current_semester.split('-')[1]        # Semester for Batch , Extracting Semester Number from file name. By SEM-I , spliting it by '-' and taking 2nd part of it.
df_utd.iloc[2:required_rows, 24] = current_semester.split('-')[0]        # Semester for Batch , Extracting Sem or Year from file name. By SEM-I , spliting it by '-' and taking 1st part of it.
df_utd.iloc[2:required_rows, 27] = total_credits                           # Total Credits for Batch
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
                32: cgpa,
                34:sgpa,
                38:student_names
                 }  # Dictionary of data to be written into UTD.csv

# Write data into UTD.csv
for column, values in variable_data.items():
    for i, value in enumerate(values):
        df_utd.iat[i + 2, column] = value


j=0
for i in range(total_subjects):
    df_utd.iloc[2:required_rows, 39 + j] = subjects_names[i]
    df_utd.iloc[2:required_rows, 40 + j] = subjects_id[i]
    df_utd.iloc[2:required_rows, 51 + j] = credits[i]
    j+=15

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
# -----------------------------
# Step 5: Save the updated CSV file
# -----------------------------
UTD_output_file = "UTD_" + session + '_' + current_semester + '.csv' # Creating a seperate CSV file for different semesters data.

try:
    df_utd.to_csv(UTD_output_file, index=False, header=False)
    print("Student Data and Semester Data has been Successfully added to UTD Output file. ")
except Exception as e:
    print("Error occurred while writing data to UTD file: ", e)
    print("Please check if the file is open in another program and close it before running the script again.")