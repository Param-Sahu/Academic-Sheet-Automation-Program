import pandas as pd

# -----------------------------
# Step 1: Read the Excel file
# -----------------------------
# Load the Excel file without headers so that rows and columns can be accessed by their index.
df_excel = pd.read_excel("SEM-I_master sheet.xls", header=None)
df_student = pd.read_excel("student_detail.xls",header=None)

# Extract Session for particular Batch
session = df_student.iloc[2,1]

# Extract Parents Names for all Students 
fathers = df_student.iloc[4:,4].reset_index(drop = True)
mothers = df_student.iloc[4:,5].reset_index(drop = True)

# Extract roll numbers from column C (index 2) starting from row 4 (index 3)
roll_numbers = df_excel.iloc[3:, 2].reset_index(drop=True)

# Extract enrollment numbers from column D (index 3) starting from row 4 (index 3)
enrollment_numbers = df_excel.iloc[3:, 3].reset_index(drop=True)

# Extract Names of Student from column 'B' (index 1) starting from row 4 (index 3)
names = df_excel.iloc[3:, 1].reset_index(drop=True)

# -----------------------------
# Step 2: Read the CSV file
# -----------------------------
# Load UTD.csv; here it's assumed that the CSV does not have a header.
df_utd = pd.read_csv("UTD.csv", header=None)

# -----------------------------
# Step 3: Ensure UTD.csv has enough rows and columns
# -----------------------------
# Calculate required number of rows (roll/enrollment numbers start at row 3)
required_rows = max(len(roll_numbers), len(enrollment_numbers)) + 2  # Account for first 2 rows
if len(df_utd) < required_rows:
    extra_rows = pd.DataFrame([[""] * df_utd.shape[1]] * (required_rows - len(df_utd)))
    df_utd = pd.concat([df_utd, extra_rows], ignore_index=True)

# Step 4: Fill constant data for all students
# -----------------------------
# For each student (starting at row 3 i.e. index 2), fill in columns A-F with constant values.
for i in range(2, required_rows):
    df_utd.iat[i, 0] = "DEVI AHILYA VISHWAVIDYALAYA INDORE"  # University Name (Column A)
    df_utd.iat[i, 1] = "SCHOOL OF INSTRUMENTATION INDORE"      # College Name (Column B)
    df_utd.iat[i, 2] = "M.TECH (IOT)"                           # Course Name in Short (Column C)
    df_utd.iat[i, 3] = "MASTER OF TECHNOLOGY"                   # Full Course Name (Column D)
    df_utd.iat[i, 4] = "MASTER OF TECHNOLOGY"                   # Full Course Name in Detail (Column E)
    df_utd.iat[i, 5] = "INTERNET OF THINGS"     # Stream (Column F)
    df_utd.iat[i,7] = session            # Session for Batch      
# -----------------------------
# Step 5: Write the data into UTD.csv
# -----------------------------
# Write enrollment numbers into column I (index 8) starting from row 3 (index 2)
for i, enrollment in enumerate(enrollment_numbers):
    df_utd.iat[i + 2, 8] = enrollment

# Write roll numbers into column J (index 9) starting from row 3 (index 2)
for i, roll_number in enumerate(roll_numbers):
    df_utd.iat[i + 2, 9] = roll_number

# Write Names into column  (index 9) starting from row 3 (index 2)
for i, name in enumerate(names):
    df_utd.iat[i + 2, 10] = name

for i, father  in enumerate(fathers):
    df_utd.iat[i + 2, 13] = father

for i, mother in enumerate(mothers):
    df_utd.iat[i + 2, 14] = mother
# -----------------------------
# Step 5: Save the updated CSV file
# -----------------------------
df_utd.to_csv("UTD.csv", index=False, header=False)

print("Enrollment numbers and roll numbers and Name ")
print("have been successfully written to UTD.csv .")
