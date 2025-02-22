import pandas as pd

# -----------------------------
# Step 1: Read the Excel file
# -----------------------------
# Load the Excel file without headers so that rows and columns can be accessed by their index.
df_excel = pd.read_excel("SEM-I_master sheet_test.xls", header=None)

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
df_utd = pd.read_csv("UTD_test.csv", header=None)

# -----------------------------
# Step 3: Ensure UTD.csv has enough rows and columns
# -----------------------------
# Calculate required number of rows (roll/enrollment numbers start at row 3)
required_rows = max(len(roll_numbers), len(enrollment_numbers)) + 2  # Account for first 2 rows
if len(df_utd) < required_rows:
    extra_rows = pd.DataFrame([[""] * df_utd.shape[1]] * (required_rows - len(df_utd)))
    df_utd = pd.concat([df_utd, extra_rows], ignore_index=True)

# -----------------------------
# Step 4: Write the data into UTD.csv
# -----------------------------
# Write enrollment numbers into column I (index 8) starting from row 3 (index 2)
for i, enrollment in enumerate(enrollment_numbers):
    df_utd.iat[i + 2, 8] = enrollment

# Write roll numbers into column J (index 9) starting from row 3 (index 2)
for i, roll in enumerate(roll_numbers):
    df_utd.iat[i + 2, 9] = roll

# Write Names into column  (index 9) starting from row 3 (index 2)
for i, name in enumerate(names):
    df_utd.iat[i + 2, 10] = name

# -----------------------------
# Step 5: Save the updated CSV file
# -----------------------------
df_utd.to_csv("UTD_test.csv", index=False, header=False)

print("Enrollment numbers and roll numbers and Name ")
print("have been successfully written to UTD.csv .")
