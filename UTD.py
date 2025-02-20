import pandas as pd

# -----------------------------
# Step 1: Read the Excel file
# -----------------------------
# Read the Excel file without inferring headers so that row indexing remains consistent.
df_excel = pd.read_excel("SEM-I_master sheet_test.xls", header=None)

# Extract the roll numbers from column C (index 2) starting from row 4 (index 3)
roll_numbers = df_excel.iloc[3:, 2].reset_index(drop=True)

# -----------------------------
# Step 2: Read the CSV file
# -----------------------------
# Read UTD.csv; if it already has headers and data, adjust accordingly.
# Here we assume the CSV file does not include a header row.
df_utd = pd.read_csv("UTD.csv", header=None)

# -----------------------------
# Step 3: Ensure UTD.csv has enough rows and columns
# -----------------------------
# We need at least as many rows as (2 + number of roll numbers) because we start writing from row 3.
required_rows = len(roll_numbers) + 2  # rows 1 and 2 are assumed to be pre-existing (headers or otherwise)
if len(df_utd) < required_rows:
    # Create a DataFrame with the extra rows filled with empty strings
    extra_rows = pd.DataFrame([[""] * df_utd.shape[1]] * (required_rows - len(df_utd)))
    df_utd = pd.concat([df_utd, extra_rows], ignore_index=True)

# We also need at least 10 columns (A to J, with column J at index 9)
if df_utd.shape[1] < 10:
    # Add extra columns as needed (fill with empty strings)
    for i in range(df_utd.shape[1], 10):
        df_utd[i] = ""

# -----------------------------
# Step 4: Write roll numbers into column J of UTD.csv
# -----------------------------
# Write the roll numbers starting at row 3 (i.e. index 2) in column J (i.e. index 9)
for i, roll in enumerate(roll_numbers):
    df_utd.iat[i + 2, 9] = roll

# -----------------------------
# Step 5: Save the updated CSV file
# -----------------------------
df_utd.to_csv("UTD.csv", index=False, header=False)

print("Roll numbers have been successfully written to UTD.csv in column J starting from row 3.")
