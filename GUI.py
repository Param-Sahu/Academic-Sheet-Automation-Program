import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from roman import toRoman
import warnings
warnings.filterwarnings("ignore", category=FutureWarning)

# Function to validate file extensions
def validate_files(sem_sheet, all_sem_file, abc_file, utd_file):
    if not sem_sheet.endswith('.xls') and not sem_sheet.endswith('.xlsx'):
        messagebox.showerror("Invalid File", "Please select a valid Excel file for 'sem_sheet'.")
        return False
    if not all_sem_file.endswith('.xls') and not all_sem_file.endswith('.xlsx'):
        messagebox.showerror("Invalid File", "Please select a valid Excel file for 'all_sem_file'.")
        return False
    if not abc_file.endswith('.xls') and not abc_file.endswith('.xlsx'):
        messagebox.showerror("Invalid File", "Please select a valid Excel file for 'abc_file'.")
        return False
    if not utd_file.endswith('.csv'):
        messagebox.showerror("Invalid File", "Please select a valid CSV file for 'UTD_file'.")
        return False
    return True

# Function to run the main program
def run_program():
    if not sem_sheet_path.get() or not all_sem_file_path.get() or not abc_file_path.get() or not utd_file_path.get():
        messagebox.showerror("Missing Files", "Please select all required files.")
        return

    sem_sheet = sem_sheet_path.get()
    all_sem_file = all_sem_file_path.get()
    abc_file = abc_file_path.get()
    utd_file = utd_file_path.get()

    if not validate_files(sem_sheet, all_sem_file, abc_file, utd_file):
        return

    try:
        # Main program logic (existing code)
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
            messagebox.showerror("Error occurred while reading the Excel files: ", e)
            messagebox.showerror("Please check if the file exists and the path is correct.")

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
            total_sem_sgpa = df_all_sem.loc[:len(roll_numbers)-1 ,'Sem1':'Sem'+str(sem_number)]
            total_sem_credits = df_all_sem.loc[:0,'Credit_1':'Credit_'+str(sem_number)]
            cgpa = total_sem_sgpa.mul(total_sem_credits.iloc[0].values, axis=1).sum(axis=1) / total_sem_credits.iloc[0].sum()
            cgpa = cgpa.round(2)

        except Exception as e:
            messagebox.showerror("Error occurred while reading Student Details: ", e)
            messagebox.showerror("Please check File contains correct Student Details at correct Index.")

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
            messagebox.showerror("Error occurred while Configuring Month and Year: ", e)
            messagebox.showerror("Please check that Batch must be in correct format, Ex:'2000-2005'")

        # -----------------------------
        # Step 2: Read the CSV file
        # -----------------------------
        # Load UTD.csv; here it's assumed that the CSV does not have a header.
        try:
            df_utd = pd.read_csv(utd_file, header=None)
        except Exception as e:
            messagebox.showerror("Error occured while reading output 'CSV' file.")
            messagebox.showerror("Please check if the file exists and in CSV format and the path is correct.")

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
            messagebox.showerror("Error Occured while adding the rows in UTD file.")
            
        try:
            # Extract Course Information for all students (Row index 0, Column index 1 assuming zero-based index)
            course_name =  (pd.read_excel(all_sem_file, header=None)).iloc[0, 1]  # Row 2, Column B in Excel (0-based index)
            # Extract stream Information for all students (Row index 1, Column index 1 assuming zero-based index)
            stream =  (pd.read_excel(all_sem_file, header=None)).iloc[1, 1]  # Row 1, Column 1 in Excel (0-based index)
            # Step 4: Fill constant data for all students
            # -----------------------------
            # For each student (starting at row 3 i.e. index 2), fill in columns A-F with constant values.
            if "five" in course_name.lower() or "dual" in course_name.lower():
                course_full_name = "B.TECH.+M.TECH. DUAL DEGREE 5 YRS."
            else:
                course_full_name = f"{course_name} {stream}"
            df_utd.iloc[2:required_rows, 0] = "SCHOOL OF INSTRUMENTATION"   # University Name, currently College Name (Column A)
            df_utd.iloc[2:required_rows, 1] = "SCHOOL OF INSTRUMENTATION"     # College Name (Column B)
            df_utd.iloc[2:required_rows, 2] = course_name              # Course Name in Short (Column C)
            df_utd.iloc[2:required_rows, 3] = course_name      # Full Course Name (Column D)
            df_utd.iloc[2:required_rows, 4] = course_name         # Full Course Name in Detail (Column E)
            df_utd.iloc[2:required_rows, 5] = stream                                   # Stream (Column F)
            df_utd.iloc[2:required_rows, 7] = batch                                  # Session for Batch  
            df_utd.iloc[2:required_rows, 18] = YEAR                                  # Year for Batch  
            df_utd.iloc[2:required_rows, 19] = MONTH                                 # MONTH for Batch  
            df_utd.iloc[2:required_rows, 23] = current_semester.split('-')[1]        # Semester for Batch , Extracting Semester Number. By SEM-I , spliting it by '-' and taking 2nd part of it.
            df_utd.iloc[2:required_rows, 24] = current_semester.split('-')[0]        # Semester for Batch , Extracting Sem or Year Exam Type. By SEM-I , spliting it by '-' and taking 1st part of it.
            df_utd.iloc[2:required_rows, 27] = total_credits                         # Total Credits for Batch
            df_utd.iloc[2:required_rows, 339] = batch.split('-')[0]                  # Admission Year 

        except Exception as e:
            messagebox.showerror("Error occurred while filling constant data for all students: ", e)

        try:
            # Extracting ABC_ID from file but only till required rows.(As per number of students)
            abc_id = pd.read_excel(abc_file,header=0).loc[:required_rows-3,"ABC ID"].astype(str) # required rows-3 indicated included rows 
        except Exception as e:
            messagebox.showerror("Error occurred while reading ABC ID from Excel: ", e)

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
            messagebox.showerror("Error occurred while writing variable data for all students: ", e)

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
            messagebox.showerror("Error occurred while filling Grades and calculating Total Grade Point and Credits: ", e)
            messagebox.showerror("Please check the integrity of the Grades data.")

        # -----------------------------
        # Step 5: Save the updated CSV file
        # -----------------------------
        try:
            output_directory = os.path.dirname(sem_sheet)  # Get the directory of the master sheet
            UTD_output_file = os.path.join(output_directory, f"UTD_{batch}_{current_semester}.csv")  # Save in the same directory
            df_utd.to_csv(UTD_output_file, index=False, header=False)
            messagebox.showinfo("Success", f"UTD Output file saved as {UTD_output_file}")
        except Exception as e:
            messagebox.showerror("Error occurred while writing data to UTD file: ", e)
            messagebox.showerror("Please check if the file is open in another program and close it before running the script again.")
        messagebox.showinfo("Success","Student Data and Semester Data has been Successfully added to UTD Output file. ")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to browse files
def browse_file(entry, file_type):
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xls *.xlsx")] if file_type == "excel" else [("CSV Files", "*.csv")]
    )
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)

# Create the main window
root = tk.Tk()
root.title("Academic Sheet Automation Program (ASAP)")
root.geometry("950x600")
root.resizable(False, False) # Disable resizing of the window
root.configure(bg="#f0f8ff")  # Light blue background for a professional look

# Heading
tk.Label(root, text="Academic Sheet Automation Program (ASAP)", font=("Helvetica", 24, "bold"), bg="#f0f8ff", fg="#000080").grid(row=0, column=0, columnspan=3, pady=20)

# Subheading
tk.Label(root, text="School of Instrumentation, DAVV", font=("Helvetica", 16, "italic"), bg="#f0f8ff", fg="#000080").grid(row=1, column=0, columnspan=3, pady=10)

# Labels and entry fields for file selection
tk.Label(root, text="Select Master Sheet for Semester (Excel) :", font=("Helvetica", 12,"bold"), bg="#f0f8ff", fg="#000000").grid(row=2, column=0, padx=10, pady=10, sticky="w")
sem_sheet_path = tk.Entry(root, width=50, font=("Helvetica", 12))
sem_sheet_path.grid(row=2, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=lambda: browse_file(sem_sheet_path, "excel"), font=("Helvetica", 12), bg="#4682b4", fg="white").grid(row=2, column=2, padx=10, pady=10)

tk.Label(root, text="Select Student Details New File (Excel) :", font=("Helvetica", 12,"bold"), bg="#f0f8ff", fg="#000000").grid(row=3, column=0, padx=10, pady=10, sticky="w")
all_sem_file_path = tk.Entry(root, width=50, font=("Helvetica", 12))
all_sem_file_path.grid(row=3, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=lambda: browse_file(all_sem_file_path, "excel"), font=("Helvetica", 12), bg="#4682b4", fg="white").grid(row=3, column=2, padx=10, pady=10)

tk.Label(root, text="Select ABC_file (Excel) :", font=("Helvetica", 12,"bold"), bg="#f0f8ff", fg="#000000").grid(row=4, column=0, padx=10, pady=10, sticky="w")
abc_file_path = tk.Entry(root, width=50, font=("Helvetica", 12))
abc_file_path.grid(row=4, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=lambda: browse_file(abc_file_path, "excel"), font=("Helvetica", 12), bg="#4682b4", fg="white").grid(row=4, column=2, padx=10, pady=10)

tk.Label(root, text="Select UTD Output file (CSV) :", font=("Helvetica", 12,"bold"), bg="#f0f8ff", fg="#000000").grid(row=5, column=0, padx=10, pady=10, sticky="w")
utd_file_path = tk.Entry(root, width=50, font=("Helvetica", 12))
utd_file_path.grid(row=5, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=lambda: browse_file(utd_file_path, "csv"), font=("Helvetica", 12), bg="#4682b4", fg="white").grid(row=5, column=2, padx=10, pady=10)

# Run button
tk.Button(root, text="Run", command=run_program, font=("Helvetica", 18, "bold"), bg="#32cd32", fg="white").grid(row=6, column=1, pady=60)

# Start the GUI event loop
root.mainloop()