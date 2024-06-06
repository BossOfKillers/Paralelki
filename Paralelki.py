import pandas as pd
import os

# Define the class capacity
class_capacity = {
    "МАТ": 26,
    "БЕЛ": 26,
    "КМИТ": 26,
    "НЕ": 26,
    "ФЕ": 26,
    "РЕ": 26
}

# Read student data from Excel file
file_dir = './mnt/data'
input_path = os.path.join(file_dir, 'student_data.xlsx')
students_df = pd.read_excel(input_path)

# Initialize class column
students_df['class'] = 'Not Assigned'

# Extract preferences from the respective columns
preferences_columns = ["МАТ", "БЕЛ", "КМИТ", "НЕ", "ФЕ", "РЕ"]

# Convert preference columns to integers
for col in preferences_columns:
    students_df[col] = pd.to_numeric(students_df[col], errors='coerce').fillna(0).astype(int)

# Calculate total points for math and BEL
students_df['Total_Math'] = students_df['НВО МАТ'] + students_df['МАТ ИУЧ']
students_df['Total_BEL'] = students_df['НВО БЕЛ'] + students_df['БЕЛ ИУЧ']

# Function to assign students to classes based on preferences and total scores
def assign_students_to_classes(students_df, class_capacity):
    class_assignments = {key: [] for key in class_capacity.keys()}
    sorted_students_df = students_df.copy()

    # Sort students by their scores and preferences
    sorted_students_df = students_df.sort_values(by=['УСПЕХ'], ascending=False).copy()

    for index, student in sorted_students_df.iterrows():
        for preference in range(1, 7):
            for class_name in preferences_columns:
                if student[class_name] == preference and len(class_assignments[class_name]) < class_capacity[class_name]:
                    if class_name in ["МАТ", "КМИТ"] and student["Total_Math"] > 0:
                        class_assignments[class_name].append(student["Име, презиме, фамилия"])
                        students_df.at[index, 'class'] = class_name
                        break
                    elif class_name == "БЕЛ" and student["Total_BEL"] > 0:
                        class_assignments[class_name].append(student["Име, презиме, фамилия"])
                        students_df.at[index, 'class'] = class_name
                        break
                    elif class_name not in ["МАТ", "КМИТ", "БЕЛ"]:
                        class_assignments[class_name].append(student["Име, презиме, фамилия"])
                        students_df.at[index, 'class'] = class_name
                        break
            if students_df.at[index, 'class'] != 'Not Assigned':
                break

    return class_assignments

# Assign students
assignments = assign_students_to_classes(students_df, class_capacity)

# Calculate total number of students assigned to each class
class_totals = {class_name: len(students) for class_name, students in assignments.items()}

# Output path
output_path = os.path.join(file_dir, 'student_assignments.xlsx')

# Ensure the directory exists
os.makedirs(file_dir, exist_ok=True)

# Save the DataFrame to an Excel file
students_df.to_excel(output_path, index=False)

# Save the class totals to the same Excel file in a new sheet
with pd.ExcelWriter(output_path, engine='openpyxl', mode='a') as writer:
    totals_df = pd.DataFrame(list(class_totals.items()), columns=['Class', 'Total Students'])
    totals_df.to_excel(writer, sheet_name='Class Totals', index=False)

print("Assignments completed and saved to Excel file.")
