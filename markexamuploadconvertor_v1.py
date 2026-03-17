import pandas as pd
import numpy as np
import os
from tkinter import Tk, filedialog, messagebox, Label, Button, Frame, StringVar
from tkinter.ttk import Progressbar
import threading
import warnings
import re

warnings.filterwarnings('ignore')


class MarksUploadConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Exam Marks Upload Converter")
        self.root.geometry("550x450")

        # UI Elements
        self.title_label = Label(root, text="Ad-hoc to Exam Marks Upload Converter", font=("Arial", 14, "bold"))
        self.title_label.pack(pady=10)

        self.file_frame = Frame(root)
        self.file_frame.pack(pady=5)

        # Ad-hoc file selection
        self.ad_hoc_label = Label(self.file_frame, text="1. Select Ad-hoc CSV file:", font=("Arial", 10))
        self.ad_hoc_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.ad_hoc_var = StringVar(value="Not selected")
        self.ad_hoc_file_label = Label(self.file_frame, textvariable=self.ad_hoc_var, font=("Arial", 9), fg="blue", width=30)
        self.ad_hoc_file_label.grid(row=0, column=1, padx=5, pady=5)

        self.ad_hoc_button = Button(self.file_frame, text="Browse", command=lambda: self.select_file("ad_hoc"), font=("Arial", 9))
        self.ad_hoc_button.grid(row=0, column=2, padx=5, pady=5)

        # Course Excel selection
        self.course_label = Label(self.file_frame, text="2. Select Course Excel:", font=("Arial", 10))
        self.course_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        self.course_var = StringVar(value="Not selected")
        self.course_file_label = Label(self.file_frame, textvariable=self.course_var, font=("Arial", 9), fg="blue", width=30)
        self.course_file_label.grid(row=1, column=1, padx=5, pady=5)

        self.course_button = Button(self.file_frame, text="Browse", command=lambda: self.select_file("course"), font=("Arial", 9))
        self.course_button.grid(row=1, column=2, padx=5, pady=5)

        # Convert button
        self.convert_button = Button(
            root,
            text="Convert Files",
            command=self.start_conversion,
            font=("Arial", 11),
            bg="#4CAF50",
            fg="white",
            padx=30,
            pady=5,
            state="disabled"
        )
        self.convert_button.pack(pady=10)

        self.progress_frame = Frame(root)
        self.progress_frame.pack(pady=10)

        self.progress_label = Label(self.progress_frame, text="")
        self.progress_label.pack()

        self.progress_bar = Progressbar(self.progress_frame, length=400, mode='indeterminate')

        self.status_label = Label(root, text="Please select both files", font=("Arial", 9))
        self.status_label.pack(pady=5)

        # Store file paths and data
        self.ad_hoc_path = None
        self.course_path = None
        self.course_name_list = []
        self.course_to_group = {}  # Map course name -> Section/Group Name

    # --------------------------
    # Cleaning / normalization
    # --------------------------
    def clean_course_name(self, course_name):
        """Remove suffixes like _202526T1 from course short name"""
        if not isinstance(course_name, str) or pd.isna(course_name):
            return course_name
        first_underscore = course_name.find('_')
        if first_underscore != -1:
            return course_name[:first_underscore].strip()
        return course_name.strip()

    def normalize_group_name(self, group_name):
        """
        Normalize group names for folder structure:
        - HD1A, HD1B → HD1
        - HD2A, HD2B → HD2
        """
        if not isinstance(group_name, str) or pd.isna(group_name):
            return group_name
        hd_match = re.match(r'^HD(\d+)[A-Z]$', group_name, re.IGNORECASE)
        if hd_match:
            return f"HD{hd_match.group(1)}"
        return group_name

    def clean_section_group_name(self, group_name):
        """Remove spaces and hyphens for Excel output content only"""
        if not isinstance(group_name, str) or pd.isna(group_name):
            return group_name
        return re.sub(r'[\s\-]+', '', group_name)

    def sanitize_folder_name(self, folder_name):
        """Make folder name safe for filesystem"""
        if not isinstance(folder_name, str) or pd.isna(folder_name):
            return "Unknown_Group"
        invalid_chars = r'[<>:"/\\|?*]'
        safe_name = re.sub(invalid_chars, '_', folder_name).strip('. ')
        return safe_name or "Group"

    def get_next_component(self, counter):
        """Components 2-8 then 10,11,12... (skip component 9)"""
        if counter < 7:
            return counter + 2
        return counter + 3

    # --------------------------
    # File reading
    # --------------------------
    def read_csv_with_encoding(self, file_path):
        encodings_to_try = [
            'utf-8', 'gbk', 'gb2312', 'big5', 'latin-1', 'cp1252', 'iso-8859-1',
            'utf-16', 'utf-16-le', 'utf-16-be'
        ]
        last_error = None
        for encoding in encodings_to_try:
            try:
                df = pd.read_csv(file_path, encoding=encoding)
                self.status_label.config(text=f"Successfully read with {encoding} encoding")
                return df, encoding
            except Exception as e:
                last_error = e

        try:
            df = pd.read_csv(file_path, encoding='utf-8', errors='ignore')
            return df, 'utf-8 (with errors ignored)'
        except Exception as e:
            last_error = e

        raise Exception(f"Failed to read CSV with any encoding. Last error: {last_error}")

    # --------------------------
    # Required-field validation (NEW)
    # --------------------------
    def validate_required_fields(self, df, required_cols):
        """
        Split df into:
          - valid_df: rows with all required fields present
          - invalid_df: rows with any missing required field
        Also returns missing_by_col counts.
        """
        missing_by_col = {}

        def is_missing_series(s: pd.Series) -> pd.Series:
            miss = s.isna()

            y = s[~miss].astype(str).str.strip().str.lower()
            blank_like = y.eq('') | y.eq('nan') | y.eq('none') | y.eq('null')
            miss.loc[~miss] = blank_like
            return miss

        missing_mask = pd.DataFrame(False, index=df.index, columns=required_cols)
        for col in required_cols:
            missing_mask[col] = is_missing_series(df[col])
            missing_by_col[col] = int(missing_mask[col].sum())

        invalid_rows_mask = missing_mask.any(axis=1)

        valid_df = df.loc[~invalid_rows_mask].copy()
        invalid_df = df.loc[invalid_rows_mask].copy()

        if not invalid_df.empty:
            invalid_df["Missing Fields"] = missing_mask.loc[invalid_df.index].apply(
                lambda r: ", ".join([c for c, v in r.items() if v]),
                axis=1
            )

        return valid_df, invalid_df, missing_by_col

    # --------------------------
    # UI actions
    # --------------------------
    def select_file(self, file_type):
        if file_type == "ad_hoc":
            file_path = filedialog.askopenfilename(
                title="Select Ad-hoc CSV file",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
            if file_path:
                self.ad_hoc_path = file_path
                self.ad_hoc_var.set(os.path.basename(file_path))

        elif file_type == "course":
            file_path = filedialog.askopenfilename(
                title="Select Course Excel file",
                filetypes=[("Excel files", "*.xlsx"), ("Excel files", "*.xls"), ("All files", "*.*")]
            )
            if file_path:
                self.course_path = file_path
                self.course_var.set(os.path.basename(file_path))
                self.load_course_names()

        if self.ad_hoc_path and self.course_path:
            self.convert_button.config(state="normal")
            self.status_label.config(text="Both files selected - ready to convert")

    def load_course_names(self):
        """Load course names and section/group mapping from Excel"""
        try:
            df = pd.read_excel(self.course_path)

            course_short_name_col = None
            section_group_col = None

            for col in df.columns:
                if 'course short name' in str(col).lower():
                    course_short_name_col = col
                    break
            if course_short_name_col is None:
                course_short_name_col = df.columns[0]

            for col in df.columns:
                if 'section' in str(col).lower() or 'group' in str(col).lower():
                    section_group_col = col
                    break

            self.course_name_list = []
            self.course_to_group = {}

            for _, row in df.iterrows():
                course_name = str(row[course_short_name_col]).strip()
                if pd.notna(course_name) and course_name and course_name != 'nan':
                    self.course_name_list.append(course_name)

                    if section_group_col and pd.notna(row[section_group_col]):
                        group = str(row[section_group_col]).strip()
                        if group and group != 'nan':
                            self.course_to_group[course_name] = group

            self.status_label.config(
                text=f"Loaded {len(self.course_name_list)} courses, {len(self.course_to_group)} with groups"
            )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load course names:\n{e}")
            self.course_path = None
            self.course_var.set("Not selected")

    def match_course_name(self, course_name_from_csv):
        if not isinstance(course_name_from_csv, str) or pd.isna(course_name_from_csv):
            return None
        clean_input = self.clean_course_name(course_name_from_csv)
        for valid_name in self.course_name_list:
            if valid_name.lower() == clean_input.lower():
                return valid_name
        return clean_input

    def start_conversion(self):
        if not self.ad_hoc_path or not self.course_path:
            messagebox.showerror("Error", "Please select both files first")
            return

        self.convert_button.config(state="disabled", text="Converting...")
        self.ad_hoc_button.config(state="disabled")
        self.course_button.config(state="disabled")
        self.progress_bar.pack()
        self.progress_bar.start(10)
        self.status_label.config(text="Starting conversion...")

        thread = threading.Thread(target=self.convert_file)
        thread.daemon = True
        thread.start()

    # --------------------------
    # Conversion
    # --------------------------
    def convert_file(self):
        try:
            self.root.after(0, lambda: self.status_label.config(text="Reading Ad-hoc CSV file..."))
            df, used_encoding = self.read_csv_with_encoding(self.ad_hoc_path)

            required_cols = ['Student Id', 'Student Name', 'Course Short Name',
                             'Assignment Name', 'Grade', 'Total Mark', 'Weight', 'Group']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")

            # Create ONE output folder for everything (failed rows + results)
            timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
            base_output_dir = os.path.join(os.path.dirname(self.ad_hoc_path), f"ExamMarksUpload_{timestamp}")
            os.makedirs(base_output_dir, exist_ok=True)

            # (NEW) Validate required field values; export failed rows if any
            self.root.after(0, lambda: self.status_label.config(text="Validating required fields..."))
            df_valid, df_invalid, missing_by_col = self.validate_required_fields(df, required_cols)

            fail_file = os.path.join(base_output_dir, "FailedRows_MissingRequiredFields.xlsx")
            if not df_invalid.empty:
                with pd.ExcelWriter(fail_file, engine="openpyxl") as writer:
                    df_invalid.to_excel(writer, index=False, sheet_name="FailedRows")
                    summary_df = pd.DataFrame(
                        {"Column": list(missing_by_col.keys()), "Missing Count": list(missing_by_col.values())}
                    ).sort_values("Missing Count", ascending=False)
                    summary_df.to_excel(writer, index=False, sheet_name="MissingSummary")

            df = df_valid
            if df.empty:
                raise ValueError(
                    "All rows have missing required fields.\n"
                    f"Failed rows exported to:\n{fail_file}"
                )

            # Remove attendance records (NaN-safe)
            self.root.after(0, lambda: self.status_label.config(text="Removing attendance records..."))
            assignment_series = df['Assignment Name'].fillna('').astype(str).str.lower()
            attendance_mask = assignment_series.str.contains(r'attendance|出席', regex=True, na=False)

            df = df[~attendance_mask].copy()
            attendance_removed = int(attendance_mask.sum())
            if df.empty:
                raise ValueError("No data left after removing attendance records")

            # Match course names
            self.root.after(0, lambda: self.status_label.config(text="Matching course names with Excel..."))
            df['Cleaned_Course_Name'] = df['Course Short Name'].apply(self.clean_course_name)

            course_mapping = {}
            unmatched_courses = set()

            for cleaned_name in df['Cleaned_Course_Name'].dropna().unique():
                matched_name = self.match_course_name(cleaned_name)

                if matched_name != cleaned_name:
                    course_mapping[cleaned_name] = matched_name
                else:
                    if cleaned_name in self.course_name_list:
                        course_mapping[cleaned_name] = cleaned_name
                    else:
                        course_mapping[cleaned_name] = cleaned_name
                        unmatched_courses.add(cleaned_name)

            df['Matched_Course_Name'] = df['Cleaned_Course_Name'].map(lambda x: course_mapping.get(x, x))
            df['Is_Matched'] = df['Cleaned_Course_Name'].apply(lambda x: (pd.notna(x)) and (x not in unmatched_courses))

            # Get group from Excel mapping, fallback to CSV group
            self.root.after(0, lambda: self.status_label.config(text="Getting Section/Group from Excel..."))

            def get_excel_group(row):
                course = row['Matched_Course_Name']
                if course in self.course_to_group:
                    return self.course_to_group[course]
                return row['Group']

            df['Excel_Group'] = df.apply(get_excel_group, axis=1)

            # Clean group for Excel output content
            self.root.after(0, lambda: self.status_label.config(text="Preparing group names for Excel..."))
            df['Cleaned_Group_For_Excel'] = df['Excel_Group'].apply(self.clean_section_group_name)

            # Normalize group for folder structure (based on CSV group)
            self.root.after(0, lambda: self.status_label.config(text="Normalizing group names for folders..."))
            df['Normalized_Group_For_Folder'] = df['Group'].apply(self.normalize_group_name)

            # Process groups
            normalized_groups = df['Normalized_Group_For_Folder'].dropna().unique()

            files_created = []
            matched_count = 0
            unmatched_count = 0

            for group_idx, normalized_group in enumerate(normalized_groups):
                self.root.after(
                    0,
                    lambda g=normalized_group, idx=group_idx, total=len(normalized_groups):
                    self.progress_label.config(text=f"Processing Group {idx+1}/{total}: {g}")
                )

                group_df = df[df['Normalized_Group_For_Folder'] == normalized_group].copy()
                matched_df = group_df[group_df['Is_Matched'] == True].copy()
                unmatched_df = group_df[group_df['Is_Matched'] == False].copy()

                if not matched_df.empty:
                    for course in matched_df['Matched_Course_Name'].dropna().unique():
                        course_df = matched_df[matched_df['Matched_Course_Name'] == course].copy()
                        files = self.process_course_group(course_df, course, normalized_group, base_output_dir, is_matched=True)
                        files_created.extend(files)
                        matched_count += len(files)

                if not unmatched_df.empty:
                    for course in unmatched_df['Matched_Course_Name'].dropna().unique():
                        course_df = unmatched_df[unmatched_df['Matched_Course_Name'] == course].copy()
                        files = self.process_course_group(course_df, course, normalized_group, base_output_dir, is_matched=False)
                        files_created.extend(files)
                        unmatched_count += len(files)

            self.root.after(
                0,
                self.conversion_complete,
                files_created,
                base_output_dir,
                matched_count,
                unmatched_count,
                unmatched_courses,
                used_encoding,
                attendance_removed,
                fail_file if not df_invalid.empty else None
            )

        except Exception as e:
            self.root.after(0, self.conversion_error, str(e))

    def process_course_group(self, course_df, course_name, normalized_group, base_output_dir, is_matched=True):
        files_created = []
        if course_df.empty:
            return files_created

        students = course_df['Student Id'].dropna().unique()
        upload_data = []

        for student_id in students:
            student_df = course_df[course_df['Student Id'] == student_id].copy().reset_index()

            # NaN-safe exam detection
            names = student_df['Assignment Name'].fillna('').astype(str).str.lower()
            exam_mask = (names.eq('exam')) | (names.str.contains(r'\bexam\b', regex=True, na=False))

            exam_df = student_df[exam_mask].copy()
            non_exam_df = student_df[~exam_mask].copy()

            non_exam_df['Weight_Priority'] = non_exam_df['Weight'].fillna(0).astype(float) > 0
            non_exam_df = non_exam_df.sort_values(by=['Weight_Priority', 'index'], ascending=[False, True])

            # Exam -> Component 1
            for _, row in exam_df.iterrows():
                upload_data.append({
                    'Student ID': student_id,
                    'Student Name': row['Student Name'],
                    'Course Short Name': course_name,
                    'Component Name': 'Component 1',
                    'Component Description': row['Assignment Name'],
                    'Grade': row['Grade'],
                    'Total Marks': row['Total Mark'],
                    'Weightage': row['Weight'],
                    'Section / Group Name': row['Cleaned_Group_For_Excel'],
                    'Mandatory': 'No'
                })

            # Others -> Component 2-8 then 10...
            non_exam_counter = 0
            for _, row in non_exam_df.iterrows():
                component_num = self.get_next_component(non_exam_counter)
                upload_data.append({
                    'Student ID': student_id,
                    'Student Name': row['Student Name'],
                    'Course Short Name': course_name,
                    'Component Name': f'Component {component_num}',
                    'Component Description': row['Assignment Name'],
                    'Grade': row['Grade'],
                    'Total Marks': row['Total Mark'],
                    'Weightage': row['Weight'],
                    'Section / Group Name': row['Cleaned_Group_For_Excel'],
                    'Mandatory': 'No'
                })
                non_exam_counter += 1

        upload_df = pd.DataFrame(upload_data)

        # Output folder
        if is_matched:
            group_folder = os.path.join(base_output_dir, self.sanitize_folder_name(normalized_group))
        else:
            group_folder = os.path.join(base_output_dir, "Unknown", self.sanitize_folder_name(normalized_group))

        os.makedirs(group_folder, exist_ok=True)

        safe_course_name = self.sanitize_folder_name(course_name)
        filename = f"{safe_course_name}.xlsx"
        output_path = os.path.join(group_folder, filename)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            upload_df.to_excel(writer, sheet_name='ExamMarksUpload', index=False)

            worksheet = writer.sheets['ExamMarksUpload']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    val = "" if cell.value is None else str(cell.value)
                    max_length = max(max_length, len(val))
                worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)

        folder_prefix = "Unknown/" if not is_matched else ""
        files_created.append(f"{folder_prefix}{self.sanitize_folder_name(normalized_group)}/{filename}")
        return files_created

    # --------------------------
    # UI results
    # --------------------------
    def conversion_complete(
        self,
        files_created,
        output_dir,
        matched_count,
        unmatched_count,
        unmatched_courses,
        used_encoding,
        attendance_removed,
        fail_file_path
    ):
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.progress_label.config(text="")
        self.convert_button.config(state="normal", text="Convert Files")
        self.ad_hoc_button.config(state="normal")
        self.course_button.config(state="normal")
        self.status_label.config(text="Conversion complete!")

        summary = f"✓ Created {len(files_created)} file(s)\n"
        summary += f"✓ Attendance records removed: {attendance_removed}\n"
        summary += f"✓ Matched courses: {matched_count}\n"
        summary += f"✓ Unmatched courses: {unmatched_count}\n"
        summary += f"✓ Output folder: {output_dir}\n"

        if fail_file_path:
            summary += "⚠️ Some rows had missing required fields and were exported:\n"
            summary += f"   {fail_file_path}\n\n"

        if unmatched_courses:
            summary += "Unmatched courses (placed in 'Unknown' folder):\n"
            for course in list(unmatched_courses)[:10]:
                summary += f"   • {course}\n"
            if len(unmatched_courses) > 10:
                summary += f"   ... and {len(unmatched_courses) - 10} more\n"
            summary += "\n"

        messagebox.showinfo("Success", summary)

    def conversion_error(self, error_msg):
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.progress_label.config(text="")
        self.convert_button.config(state="normal", text="Convert Files")
        self.ad_hoc_button.config(state="normal")
        self.course_button.config(state="normal")
        self.status_label.config(text="Error occurred!")
        messagebox.showerror("Error", f"Conversion failed:\n{error_msg}")


def main():
    root = Tk()
    app = MarksUploadConverter(root)
    root.mainloop()


if __name__ == "__main__":
    print("=" * 70)
    print("EXAM MARKS UPLOAD CONVERTER")
    print("=" * 70)
    print("\nFeatures:")
    print("• Both files are required (Ad-hoc CSV and Course Excel)")
    print("• Attendance records completely removed")
    print("• Rows with missing required fields are exported to FailedRows_MissingRequiredFields.xlsx")
    print("• Component mapping: Exam=Comp1, Others=Comp2-8, then Comp10,11,12...")
    print("• Section/Group Name from Excel file (spaces and hyphens removed)")
    print("• Folder names keep original Ad-hoc Group names")
    print("• Unmatched courses go to 'Unknown' folder")
    print("-" * 40)
    print("\nRunning application now...")
    print("=" * 70)

    main()