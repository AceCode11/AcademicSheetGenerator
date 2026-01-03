import sys
import pandas as pd
from io import BytesIO
from datetime import datetime
import zipfile
import re

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QHBoxLayout, QPushButton, QLabel, QFileDialog,
    QComboBox, QCheckBox, QSpinBox, QTextEdit,
    QLineEdit, QMessageBox, QTabWidget, QScrollArea,
    QGroupBox, QRadioButton
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

# ----- PDF generation (ReportLab) -----
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Paragraph, Spacer
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.colors import HexColor


# =====================================================================
#                    UTILITY FUNCTIONS
# =====================================================================

def extract_department_from_sheet_name(sheet_name):
    """Extract department name dynamically from sheet name"""
    sheet_lower = sheet_name.lower()

    dept_patterns = {
        r'\b(bsc\s*it|b\.sc\s*it|information\s*technology)\b': 'IT',
        r'\b(bsc\s*cs|b\.sc\s*cs|computer\s*science)\b': 'CS',
        r'\b(bsc\s*ds|b\.sc\s*ds|data\s*science)\b': 'DS',
        r'\b(bcom|b\.com|commerce)\b': 'Commerce',
        r'\b(bms|b\.m\.s|management\s*studies)\b': 'BMS',
        r'\b(baf|b\.a\.f|accounting\s*finance)\b': 'BAF'
    }

    for pattern, dept_code in dept_patterns.items():
        if re.search(pattern, sheet_lower):
            return dept_code

    custom_match = re.search(r'[fy|sy|ty][-_\s]*([a-z]+)', sheet_lower)
    if custom_match:
        return custom_match.group(1).upper()

    standalone_match = re.search(r'\b([a-z]{2,4})\b', sheet_lower)
    if standalone_match and len(standalone_match.group(1)) <= 4:
        return standalone_match.group(1).upper()

    return None


def get_all_departments_from_sheets(sheet_info):
    """Extract all unique departments from sheet names"""
    departments = set()

    for sheet_name, info in sheet_info.items():
        if 'error' not in info and info['rows'] > 0:
            detected_dept = extract_department_from_sheet_name(sheet_name)
            if detected_dept:
                departments.add(detected_dept)

    common_depts = ["IT", "CS", "DS", "Commerce", "BMS", "BAF"]
    departments.update(common_depts)

    return sorted(list(departments))


def get_full_department_name(dept_code):
    """Get full department name from code"""
    dept_mapping = {
        'IT': 'Information Technology',
        'CS': 'Computer Science',
        'DS': 'Data Science',
        'COMMERCE': 'Commerce',
        'BMS': 'Management Studies',
        'BAF': 'Accounting & Finance'
    }
    return dept_mapping.get(dept_code.upper(), dept_code)


def extract_year_from_sheet_name(sheet_name):
    """Extract year information from sheet name"""
    sheet_lower = sheet_name.lower()

    year_patterns = {
        r'\b(fy|first\s*year|1st\s*year)\b': 'FY',
        r'\b(sy|second\s*year|2nd\s*year)\b': 'SY',
        r'\b(ty|third\s*year|3rd\s*year)\b': 'TY'
    }

    for pattern, year_code in year_patterns.items():
        if year_code in sheet_lower:
            return year_code

    return None


def clean_dataframe(df):
    """Find and standardize Roll No & Name columns"""
    original_columns = df.columns.tolist()
    df_lower = df.copy()
    df_lower.columns = [col.strip().lower() for col in df_lower.columns]

    col_map = {}
    roll_keywords = ['roll', 'rollno', 'roll_no', 'roll no', 'student_id',
                     'id', 'registration', 'reg_no', 'number']
    name_keywords = ['name', 'student_name', 'student name',
                     'full_name', 'full name', 'student']

    for i, col in enumerate(df_lower.columns):
        col_clean = col.replace('_', '').replace(' ', '').replace('-', '').lower()
        for keyword in roll_keywords:
            keyword_clean = keyword.replace('_', '').replace(' ', '').replace('-', '').lower()
            if keyword_clean in col_clean or col_clean in keyword_clean:
                col_map['Roll No'] = original_columns[i]
                break
        if 'Roll No' in col_map:
            break

    for i, col in enumerate(df_lower.columns):
        col_clean = col.replace('_', '').replace(' ', '').replace('-', '').lower()
        for keyword in name_keywords:
            keyword_clean = keyword.replace('_', '').replace(' ', '').replace('-', '').lower()
            if keyword_clean in col_clean or col_clean in keyword_clean:
                col_map['Name'] = original_columns[i]
                break
        if 'Name' in col_map:
            break

    if 'Roll No' not in col_map or 'Name' not in col_map:
        available_cols = ', '.join(original_columns)
        missing = []
        if 'Roll No' not in col_map:
            missing.append("Roll number column")
        if 'Name' not in col_map:
            missing.append("Name column")
        return None, f"Could not find: {' and '.join(missing)}. Available: {available_cols}"

    df_clean = df[[col_map['Roll No'], col_map['Name']]].copy()
    df_clean.columns = ['Roll No', 'Name']
    original_count = len(df_clean)

    df_clean = df_clean.dropna(subset=['Roll No', 'Name'])
    df_clean = df_clean[df_clean['Roll No'].astype(str).str.strip() != '']
    df_clean = df_clean[df_clean['Name'].astype(str).str.strip() != '']
    df_clean = df_clean.reset_index(drop=True)

    removed_count = original_count - len(df_clean)
    found_msg = f"Found columns: '{col_map['Roll No']}' → Roll No, '{col_map['Name']}' → Name"
    clean_msg = (f"Cleaned: {removed_count} incomplete rows removed"
                 if removed_count > 0 else "No incomplete rows")

    return df_clean, f"{found_msg}. {clean_msg}"


def get_sheet_info(file_path):
    """Read workbook and basic info for each sheet"""
    try:
        excel_file = pd.ExcelFile(file_path)
        sheet_info = {}
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                detected_dept = extract_department_from_sheet_name(sheet_name)
                detected_year = extract_year_from_sheet_name(sheet_name)

                sheet_info[sheet_name] = {
                    'rows': len(df),
                    'columns': len(df.columns),
                    'detected_department': detected_dept,
                    'detected_year': detected_year,
                    'column_names': df.columns.tolist()
                }
            except Exception as e:
                sheet_info[sheet_name] = {
                    'error': str(e),
                    'detected_department': None,
                    'detected_year': None,
                    'rows': 0,
                    'columns': 0
                }
        return sheet_info
    except Exception:
        return None


def parse_custom_columns(columns_input):
    """Parse comma-separated column names"""
    if not columns_input.strip():
        return []
    columns = [col.strip() for col in columns_input.split(',')]
    return [col for col in columns if col]


# =====================================================================
#                    PDF GENERATION HELPERS
# =====================================================================

def create_pdf_file(dataframe, sheet_title, college_name, department):
    """
    Create a landscape A4 PDF with:
    - Polished header (college, department, title)
    - Class / Subject / Semester / Code line
    - Table with evenly spaced columns
    """
    buffer = BytesIO()
    # Landscape A4
    page_size = landscape(A4)
    left_margin = right_margin = top_margin = bottom_margin = 36  # points
    doc = SimpleDocTemplate(
        buffer,
        pagesize=page_size,
        leftMargin=left_margin,
        rightMargin=right_margin,
        topMargin=top_margin,
        bottomMargin=bottom_margin
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "CollegeTitle",
        parent=styles["Title"],
        alignment=1,  # center
        fontSize=18,
        leading=22
    )
    dept_style = ParagraphStyle(
        "DeptTitle",
        parent=styles["Heading2"],
        alignment=1,
        fontSize=13,
        leading=16
    )
    sheet_title_style = ParagraphStyle(
        "SheetTitle",
        parent=styles["Heading3"],
        alignment=1,
        fontSize=12,
        leading=15
    )
    normal_center = ParagraphStyle(
        "NormalCenter",
        parent=styles["Normal"],
        alignment=1,
        fontSize=9
    )
    normal_left = ParagraphStyle(
        "NormalLeft",
        parent=styles["Normal"],
        alignment=0,
        fontSize=9
    )

    elements = []

    # ---- Header ----
    if college_name.strip():
        elements.append(Paragraph(college_name.strip(), title_style))
        elements.append(Spacer(1, 4))

    if department:
        dept_full = get_full_department_name(department)
        elements.append(Paragraph(f"Department of {dept_full}", dept_style))
        elements.append(Spacer(1, 2))

    elements.append(Paragraph(sheet_title, sheet_title_style))
    elements.append(Spacer(1, 6))

    # small horizontal divider
    divider_table = Table(
        [[""]],
        colWidths=[(page_size[0] - left_margin - right_margin)]
    )
    divider_table.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -1), 0.6, colors.black),
    ]))
    elements.append(divider_table)
    elements.append(Spacer(1, 4))

    # Class / Subject / Semester / Code line (text only, like your sheet)
    class_line = "Class: ____________    Subject: ______________________    Semester: ______    Code: ______"
    elements.append(Paragraph(class_line, normal_left))
    elements.append(Spacer(1, 8))

    # ---- Table ----
    if dataframe.empty:
        elements.append(Paragraph("No data", normal_center))
    else:
        df_str = dataframe.astype(str)
        data = [list(df_str.columns)] + df_str.values.tolist()

        # Even column widths across full printable width
        printable_width = page_size[0] - left_margin - right_margin
        num_cols = len(df_str.columns)
        col_width = printable_width / float(num_cols)
        col_widths = [col_width] * num_cols

        table = Table(data, colWidths=col_widths, repeatRows=1)

        header_color = HexColor("#366092")

        table.setStyle(TableStyle([
            # Header row
            ("BACKGROUND", (0, 0), (-1, 0), header_color),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 9),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 5),

            # Body
            ("ALIGN", (0, 1), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 1), (-1, -1), 8),

            # Grid
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),

            # Row height slightly larger
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))

        elements.append(table)

    doc.build(elements)
    buffer.seek(0)
    return buffer


def create_zip_file(files_dict):
    """Pack multiple PDFs into a ZIP"""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for filename, file_data in files_dict.items():
            zip_file.writestr(filename, file_data.getvalue())
    zip_buffer.seek(0)
    return zip_buffer


# =====================================================================
#                    SHEET GENERATORS (DATAFRAME)
# =====================================================================

def generate_assignment_sheet(attendance_df, year_type, num_assignments=None):
    """Generate assignment sheet dataframe"""
    assignment_df = attendance_df[['Roll No', 'Name']].copy()

    if num_assignments:
        for i in range(1, num_assignments + 1):
            assignment_df[f'Assignment {i}'] = ''
    else:
        if year_type == "FY/SY":
            assignment_df['Assignment 1'] = ''
            assignment_df['Assignment 2'] = ''
        else:
            for i in range(1, 6):
                assignment_df[f'Assignment {i}'] = ''

    assignment_df['Date Submitted'] = ''
    return assignment_df


def generate_practical_journal_sheet(attendance_df):
    practical_df = attendance_df[['Roll No', 'Name']].copy()
    for i in range(1, 11):
        practical_df[f'Prac No {i}'] = ''
    practical_df['Total Completed'] = ''
    return practical_df


def generate_journal_completion_sheet(attendance_df):
    journal_df = attendance_df[['Roll No', 'Name']].copy()
    journal_df['Certificate Issued'] = ''
    journal_df['Date'] = ''
    return journal_df


def create_custom_test_sheet(attendance_df, test_name, semester, custom_columns):
    """Create test sheet with custom columns"""
    test_df = attendance_df[['Roll No', 'Name']].copy()

    for column in custom_columns:
        if column.strip():
            test_df[column.strip()] = ''

    common_columns = ['Date', 'Remarks']
    for col in common_columns:
        if col not in test_df.columns:
            test_df[col] = ''

    return test_df


# =====================================================================
#                    MAIN APPLICATION (PyQt)
# =====================================================================

class AcademicManagementSystem(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Academic Management System")
        self.setGeometry(100, 100, 1200, 800)

        # Data storage
        self.uploaded_file_path = None
        self.sheet_info = None
        self.cleaned_df = None

        self.init_ui()

    # ------------------------------------------------------------
    # UI SETUP
    # ------------------------------------------------------------
    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Title
        title_label = QLabel("📚 Academic Management System")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)

        # Tabs
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        self.tabs.addTab(self.create_home_tab(), "🏠 Home")
        self.tabs.addTab(self.create_sheet_generator_tab(), "📄 Sheet Generator")

    def create_home_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        welcome_label = QLabel("Welcome to Academic Management System")
        welcome_font = QFont()
        welcome_font.setPointSize(14)
        welcome_font.setBold(True)
        welcome_label.setFont(welcome_font)
        layout.addWidget(welcome_label)

        info_text = QTextEdit()
        info_text.setReadOnly(True)
        info_text.setHtml("""
        <h3>Sheet Generator</h3>
        <p>Generate multiple academic sheets from your Excel data, now directly in <b>PDF format</b>:</p>
        <ul>
            <li>Assignment Sheets</li>
            <li>Internal / Assessment Sheets</li>
            <li>Practical / Journal tracking Sheets</li>
            <li>Journal Completion Certificates</li>
            <li>Clean Class Lists</li>
        </ul>
        <p><b>Use the "Sheet Generator" tab to get started.</b></p>
        """)
        layout.addWidget(info_text)

        return widget

    def create_sheet_generator_tab(self):
        widget = QWidget()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(widget)

        layout = QVBoxLayout(widget)

        # College Name
        college_group = QGroupBox("College Configuration")
        college_layout = QVBoxLayout()
        self.college_name_input = QLineEdit()
        self.college_name_input.setText("Thakur Shyamnarayan Degree College")


        college_layout.addWidget(QLabel("College Name:"))
        college_layout.addWidget(self.college_name_input)
        college_group.setLayout(college_layout)
        layout.addWidget(college_group)

        # File Upload
        file_group = QGroupBox("Step 1: Upload Excel File")
        file_layout = QVBoxLayout()

        file_btn_layout = QHBoxLayout()
        self.upload_btn = QPushButton("Choose Excel File")
        self.upload_btn.clicked.connect(self.upload_file_sheet_gen)
        file_btn_layout.addWidget(self.upload_btn)

        self.file_label = QLabel("No file selected")
        file_btn_layout.addWidget(self.file_label)
        file_btn_layout.addStretch()

        file_layout.addLayout(file_btn_layout)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # Sheet Selection
        sheet_group = QGroupBox("Step 2: Select Sheet and Configuration")
        sheet_layout = QVBoxLayout()

        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_selected)
        sheet_layout.addWidget(QLabel("Select Sheet:"))
        sheet_layout.addWidget(self.sheet_combo)

        self.dept_combo = QComboBox()
        sheet_layout.addWidget(QLabel("Select Department:"))
        sheet_layout.addWidget(self.dept_combo)

        self.custom_dept_check = QCheckBox("Use Custom Department")
        self.custom_dept_check.stateChanged.connect(self.toggle_custom_dept)
        sheet_layout.addWidget(self.custom_dept_check)

        self.custom_dept_input = QLineEdit()
        self.custom_dept_input.setPlaceholderText("Enter custom department code")
        self.custom_dept_input.setEnabled(False)
        sheet_layout.addWidget(self.custom_dept_input)

        sheet_group.setLayout(sheet_layout)
        layout.addWidget(sheet_group)

        # Sheet Generation Options
        options_group = QGroupBox("Step 3: Select Sheets to Generate (PDF)")
        options_layout = QVBoxLayout()

        # Assignment options
        self.assignment_check = QCheckBox("Generate Assignment Sheet")
        options_layout.addWidget(self.assignment_check)

        assignment_options = QHBoxLayout()
        self.assignment_year_fy = QRadioButton("FY/SY")
        self.assignment_year_ty = QRadioButton("TY")
        self.assignment_year_fy.setChecked(True)
        assignment_options.addWidget(QLabel("Year Type:"))
        assignment_options.addWidget(self.assignment_year_fy)
        assignment_options.addWidget(self.assignment_year_ty)
        assignment_options.addStretch()
        options_layout.addLayout(assignment_options)

        self.custom_assignment_check = QCheckBox("Custom Number of Assignments")
        options_layout.addWidget(self.custom_assignment_check)

        self.num_assignments_spin = QSpinBox()
        self.num_assignments_spin.setRange(1, 15)
        self.num_assignments_spin.setValue(5)
        self.num_assignments_spin.setEnabled(False)
        self.custom_assignment_check.stateChanged.connect(
            lambda: self.num_assignments_spin.setEnabled(self.custom_assignment_check.isChecked())
        )
        options_layout.addWidget(self.num_assignments_spin)

        # Test / Assessment sheet
        self.test_check = QCheckBox("Generate Test / Assessment Sheet")
        options_layout.addWidget(self.test_check)

        self.test_name_input = QLineEdit()
        self.test_name_input.setPlaceholderText("Test Name (e.g., Internal Assessment, Unit Test 1)")
        options_layout.addWidget(self.test_name_input)

        self.test_columns_input = QTextEdit()
        self.test_columns_input.setPlaceholderText(
            "Enter column names (comma-separated)\n"
            "Example: Technique, Practical, Assignment, Date, Remarks"
        )
        self.test_columns_input.setMaximumHeight(80)
        options_layout.addWidget(self.test_columns_input)

        # Practical/Journal Sheet
        self.practical_check = QCheckBox("Generate Practical / Journal Sheet")
        options_layout.addWidget(self.practical_check)

        # Certificate
        self.certificate_check = QCheckBox("Generate Journal Completion Certificate")
        options_layout.addWidget(self.certificate_check)

        # Class List
        self.class_list_check = QCheckBox("Generate Clean Class List")
        options_layout.addWidget(self.class_list_check)

        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        # Generate Button
        self.generate_btn = QPushButton("Generate Selected Sheets (PDF)")
        self.generate_btn.clicked.connect(self.generate_sheets)
        self.generate_btn.setMinimumHeight(40)
        layout.addWidget(self.generate_btn)

        layout.addStretch()

        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.addWidget(scroll)
        return container

    # ------------------------------------------------------------
    # SHEET GENERATOR METHODS
    # ------------------------------------------------------------
    def upload_file_sheet_gen(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.uploaded_file_path = file_path
            self.file_label.setText(f"Selected: {file_path.split('/')[-1]}")

            self.sheet_info = get_sheet_info(file_path)
            if self.sheet_info:
                self.sheet_combo.clear()
                self.sheet_combo.addItem("-- Select Sheet --")
                valid_sheets = [
                    name for name, info in self.sheet_info.items()
                    if 'error' not in info and info['rows'] > 0
                ]
                self.sheet_combo.addItems(valid_sheets)

                all_depts = get_all_departments_from_sheets(self.sheet_info)
                self.dept_combo.clear()
                self.dept_combo.addItem("-- Select Department --")
                self.dept_combo.addItems(all_depts)

                QMessageBox.information(
                    self, "Success",
                    f"File loaded successfully! Found {len(valid_sheets)} sheet(s)"
                )
            else:
                QMessageBox.critical(self, "Error", "Failed to read Excel file")

    def on_sheet_selected(self, sheet_name):
        if sheet_name and sheet_name != "-- Select Sheet --" and self.sheet_info:
            sheet_data = self.sheet_info[sheet_name]
            try:
                df = pd.read_excel(self.uploaded_file_path, sheet_name=sheet_name)
                cleaned_df, message = clean_dataframe(df)

                if cleaned_df is not None:
                    self.cleaned_df = cleaned_df

                    detected_dept = sheet_data.get('detected_department')
                    if detected_dept:
                        index = self.dept_combo.findText(detected_dept)
                        if index >= 0:
                            self.dept_combo.setCurrentIndex(index)

                    QMessageBox.information(
                        self, "Sheet Loaded",
                        f"{message}\nTotal students: {len(cleaned_df)}"
                    )
                else:
                    QMessageBox.warning(self, "Warning", message)
                    self.cleaned_df = None
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error processing sheet: {str(e)}")
                self.cleaned_df = None

    def toggle_custom_dept(self, state):
        self.custom_dept_input.setEnabled(state == 2)

    def generate_sheets(self):
        # Basic validation
        college_name = self.college_name_input.text().strip()
        if not college_name:
            QMessageBox.warning(self, "Warning", "Please enter college name")
            return

        if self.cleaned_df is None:
            QMessageBox.warning(self, "Warning", "Please select a valid sheet first")
            return

        # Department
        if self.custom_dept_check.isChecked() and self.custom_dept_input.text().strip():
            department = self.custom_dept_input.text().strip().upper()
        else:
            department = self.dept_combo.currentText()
            if department == "-- Select Department --":
                QMessageBox.warning(self, "Warning", "Please select a department")
                return

        selected_sheet = self.sheet_combo.currentText()

        files_dict = {}
        generated_count = 0

        try:
            # Assignment PDF
            if self.assignment_check.isChecked():
                year_type = "FY/SY" if self.assignment_year_fy.isChecked() else "TY"
                num_assignments = (
                    self.num_assignments_spin.value()
                    if self.custom_assignment_check.isChecked()
                    else None
                )

                assignment_data = generate_assignment_sheet(
                    self.cleaned_df, year_type, num_assignments
                )
                if num_assignments:
                    sheet_title = f"Assignment Sheet ({num_assignments} Assignments)"
                else:
                    sheet_title = f"Assignment Sheet ({year_type})"

                pdf_buffer = create_pdf_file(
                    assignment_data,
                    sheet_title,
                    college_name,
                    department
                )

                filename = (
                    f"Assignment_Sheet_{year_type}_{selected_sheet}_"
                    f"{datetime.now().strftime('%Y%m%d')}.pdf"
                )
                files_dict[filename] = pdf_buffer
                generated_count += 1

            # Test / Assessment PDF
            if self.test_check.isChecked():
                test_name = self.test_name_input.text().strip()
                test_columns = self.test_columns_input.toPlainText().strip()

                if not (test_name and test_columns):
                    QMessageBox.warning(
                        self, "Warning",
                        "Please enter Test Name and Columns for Test Sheet"
                    )
                else:
                    columns = parse_custom_columns(test_columns)
                    if not columns:
                        QMessageBox.warning(
                            self, "Warning",
                            "Please enter at least one valid column name for Test Sheet"
                        )
                    else:
                        test_data = create_custom_test_sheet(
                            self.cleaned_df,
                            test_name,
                            "I",
                            columns
                        )
                        sheet_title = f"{test_name} - Assessment Sheet"

                        pdf_buffer = create_pdf_file(
                            test_data,
                            sheet_title,
                            college_name,
                            department
                        )

                        safe_name = test_name.replace(' ', '_')
                        filename = (
                            f"{safe_name}_{selected_sheet}_"
                            f"{datetime.now().strftime('%Y%m%d')}.pdf"
                        )
                        files_dict[filename] = pdf_buffer
                        generated_count += 1

            # Practical / Journal PDF
            if self.practical_check.isChecked():
                practical_data = generate_practical_journal_sheet(self.cleaned_df)
                sheet_title = "Practical / Journal Sheet"

                pdf_buffer = create_pdf_file(
                    practical_data,
                    sheet_title,
                    college_name,
                    department
                )

                filename = (
                    f"Practical_Journal_{selected_sheet}_"
                    f"{datetime.now().strftime('%Y%m%d')}.pdf"
                )
                files_dict[filename] = pdf_buffer
                generated_count += 1

            # Journal Completion Certificate PDF
            if self.certificate_check.isChecked():
                journal_data = generate_journal_completion_sheet(self.cleaned_df)
                sheet_title = "Journal Completion Certificate"

                pdf_buffer = create_pdf_file(
                    journal_data,
                    sheet_title,
                    college_name,
                    department
                )

                filename = (
                    f"Journal_Completion_Certificate_{selected_sheet}_"
                    f"{datetime.now().strftime('%Y%m%d')}.pdf"
                )
                files_dict[filename] = pdf_buffer
                generated_count += 1

            # Class List PDF
            if self.class_list_check.isChecked():
                class_list_data = self.cleaned_df[['Roll No', 'Name']].copy()
                sheet_title = "Class List"

                pdf_buffer = create_pdf_file(
                    class_list_data,
                    sheet_title,
                    college_name,
                    department
                )

                filename = (
                    f"Class_List_{selected_sheet}_"
                    f"{datetime.now().strftime('%Y%m%d')}.pdf"
                )
                files_dict[filename] = pdf_buffer
                generated_count += 1

            if generated_count == 0:
                QMessageBox.warning(
                    self, "Warning",
                    "Please select at least one sheet to generate"
                )
                return

            # Saving
            if generated_count == 1:
                single_filename = list(files_dict.keys())[0]
                single_file = list(files_dict.values())[0]

                save_path, _ = QFileDialog.getSaveFileName(
                    self,
                    "Save PDF File",
                    single_filename,
                    "PDF Files (*.pdf)"
                )
                if save_path:
                    with open(save_path, 'wb') as f:
                        f.write(single_file.getvalue())
                    QMessageBox.information(
                        self, "Success",
                        "PDF file saved successfully!"
                    )
            else:
                zip_filename = (
                    f"Academic_Sheets_{department}_{selected_sheet}_"
                    f"{datetime.now().strftime('%Y%m%d_%H%M')}.zip"
                )
                save_path, _ = QFileDialog.getSaveFileName(
                    self,
                    "Save ZIP of PDFs",
                    zip_filename,
                    "ZIP Files (*.zip)"
                )
                if save_path:
                    zip_file = create_zip_file(files_dict)
                    with open(save_path, 'wb') as f:
                        f.write(zip_file.getvalue())
                    QMessageBox.information(
                        self, "Success",
                        f"Successfully generated {generated_count} PDF sheet(s)!"
                    )

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error generating sheets: {str(e)}")


# =====================================================================
#                    MAIN
# =====================================================================

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = AcademicManagementSystem()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
