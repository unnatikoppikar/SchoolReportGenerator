# Report Card Generator

## Download Application
[ReportGenV2.exe](dist/ReportGenV2.exe)

## Project Overview
A Python-based application for generating report cards from Excel files using a Word template.

## Prerequisites
- Python 3.8+
- Windows Operating System (due to Win32 dependencies)

## Project Structure (Only for Convention not compulsory)
```
report_generator/
│
├── input_files/
│   ├── template-word1A.docx
│   └── student_marks.xlsx
│
├── mappings/
│   ├── I_A_mapping.json
│   ├── II_B_mapping.json
│   └── ... (other class mappings)
│
├── requirements.txt
├── main.py
└── setup.py
```

## Setup Instructions

### 1. Clone the Repository
```bash
git clone https://github.com/unnatikoppikar/SchoolReportGenerator.git
cd SchoolReportGenerator/
```

### 2. Create Virtual Environment
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/Scripts/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

## Input File Preparations

### Excel File Requirements
- Located in `input_files/`
- Filename format: `{class_name}_marks.xlsx`
- Columns must match the corresponding mapping JSON

### Mapping JSON Files
- Located in `mappings/`
- Filename format: `{Class_Name}_mapping.json`
- Maps Excel column names to report card template fields

### Word Template
- Located in `input_files/{class_name}_word.docx`
- Must have placeholders matching the mapping JSON

## Configuration

### Mapping JSON Example
```json
{
    "name": "Name", //Excel Column Name
    "percentage": "Total Percentage",
    "remark": "Remarks"
}
```

## Running the Application

### Direct Python Execution
```bash
python ReportGenV2.py
```

### Build Executable

#### Using PyInstaller
```bash
# Create executable
pyinstaller --onefile --windowed ReportGenV2.py
```
## Running Tests

To run the tests, first fill the test file paths in test_config.txt

the run this command:

```
python test.py run
```
To test only data loading from excel files run this command:
```
python test.py read_excel
```

## Troubleshooting

### Common Issues
- Ensure all required dependencies are installed
- Check Excel file format matches expected structure
- Verify mapping JSON is correctly configured

### Dependencies
Detailed dependencies are listed in `requirements.txt`. Key libraries include:
- pandas
- python-docx
- docxtpl
- pywin32
- tkinter

## Contributing
1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## Additional Notes for Setup

### Requirements.txt Content
```
pandas==1.3.5
python-docx==0.8.11
docxtpl==0.16.4
pywin32==303
openpyxl==3.0.9
tk==0.1.0
```