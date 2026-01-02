# Report Card Generator

A portable, self-contained application for generating PDF report cards from Excel data and Word templates.

**No Microsoft Word required!** Uses LibreOffice for PDF conversion.

---

## Features

- ✅ **Web-based UI** - Easy to use interface in your browser
- ✅ **Word template support** - Use your existing Word templates with placeholders
- ✅ **Excel data input** - Read student data from Excel files
- ✅ **Configurable** - Skip header rows, custom placeholder format, etc.
- ✅ **No Word dependency** - Uses LibreOffice for PDF conversion
- ✅ **Portable** - Extract and run, no installation needed

---

## Requirements

- **LibreOffice** - Required for PDF conversion
  - Download: https://www.libreoffice.org/download/
  - Or place LibreOffice Portable in the `libreoffice` folder
- **Python 3.10+** (or use bundled Python)

---

## Quick Start

### Windows

1. Install LibreOffice (or place LibreOffice Portable in `libreoffice` folder)
2. Double-click `start.bat`
3. Open browser to `http://localhost:8080`

### macOS / Linux

1. Install LibreOffice: `brew install --cask libreoffice` (macOS)
2. Run: `chmod +x start.sh && ./start.sh`
3. Open browser to `http://localhost:8080`

---

## How to Use

### 1. Prepare Your Files

**Excel File:**
- First 4 rows are headers (configurable)
- Data starts from row 5
- Each column contains different fields (Roll No, Name, Marks, etc.)

Example structure:
```
Row 1: School Name
Row 2: Class Info
Row 3: Date
Row 4: Column Headers (Roll No | Name | English | ...)
Row 5: Student 1 data
Row 6: Student 2 data
...
```

**Word Template:**
- Create your report card design in Word
- Use placeholders like `{{name}}`, `{{english}}`, `{{result}}`
- Placeholders are replaced with student data

**Mapping File (JSON):**
- Maps placeholder names to Excel columns
- Example:
```json
{
    "rollno": "A",
    "name": "B",
    "english": "C",
    "maths": "D",
    "result": "E"
}
```

### 2. Generate Report Cards

1. Open the app at `http://localhost:8080`
2. Upload your Excel file, Word template, and mapping JSON
3. Enter the class name
4. Click "Generate Report Cards"
5. Download the ZIP file with all PDFs

---

## Configuration

Edit `config/settings.json` to customize:

```json
{
    "header_rows_to_skip": 4,
    "placeholder_prefix": "{{",
    "placeholder_suffix": "}}",
    "default_null_value": "---",
    "null_indicators": ["NAN", "NONE", "NA", "NULL", ""],
    "libreoffice_timeout_seconds": 60
}
```

| Setting | Description |
|---------|-------------|
| `header_rows_to_skip` | Number of rows to skip before student data |
| `placeholder_prefix/suffix` | Format of placeholders in template |
| `default_null_value` | Value to use for empty/null cells |
| `null_indicators` | Values to treat as null |
| `libreoffice_timeout_seconds` | Timeout for PDF conversion |

---

## Sample Files

Sample files are included in the `input_files` folder:
- `sample_students.xlsx` - Example Excel file
- `sample_template.docx` - Example Word template
- `sample_mapping.json` - Example mapping file

---

## Folder Structure

```
ReportCardGenerator/
├── app/                    # Application code
│   ├── main.py             # Flask web server
│   ├── services/           # Core services
│   ├── templates/          # HTML templates
│   └── static/             # CSS/JS files
├── config/
│   └── settings.json       # Configuration
├── input_files/            # Sample files
├── tests/                  # Test files
├── libreoffice/            # (Optional) LibreOffice Portable
├── start.bat               # Windows launcher
├── start.sh                # macOS/Linux launcher
└── requirements.txt        # Python dependencies
```

---

## Troubleshooting

### "LibreOffice not found"
- Install LibreOffice from https://www.libreoffice.org/download/
- Or place LibreOffice Portable in the `libreoffice` folder

### "No student data found"
- Check that `header_rows_to_skip` in settings matches your Excel structure
- Make sure data starts after the header rows

### "Placeholder not replaced"
- Verify placeholder names in template match keys in mapping JSON
- Check that placeholders use the correct format: `{{name}}`

### PDF conversion slow
- This is normal for LibreOffice headless mode
- Each PDF takes 2-5 seconds to generate

---

## Development

### Setup

```bash
python3 -m venv venv
source venv/bin/activate  # or `venv\Scripts\activate` on Windows
pip install -r requirements.txt
```

### Run Tests

```bash
python tests/run_all_tests.py
```

### Run App

```bash
cd app
python main.py
```

---

## License

MIT License - Use freely for any purpose.

