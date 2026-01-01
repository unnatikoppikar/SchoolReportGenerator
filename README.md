# School Report Card Generator

A cross-platform application for generating PDF report cards from Excel data.

## Download

**[Download SchoolReportGenerator.exe](https://github.com/utkarsh-koppikar/SchoolReportGenerator/releases/latest/download/SchoolReportGenerator.exe)** - Windows executable (self-contained, no installation required)

[View all releases](https://github.com/utkarsh-koppikar/SchoolReportGenerator/releases)

## Features

- ✅ **No Microsoft Word required** - Generates PDFs directly
- ✅ **Self-contained executable** - No .NET or other runtime needed on target machine
- ✅ **Cross-platform development** - Built with .NET 9 and Avalonia UI
- ✅ **Live progress tracking** - Shows current student and remaining count
- ✅ **Simple GUI** - Easy file selection and one-click generation

## How It Works

1. **Excel File** - Contains student data (names, grades, etc.)
2. **Mapping File** - JSON file that maps Excel columns to report fields
3. **Class Name** - Used for output folder naming
4. **Output** - PDF report cards generated in `{ClassName} report_cards/` folder

## Usage

### GUI Mode
1. Run `SchoolReportGenerator.exe`
2. Browse and select your Excel file
3. Browse and select your Word template (used for reference only)
4. Browse and select your mapping JSON file
5. Enter the class name
6. Click "Generate Report Cards"
7. Watch the progress bar as reports are generated

### Command Line Mode
```bash
SchoolReportGenerator.exe <excel_path> <template_path> <mapping_path> <class_name>
```

Example:
```bash
SchoolReportGenerator.exe "./marks.xlsx" "./template.docx" "./mapping.json" "Class_5A"
```

## File Formats

### Excel File
- First row: Headers (ignored)
- Subsequent rows: Student data
- Columns referenced by letter (A, B, C, etc.)

### Mapping JSON
Maps field names to Excel column letters:
```json
{
    "name": "A",
    "class": "B",
    "result": "C",
    "marks": "D"
}
```

## Development

### Prerequisites
- .NET 9 SDK
- Visual Studio Code or Visual Studio

### Build from Source
```bash
cd SchoolReportGeneratorCSharp
dotnet restore
dotnet build
dotnet run
```

### Build Windows Executable
```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

Output: `bin/Release/net9.0/win-x64/publish/SchoolReportGenerator.exe`

## Project Structure

```
SchoolReportGenerator/
├── SchoolReportGeneratorCSharp/     # C# source code
│   ├── Program.cs                   # Entry point
│   ├── MainWindow.axaml             # UI layout
│   ├── MainWindow.axaml.cs          # UI logic
│   └── Services/
│       ├── DataProcessor.cs         # Excel reading
│       └── ReportCardGenerator.cs   # PDF generation
├── input_files/                     # Test files
├── mappings/                        # Mapping templates
├── dist/                            # Compiled executables
└── README.md
```

## Technology Stack

- **C# / .NET 9** - Core application
- **Avalonia UI** - Cross-platform GUI framework
- **ClosedXML** - Excel file reading
- **QuestPDF** - PDF generation
- **No external dependencies** - Runs standalone on Windows

## Legacy Python Version

The original Python version is preserved in `README_old.md`. It required Microsoft Word for PDF conversion and had DLL compatibility issues across different Windows machines. The C# version solves these issues with a fully self-contained approach.

## License

MIT License

