# School Report Card Generator

Generate PDF report cards from Excel data and Word templates. **No Microsoft Word required!**

---

## 🚀 Quick Start (Recommended)

### Option 1: Docker (Works on any Windows PC)

#### One-Time Setup (5 minutes)

| Step | Action |
|------|--------|
| 1 | Go to [docker.com/products/docker-desktop](https://www.docker.com/products/docker-desktop/) |
| 2 | Click **"Download for Windows"** |
| 3 | Run installer → Next → Next → Finish |
| 4 | **Restart computer** |
| 5 | Open **Docker Desktop** from Start Menu |
| 6 | Wait until you see ✅ "Running" (bottom-left corner) |

#### Run the App

| Step | Action |
|------|--------|
| 1 | Make sure Docker Desktop is running (green whale 🐳 in taskbar) |
| 2 | Download [`Run-ReportGenerator.bat`](https://github.com/utkarsh-koppikar/SchoolReportGenerator/raw/main/Run-ReportGenerator.bat) |
| 3 | **Double-click** the `.bat` file |
| 4 | Browser opens automatically at `http://localhost:8080` |
| 5 | Upload your files and generate reports! |
| 6 | Press any key in the black window to stop |

**That's it! No command line needed.** 🎉

---

### Option 2: Online (No Installation)

Use the hosted version (may be slow on free tier):

🌐 **[reportgen-9eyd.onrender.com](https://reportgen-9eyd.onrender.com)**

---

### Option 3: Standalone EXE (Requires Microsoft Word)

**[Download SchoolReportGenerator.exe](https://github.com/utkarsh-koppikar/SchoolReportGenerator/releases/tag/latest)**

> ⚠️ This option requires Microsoft Word installed for PDF conversion.

---

## 📋 How It Works

1. **Excel File** (.xlsx) - Contains student data (names, grades, etc.)
2. **Word Template** (.docx) - Your report card design with `{{placeholders}}`
3. **Mapping File** (.json) - Maps Excel columns to template placeholders
4. **Output** - Individual PDF report cards for each student

### Example Mapping File

```json
{
    "name": "A",
    "class": "B",
    "english": "C",
    "math": "D",
    "science": "E",
    "result": "F"
}
```

This maps:
- Column A → `{{name}}`
- Column B → `{{class}}`
- Column C → `{{english}}`
- etc.

---

## ✨ Features

| Feature | Docker Version | EXE Version |
|---------|---------------|-------------|
| PDF Generation | ✅ LibreOffice (bundled) | ✅ Requires Word |
| No Installation | ✅ Just Docker | ✅ Self-contained |
| Web UI | ✅ Browser-based | ❌ Desktop app |
| Works Offline | ✅ Yes | ✅ Yes |
| Progress Tracking | ✅ Yes | ✅ Yes |

---

## 🛠️ For Developers

### Run with Docker

```bash
docker run -p 8080:8080 ghcr.io/utkarsh-koppikar/reportgen:latest
```

### Build from Source

```bash
# Clone
git clone https://github.com/utkarsh-koppikar/SchoolReportGenerator.git
cd SchoolReportGenerator

# Docker
cd portable-app
docker build -t reportgen .
docker run -p 8080:8080 reportgen

# Or C# version
cd SchoolReportGeneratorCSharp
dotnet restore
dotnet run
```

### Project Structure

```
SchoolReportGenerator/
├── portable-app/                    # Python/Flask web app
│   ├── Dockerfile                   # Docker config
│   ├── app/                         # Flask application
│   │   ├── main.py                  # Web server
│   │   ├── services/                # Core logic
│   │   └── templates/               # HTML UI
│   └── docker-compose.yml
├── SchoolReportGeneratorCSharp/     # C# desktop app
├── Run-ReportGenerator.bat          # Windows launcher
├── render.yaml                      # Cloud deployment
└── README.md
```

---

## 📦 Releases

| Release | Description | Download |
|---------|-------------|----------|
| Docker Image | Web app with LibreOffice | `ghcr.io/utkarsh-koppikar/reportgen:latest` |
| Portable App | ZIP with bundled dependencies | [portable-latest](https://github.com/utkarsh-koppikar/SchoolReportGenerator/releases/tag/portable-latest) |
| Windows EXE | Desktop app (needs Word) | [latest](https://github.com/utkarsh-koppikar/SchoolReportGenerator/releases/tag/latest) |

---

## 📄 License

MIT License
