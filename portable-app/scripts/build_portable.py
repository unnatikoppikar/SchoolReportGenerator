"""
Build Script for Portable Report Card Generator

This script:
1. Downloads LibreOffice Portable (if not present)
2. Downloads Python Embedded (if not present)
3. Packages everything into a self-contained ZIP

Final package requires NO installation - just extract and run.
"""

import os
import sys
import shutil
import zipfile
import urllib.request
import tempfile
from pathlib import Path

# Configuration
LIBREOFFICE_PORTABLE_URL = "https://download.documentfoundation.org/libreoffice/portable/7.6.4/LibreOfficePortable_7.6.4_MultilingualStandard.paf.exe"
PYTHON_EMBEDDED_URL = "https://www.python.org/ftp/python/3.11.7/python-3.11.7-embed-amd64.zip"

# Paths
ROOT_DIR = Path(__file__).parent.parent
BUILD_DIR = ROOT_DIR / "build"
DIST_DIR = ROOT_DIR / "dist"


def download_file(url: str, dest: Path, desc: str = "Downloading"):
    """Download a file with progress."""
    print(f"{desc}: {url}")
    
    def progress_hook(count, block_size, total_size):
        percent = int(count * block_size * 100 / total_size) if total_size > 0 else 0
        print(f"\r  Progress: {percent}%", end="", flush=True)
    
    urllib.request.urlretrieve(url, str(dest), progress_hook)
    print()  # New line after progress


def setup_libreoffice_portable():
    """
    Download and extract LibreOffice Portable.
    
    For Windows: Downloads the portable installer
    Note: The .paf.exe is a self-extracting installer, 
    we'll need to extract it or instruct user to run it.
    """
    lo_dir = BUILD_DIR / "libreoffice"
    
    if lo_dir.exists() and (lo_dir / "App" / "libreoffice" / "program" / "soffice.exe").exists():
        print("✓ LibreOffice Portable already present")
        return lo_dir
    
    print("\n" + "="*50)
    print("Setting up LibreOffice Portable")
    print("="*50)
    
    # For automated builds, we use a pre-extracted version
    # The LibreOffice Portable installer (.paf.exe) needs manual extraction
    
    print("""
NOTE: LibreOffice Portable needs to be set up manually:

1. Download LibreOffice Portable from:
   https://www.libreoffice.org/download/portable-versions/

2. Run the installer and extract to:
   {build_dir}/libreoffice/

3. After extraction, verify this path exists:
   {build_dir}/libreoffice/App/libreoffice/program/soffice.exe

4. Re-run this build script.
""".format(build_dir=BUILD_DIR))
    
    # Create placeholder directory
    lo_dir.mkdir(parents=True, exist_ok=True)
    
    return None


def setup_python_embedded():
    """Download and set up Python Embedded for Windows."""
    python_dir = BUILD_DIR / "python"
    
    if python_dir.exists() and (python_dir / "python.exe").exists():
        print("✓ Python Embedded already present")
        return python_dir
    
    print("\n" + "="*50)
    print("Setting up Python Embedded")
    print("="*50)
    
    python_dir.mkdir(parents=True, exist_ok=True)
    
    # Download Python Embedded
    zip_path = BUILD_DIR / "python-embedded.zip"
    download_file(PYTHON_EMBEDDED_URL, zip_path, "Downloading Python Embedded")
    
    # Extract
    print("Extracting Python...")
    with zipfile.ZipFile(zip_path, 'r') as zf:
        zf.extractall(python_dir)
    
    # Enable pip by modifying python311._pth
    pth_file = python_dir / "python311._pth"
    if pth_file.exists():
        content = pth_file.read_text()
        content = content.replace("#import site", "import site")
        pth_file.write_text(content)
    
    # Download get-pip.py and install pip
    getpip_path = python_dir / "get-pip.py"
    download_file("https://bootstrap.pypa.io/get-pip.py", getpip_path, "Downloading pip")
    
    os.remove(zip_path)
    print("✓ Python Embedded set up")
    
    return python_dir


def copy_app_files():
    """Copy application files to build directory."""
    print("\n" + "="*50)
    print("Copying Application Files")
    print("="*50)
    
    app_dest = BUILD_DIR / "app"
    
    # Remove old app directory
    if app_dest.exists():
        shutil.rmtree(app_dest)
    
    # Copy app
    shutil.copytree(ROOT_DIR / "app", app_dest)
    print("✓ Copied app/")
    
    # Copy config
    config_dest = BUILD_DIR / "config"
    if config_dest.exists():
        shutil.rmtree(config_dest)
    shutil.copytree(ROOT_DIR / "config", config_dest)
    print("✓ Copied config/")
    
    # Copy sample files
    input_dest = BUILD_DIR / "input_files"
    if input_dest.exists():
        shutil.rmtree(input_dest)
    shutil.copytree(ROOT_DIR / "input_files", input_dest)
    print("✓ Copied input_files/")
    
    # Copy launcher scripts
    shutil.copy(ROOT_DIR / "start.bat", BUILD_DIR / "start.bat")
    shutil.copy(ROOT_DIR / "start.sh", BUILD_DIR / "start.sh")
    print("✓ Copied launcher scripts")
    
    # Copy README
    shutil.copy(ROOT_DIR / "README.md", BUILD_DIR / "README.md")
    print("✓ Copied README.md")
    
    # Copy requirements
    shutil.copy(ROOT_DIR / "requirements.txt", BUILD_DIR / "requirements.txt")
    print("✓ Copied requirements.txt")


def create_distribution():
    """Create final distribution ZIP."""
    print("\n" + "="*50)
    print("Creating Distribution")
    print("="*50)
    
    DIST_DIR.mkdir(parents=True, exist_ok=True)
    
    zip_name = "ReportCardGenerator-Portable.zip"
    zip_path = DIST_DIR / zip_name
    
    print(f"Creating: {zip_path}")
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(BUILD_DIR):
            # Skip __pycache__ directories
            dirs[:] = [d for d in dirs if d != "__pycache__"]
            
            for file in files:
                file_path = Path(root) / file
                arc_name = file_path.relative_to(BUILD_DIR)
                zf.write(file_path, f"ReportCardGenerator/{arc_name}")
    
    size_mb = zip_path.stat().st_size / (1024 * 1024)
    print(f"✓ Created: {zip_path} ({size_mb:.1f} MB)")
    
    return zip_path


def main():
    print("\n" + "="*60)
    print("  Report Card Generator - Build Script")
    print("="*60)
    
    # Create build directory
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    
    # Setup components
    lo_path = setup_libreoffice_portable()
    if lo_path is None:
        print("\n⚠ LibreOffice Portable not set up. Please follow the instructions above.")
        print("Build will continue, but PDF generation won't work without LibreOffice.\n")
    
    # For Windows distribution, we'd set up Python Embedded
    # setup_python_embedded()  # Uncomment for full Windows build
    
    # Copy app files
    copy_app_files()
    
    # Create distribution
    zip_path = create_distribution()
    
    print("\n" + "="*60)
    print("  Build Complete!")
    print("="*60)
    print(f"\nDistribution: {zip_path}")
    print("\nTo create a complete portable package:")
    print("1. Extract LibreOffice Portable to: build/libreoffice/")
    print("2. Run this script again")
    print("3. The ZIP will include everything needed")


if __name__ == "__main__":
    main()

