#!/usr/bin/env python3
import os
import platform
import subprocess
import shutil
import sys
from pathlib import Path

print("=" * 50)
print("Building PO Formatter Portable Application")
print("=" * 50)

# Use Python and pip from virtual environment if running in one
in_venv = hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix)
print(f"Running in virtual environment: {in_venv}")

# No need to reinstall dependencies if we're already in a venv
if not in_venv:
    print("Please run this script from within a virtual environment")
    print("Create one with: python3 -m venv venv")
    print("Activate it with: source venv/bin/activate (Linux/Mac) or venv\\Scripts\\activate (Windows)")
    sys.exit(1)

# Determine OS-specific parameters
is_windows = platform.system() == "Windows"
is_mac = platform.system() == "Darwin"

print(f"Building for: {platform.system()}")

# Create runtime hook to ensure temp directory access
print("Creating runtime hook...")
with open("runtime_hook.py", "w") as f:
    f.write("""
import os
import sys
import tempfile

# Ensure temp directory is accessible
if hasattr(sys, '_MEIPASS'):
    os.environ['TMPDIR'] = os.path.join(sys._MEIPASS, 'temp')
    if not os.path.exists(os.environ['TMPDIR']):
        os.makedirs(os.environ['TMPDIR'])
""")

def main():
    print("Starting PO Formatter build process...")
    
    # Make sure PyInstaller is installed
    try:
        import PyInstaller
        print(f"Using PyInstaller version {PyInstaller.__version__}")
    except ImportError:
        print("PyInstaller not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # Check for the icon file
    icon_path = os.path.join("Images", "POChange.png")
    if not os.path.exists(icon_path):
        print(f"Warning: {icon_path} not found! Please place the icon file in the Images directory.")
        return
    
    # Create a directory for the build files if it doesn't exist
    build_dir = Path("build")
    dist_dir = Path("dist")
    
    if not build_dir.exists():
        build_dir.mkdir()
    
    # Clean previous build files if they exist
    if dist_dir.exists():
        print("Cleaning previous build...")
        shutil.rmtree(dist_dir)
    
    # Run PyInstaller
    print("Building executable with PyInstaller...")
    subprocess.check_call([
        "pyinstaller",
        "--clean",
        "po_formatter.spec"
    ])
    
    # Create images directory in dist if it doesn't exist
    dist_images_dir = os.path.join("dist", "Images")
    if not os.path.exists(dist_images_dir):
        os.makedirs(dist_images_dir)
    
    # Copy the image file to the dist/images directory
    print("Copying resources...")
    if os.path.exists(os.path.join(dist_images_dir, "POChange.png")):
        print("Image file already included in the build.")
    else:
        shutil.copy2(icon_path, dist_images_dir)
        print("Image file copied to the dist directory.")
    
    print("\nBuild completed successfully!")
    print("Executable location: dist/PO_Formatter.exe")
    
    # Copy the executable to the root directory for easy access
    shutil.copy2(os.path.join("dist", "PO_Formatter.exe"), "PO_Formatter.exe")
    print("Executable copied to the root directory as PO_Formatter.exe")

if __name__ == "__main__":
    main()

# Clean up temporary files
print("Cleaning up temporary files...")
temp_files = ["runtime_hook.py"]
for file in temp_files:
    if os.path.exists(file):
        try:
            os.remove(file)
        except:
            pass

print("\nTo use the portable application:")
print("1. Copy PO_Formatter.exe to your USB drive")
print("2. Run directly from the USB drive - no installation needed")
print("=" * 50) 