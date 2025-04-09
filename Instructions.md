# PO File Formatter - Developer Instructions

## Project Overview

This is a portable standalone application that formats Purchase Order Excel files for various vendor systems. The application allows users to:

1. Select an Excel file containing PO data
2. Choose the target vendor format
3. Automatically extract the PO number from the filename
4. Format and save the output file according to vendor specifications

The application runs directly from a USB drive with no installation required on either Windows or macOS systems.

## Technical Implementation

### Architecture

- **Frontend**: PySide6 (Qt) for cross-platform GUI
- **Data Processing**: Pandas and openpyxl for Excel file handling
- **Packaging**: PyInstaller for creating standalone executables

### Key Files

- `po_formatter.py` - Main application with GUI and processing logic
- `po_formatter.spec` - PyInstaller spec file for optimized builds
- `build.py` - Build script to create standalone executables
- `requirements.txt` - Python dependencies

### Vendor Formats

The application supports multiple vendor formats, each with its own processing function:

1. **Vendor 1** - Formats to CSV with specific columns and PO number
2. **Vendor 2** - Creates tab-delimited text files
3. **Vendor 3** - Formats to CSV with custom headers
4. **Vendor 4** - Creates CSV with specific column ordering
5. **Vendor 5** - Outputs as Excel file with specific sheet name

## Development Setup

### macOS

```bash
# Create virtual environment
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Build the application
python build.py
```

### Windows

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Build the application
python build.py
```

## Building the Portable Executable

The build process:

1. Creates a runtime hook to handle temporary files
2. Uses a PyInstaller spec file with optimized settings
3. Strips unnecessary libraries to reduce file size
4. Copies the executable to the root directory for easy access

The output is a single executable file that can be copied to a USB drive and run on any compatible system without installation.

## Testing

For testing, use the included `create_sample_data.py` script to generate sample PO Excel files:

```bash
python create_sample_data.py
```

This creates a sample file in the `sample` directory named `PO12345.xlsx`.

## Design Decisions

1. **Standalone Executable** - No dependencies or installation required
2. **Cross-platform** - Works on both Windows and macOS
3. **Self-contained UI** - Intuitive interface for non-technical users
4. **File Dialog Integration** - Native file selection and saving
5. **Error Handling** - Robust validation and user-friendly error messages

## Original Requirements

This application replaces an Excel macro that formatted PO data for vendor uploads. The portable executable approach allows users to run the application from a USB drive without installing Python or any dependencies on the target computer.

## Deployment

1. Build the executable on your development machine
2. Copy the single executable file to a USB drive
3. Distribute the USB drive to users
4. Users can run the application directly from the USB with no installation 