# PO File Formatter - User Guide

## Quick Start Guide

This standalone application helps you format Purchase Order (PO) Excel files for different vendor ordering systems. The application runs directly from your USB drive - no installation needed.

### How to Use

1. **Start the Application**
   - On Windows: Double-click the `PO_Formatter.exe` file on your USB drive
   - On Mac: Double-click the `PO_Formatter` file on your USB drive

2. **Select Your PO File**
   - Click the "Browse..." button
   - Navigate to and select your Excel file containing PO data
   - The application will automatically extract the PO number from the filename

3. **Confirm PO Number**
   - The PO number field will automatically populate based on the filename
   - You can edit this if needed

4. **Select Vendor Format**
   - Choose the appropriate vendor from the dropdown menu

5. **Process Your File**
   - Click the "Process" button
   - Choose where to save the formatted output file
   - Click "Save"

6. **Upload to Vendor**
   - Use the formatted file to upload to your vendor's ordering system

## Supported Vendor Formats

The application currently supports five vendor formats:

1. **Vendor 1** - CSV format with PO number, item number, description, quantity, unit price, and total
2. **Vendor 2** - Tab-delimited text file
3. **Vendor 3** - CSV with custom headers
4. **Vendor 4** - CSV with specific column ordering
5. **Vendor 5** - Excel file with specific sheet name

## Tips and Troubleshooting

- **Filename as PO Number**: For best results, name your Excel file with the PO number (e.g., "PO12345.xlsx")
- **Missing Columns**: The application will notify you if your Excel file is missing required columns
- **File Format Issues**: Make sure your Excel file follows your company's standard PO format
- **Save Location**: You can save the output file anywhere - on your USB drive, desktop, or network location

## Need Help?

If you encounter any issues or have questions about using the PO File Formatter, please contact your IT department or system administrator.

---

*PO File Formatter v1.0 - Runs directly from USB with no installation required* 