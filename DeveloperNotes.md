# PO File Formatter - Developer Notes

## Technical Implementation

### Architecture Decisions

1. **PySide6 vs. PyQt5**: We initially planned to use PyQt5 but switched to PySide6 due to compatibility issues with macOS. PySide6 provides the same Qt functionality with better cross-platform support.

2. **Pandas for Data Processing**: We chose pandas for Excel processing due to its robust handling of tabular data and excellent support for various output formats including CSV, TSV, and Excel.

3. **PyInstaller for Packaging**: We selected PyInstaller over alternatives like cx_Freeze because it:
   - Creates a single executable file
   - Works well on both Windows and macOS
   - Allows fine-tuning through spec files
   - Handles runtime dependencies efficiently

4. **Spec File Optimization**: We created a custom spec file to:
   - Exclude unnecessary libraries (matplotlib, scipy)
   - Include all required dependencies
   - Configure proper runtime behavior
   - Fix macOS path handling

### Code Structure

- **Main Window Class**: The `POFormatter` class handles the UI and orchestrates the workflow
- **Format Methods**: Each vendor format has its own method that handles formatting logic
- **Validation**: Input validation happens at multiple levels:
  - File selection validation
  - Required fields checking
  - PO number validation
  - Vendor-specific data validation

### Platform-Specific Considerations

#### macOS

- Required runtime hooks for proper temporary file handling
- Built using a virtual environment to avoid system Python issues
- Used macOS-specific build flags for compatibility

#### Windows

- Used Windows-native path separators in the build script
- Added Windows-specific version info
- Configured runtime tmp directory for USB operation

## Vendor Formatting Implementation

The vendor-specific formatting follows a pattern:

1. Copy the DataFrame to avoid modifying the original
2. Apply vendor-specific transformations:
   - Select/rename columns
   - Format numeric values
   - Add/modify required fields
   - Convert to vendor-specific format
3. Save to user-selected location in correct format

## Future Enhancements

1. **Custom Vendor Configuration**: Add ability to configure vendor formatting rules without code changes

2. **Template System**: Implement a template system to allow users to create and save custom formatting profiles

3. **Batch Processing**: Add support for processing multiple PO files in a single operation

4. **Preview Mode**: Add a preview window to show how the formatted output will look before saving

5. **Logging**: Implement better logging for troubleshooting issues

6. **Auto-Updates**: Add a mechanism to check for and apply updates to the application

7. **Advanced Error Recovery**: Implement more sophisticated error handling and recovery

## Deployment Considerations

### USB Drive Requirements

- The application runs from any USB drive with no special requirements
- Executable size is approximately 54MB (macOS) and 40-60MB (Windows)
- No additional files or folders needed

### Security

- The application does not require administrative privileges
- It does not modify system files or registry
- All file operations happen within user-selected locations
- No network access required for core functionality

### Performance

- Loading large Excel files may require more memory
- Consider the following optimizations for large files:
  - Chunk processing for very large datasets
  - Progress indicators for long-running operations
  - Memory usage optimizations

## Maintenance Notes

- The PyInstaller spec file should be updated if dependencies change
- When upgrading Python or dependencies, rebuild the executable and test thoroughly
- Vendor format changes will require code updates in the corresponding format method 