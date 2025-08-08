# 📊 Simple Excel Diff Tool v.0.1.0-beta

A streamlined web application for identifying differences between Excel files quickly and efficiently. Simple-excel-dif helps you compare two Excel spreadsheets and visualize the changes with an intuitive, user-friendly interface.

## 🌟 Features

- **Easy-to-Use Interface**: Simple drag-and-drop or file selection for two Excel files
- **Smart Comparison and Visual Highlighting**: Automatically  highlights changes between files for easy identification
- **Flexible Filtering**: Filter results by any column to focus on specific data
- **Export Capabilities**: Download comparison results as an Excel file
- **Large File Support**: Handles files up to 100MB 
- **Preview Mode**: View contents of both files before comparison
 


## 🥺 Limitations

- Only supports .xlsx format (.xls not supported)
- Excel internal data filters must be disabled.
- It uses columns from the old file to compare. 
- Newly added columns in new file will be ignored.
- First row considered as header.
- Huge files may require increased memory allocation