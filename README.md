# Excel_To_Multiple_Word_Files

# Mail Merge Application

This Python application automates the process of performing a mail merge using a Word template and data from an Excel file. It processes each record in the Excel sheet, replaces placeholders in the Word template with corresponding data, and generates individual Word documents. The output documents are saved in a folder structure based on the data from the Excel sheet.

## Features
- Select a Word template with placeholders to be replaced with Excel data.
- Process Excel data and replace placeholders in the Word template.
- Save the output documents in organized folders.
- Track the progress with a progress bar and log messages.
- Customize the file/folder names by stripping invalid characters.

## Requirements

To run this program, ensure you have Python installed along with the following required libraries:

### Required Libraries

1. **`pandas`**:
   - Description: Data manipulation and analysis library.
   - Installation Command:
     ```bash
     pip install pandas
     ```

2. **`python-docx`**:
   - Description: Library for reading and writing Word documents.
   - Installation Command:
     ```bash
     pip install python-docx
     ```

3. **`ttkbootstrap`**:
   - Description: A set of modern, enhanced themes for Tkinter.
   - Installation Command:
     ```bash
     pip install ttkbootstrap
     ```

4. **`openpyxl`**:
   - Description: A library to read/write Excel 2010 xlsx/xlsm files.
   - Installation Command:
     ```bash
     pip install openpyxl
     ```

5. **`tkinter`**:
   - Description: A standard Python library for creating GUI applications.
   - **Note**: `tkinter` is bundled with Python, so it doesn't need to be installed separately.

## How to Use

1. **Select Template File**: Click "Browse" to select the Word template that contains placeholders (e.g., `[Name]`, `[Date]`).
2. **Select Excel File**: Click "Browse" to select the Excel file containing data that will replace the placeholders in the template.
3. **Select Output Folder**: Click "Browse" to select the folder where the generated Word documents will be saved.
4. **Start Mail Merge**: Click "Start Mail Merge" to begin processing. The program will replace placeholders with the data from the Excel file and save the resulting documents.

## Example Output

The output files are saved in folders named according to the `PlaceholderName` and folder identifier (e.g., Placeholder1, Placeholder2) from the Excel data. Each generated file will be named `Letter.docx`.

## Troubleshooting

If you encounter issues, ensure all required libraries are installed correctly and that the input files are not corrupted. Check the log for detailed information on any errors during the process.

