# Spreadsheet Metadata

<details>

<summary>Table of Contents</summary>

- [Prerequisites](#prerequisites)
- [Usage](#usage)
- [Script Explanation](#script-explanation)
- [Example Output](#example-output)
- [Compiler script](#compiler-script)

</details>

The [spreadsheet_extractor script](https://github.com/a-graziano/ADS_metadata_processing_scripts/blob/main/spreadsheet_metadata/spreadsheet_extractor.py) is designed to process Excel files using the `openpyxl` library. It extracts metadata about the workbook, counts the number of non-empty rows in each sheet, and also extracts header names from the first sheet.<br>
The [spreadsheet_compiler script](https://github.com/a-graziano/ADS_metadata_processing_scripts/blob/main/spreadsheet_metadata/spreadsheet_compiler.py) is designed to generate an Excel spreadsheet containing metadata for a collection of documents. It reads metadata information from a text file and organizes it into columns in an Excel file.

### Prerequisites

Before running the script, make sure you have Python installed on your system. You also need to install the `openpyxl` library if it's not already installed. You can install it using pip:

```bash
pip install openpyxl
```

## Usage
1. Clone the repository or download the script files to your local machine.
2. Update the excel_files list in the script with the paths of the Excel files you want to process.
3. Run the script:
```bash
python spreadsheet_extractor.py
```
5. The script will generate a spreadsheet_metadata.txt file containing metadata and row count information for each Excel file. Additionally, it will create _header.txt files for each Excel file, containing header names from the first sheet.

### Script Explanation
<ul>
<li>The script iterates through each Excel file specified in the excel_files list.</li>
<li>It extracts metadata information such as title, description, creator, and timestamps from the workbook's properties.</li>
<li>It counts the number of non-empty rows in each sheet of the workbook.</li>
<li>It creates a separate text file named after each Excel file with _header.txt appended and writes the header names from the first sheet to this file.</li>
</ul>

### Example output
Excel File: file path</br>
Title: Excel Title</br>
Description: Excel Description</br>
Creator: Excel Creator</br>
Copyright Holder: Excel Creator</br>
Start Date: 2023-01-01</br>
End Date: 2023-09-01</br>
Language: English</br>
Sheet Name: Sheet1</br>
Number of rows: 100</br>

...

Excel File: file path</br>
Title: Feature List Title</br>
Description: Feature List Description</br>
Creator: Feature List Creator</br>
Copyright Holder: Feature List Creator</br>
Start Date: 2023-01-01</br>
End Date: 2023-09-01</br>
Language: English</br>
Sheet Name: Data</br>
Number of rows: 50</br>
</br>

### Compiler script
#### Modify the script's configuration to suit your needs:

1. Set the doc_folder variable to the path of the folder containing the spreadsheets you want to generate metadata for.
2. Set the metadata_file_path variable to the path of the metadata text file.
3. Adjust the organization name by changing the organisation variable.
4. Run the script:

```bash
python spreadsheet_compiler.py
```
The script will generate an Excel file named spreadsheet_metadata.xlsx in the same folder as the script, containing the metadata for the documents.
