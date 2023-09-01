# Plain Text Document Metadata

This Python script is designed to extract metadata from PDF and Word documents in a specified folder and create an Excel spreadsheet with the extracted information. It also extracts software information from these documents and populates the spreadsheet with it.

## Prerequisites

Before running the script, make sure you have the following Python libraries installed:

- PyPDF2
- python-docx
- openpyxl
- Pillow (PIL)

You can install these libraries using pip:

```bash
pip install PyPDF2 python-docx openpyxl Pillow
```
## Usage
1. Clone this repository to your local machine.
2. Modify the doc_folder variable in the script to specify the path to the folder containing the documents you want to process.
3. Ensure that a metadata text file (metadata.txt) is present in the specified folder
4. This text file should contain metadata information for the documents in the folder.
5. Customize the organisation variable with the name of your organization.
6. Run the script using the following command:
```bash
python text_compiler.py
```
The script will generate an Excel file named text_metadata.xlsx in the same folder as the documents.
This file will contain the extracted metadata and software information.

### metadata template
The .txt file will have the following column headers:<br>

Title:<br>
Abstract: <br>
Type:<br>
First Name:<br>
Last Name:<br>
Organisation:<br>
Page Count:<br>
Date Published:<br>
Publisher:<br>
Place Published:<br>
Volume/Issue/Report Number:<br>
ISBN:<br>
DOI:<br>
URL:<br>
Language: English<br>
Software:<br>
Software Version:<br>


### Notes
Make sure to review and adjust the script according to your specific requirements and the format of your metadata text file.
This script may require additional configuration or adjustments based on the structure and content of your documents.

