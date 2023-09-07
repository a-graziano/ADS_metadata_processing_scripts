# Raster Metadata Script
This script extracts metadata from photos and drawings and compiles it into an Excel spreadsheet. It also applies styling and formatting to the spreadsheet.

## Dependancies
```bash
os
datetime
openpyxl
PIL
```
## Usage
Set up the following variables in the script:

2. photo_folder: Path to the folder containing the photos.
3. metadata_file_path: Path to the photo metadata text file.
4. dwg_file_path: Path to the dwg metadata text file.
5. Ensure you have a captions spreadsheet named captions.xlsx with two columns: Filename and Caption.
6. Run the script.

## Script Overview
Loading Dependencies:

This script uses various libraries including os, datetime, openpyxl, and PIL for handling files, dates, Excel operations, and image processing respectively.<br>

#### Setting Paths:

Paths to the photo folder and metadata files are defined.

#### Loading Captions:

Captions are loaded from the captions.xlsx file.

#### Reading Metadata:

Metadata is extracted from the text files (metadata.txt and dwg_metadata.txt).
#### Creating Excel Workbook:

An Excel workbook is created.
#### Setting Headers and Styles:

Column headers are defined and styled.
#### Looping through Photos:

Metadata is extracted from each photo and added to the Excel sheet.
#### Looping through Drawings:

Metadata is extracted from each drawing and added to the Excel sheet, following the photos.
#### Saving the Excel File:

The Excel file is saved in the photo folder as raster_metadata.xlsx.

## Important Notes
Ensure the paths are correctly set for photo_folder, metadata_file_path, and dwg_file_path.
Make sure you have a captions spreadsheet named captions.xlsx with columns Filename and Caption.
The script assumes that metadata in the text files is formatted as key-value pairs, e.g., Key: Value.
The script creates a new Excel workbook with the metadata.
