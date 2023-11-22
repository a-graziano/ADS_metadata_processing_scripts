# Raster Metadata
This script extracts metadata from photos and drawings and compiles it into an Excel spreadsheet. It also applies styling and formatting to the spreadsheet.

<details>

<summary>Table of Contents</summary>

- [Prerequisites](#prerequisites)
- [Usage](#usage)
- [Important Notes](#important-notes)

</details>


## Prerequisites

Before running the script, make sure you have the following Python libraries installed:

- DateTime
- python-docx
- openpyxl
- Pillow (PIL)

You can install these libraries using pip:

```bash
pip install DateTime python-docx openpyxl Pillow
```


### Usage

Set up the following variables in the script:

1. photo_folder: Path to the folder containing the photos.
2. metadata_file_path: Path to the photo metadata text file.
3. dwg_file_path: Path to the dwg metadata text file.
4. Ensure you have a captions spreadsheet named captions.xlsx with two columns: Filename and Caption(don't use any headers, just the filename and its caption)
5. Run the script:
```bash
python raster_compiler.py
```
6. Metadata is extracted from each photo and added to the Excel sheet.
7. The Excel file is saved in the photo folder as raster_metadata.xlsx.

## Important Notes
<ul>
  <li>Ensure the paths are correctly set for photo_folder, metadata_file_path, and dwg_file_path.</li>
  <li>Make sure you have a captions spreadsheet named captions.xlsx with columns Filename and Caption.</li>
<li>The script assumes that metadata in the text files is formatted as key-value pairs, e.g., Key: Value.</li>
<li>The script creates a new Excel workbook with the metadata.</li>
</ul>

