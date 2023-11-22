<h1> Metadata Processing Scripts </h1>

<p style="font-size: 20px">If you need to create metadata for an archaeological project with more than <strong>1000 files</strong>, you are in the right place. This suite of scripts can help you save a few hours of work.</p>
<p style="font-size: 20px">This repository contains a collection of Python scripts for processing and generating metadata from various sources, including Excel spreadsheets, plain text documents, shapefiles, and a web scraping script to gather all the projects already on ADS Library.</p><br>

<details>

<summary>Table of Contents</summary>

- [Contributing](#contributing)
- [Raster](#raster)
- [Spreadsheet](#spreadsheet)
- [Plain Text Document](#plain-text-document)
- [GIS (.shp)](#gis-shp)
- [Projects Scraping](projects_scraping)
- [ADS - Data Type requirements](https://archaeologydataservice.ac.uk/help-guidance/instructions-for-depositors/files-and-metadata/)

</details>

### Contributing
Contributions are welcome! If you have suggestions, bug reports, or improvements, please open an issue or submit a pull request.


### [Raster](raster_metadata)

This script extract metadata from photos using the exif tags and from the drawings (usually permatraces in .png format) and compiles it into an Excel spreadsheet. It also applies styling and formatting to the spreadsheet following the ADS template. The text_compiler.py is designed to extract metadata from PDF and Word documents from a specified folder and create an Excel spreadsheet (ADS template) with the extracted information. It also extracts software information from these documents and populates the spreadsheet with it.

### [Spreadsheet](spreadsheet_metadata)

The spreadsheet_extractor script is designed to process Excel files using the openpyxl library. It extracts metadata about the workbook, counts the number of non-empty rows in each sheet, and also extracts header names from the first sheet.
The spreadsheet_compiler script is designed to generate an Excel spreadsheet containing metadata for a collection of documents. It reads metadata information from a text file and organizes it into columns in an Excel file.

### [Plain Text Document](plain_text_document_metadata)

The text_compiler.py is designed to extract metadata from PDF and Word documents from a specified folder and create an Excel spreadsheet (ADS template) with the extracted information. It also extracts software information from these documents and populates the spreadsheet with it. This Python script scrapes project name from a specific organization URLs, processes it, and exports it to an Excel file.

### [GIS (.shp)](gis_metadata)

This Python script is designed to extract metadata from shapefiles located in a specified directory. It generates both general metadata about the shapefiles and attribute tables for each shapefile's features, 1 attributes.txt file for each shapefiles and 1 gis_metadata.txt for all the shapefiles inside the directory.
```bash
<shapefile_name>_attributes.txt
gis_metadata.txt
```

### [Projects Scraping](projects_scraping)

This Python script scrapes project name from a specific organization URLs, processes it, and exports it to an Excel file.
