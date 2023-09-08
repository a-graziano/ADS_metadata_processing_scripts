# ADS Metadata Processing Scripts

This repository contains a collection of Python scripts for processing metadata from various sources, including Excel spreadsheets, plain text documents, shapefiles, and web scraping to gather all the projects already on ADS Library.

## Table of Content
- [Spreadsheet Metadata](spreadsheet_metadata)
- [Plain Text Document Metadata](plain_text_document_metadata)
- [GIS Shapefile Metadata](gis_metadata)
- [Raster Metadata](raster_metadata)
- [Projects Scraping](projects_scraping)

## Raster Metadata

### First Step

- **File**: [raster_metatadata.txt](raster_metadata)
- **Description**: This text file contains all the general information relating to the raster files(.jpg, .png, .tiff) which will then be extracted and inserted into a spreadsheet after performing the next step.

### Second Step

- **File**: [raster_compiler.py](raster_metadata)
- **Description**: This script processes metadata from photos and DWG drawings, extracts information such as filenames, captions, keywords, periods, creators, and more. It creates an Excel workbook with organised metadata and applies styling to the column headers. To create the metadata file, it also uses the information relating to each individual file which is contained in another spreadsheet. Usually this is the File Name and Caption.


## Projects Scraping (Additional Script)

- **File**: [projects_scraping.py](projects_scraping)
- **Description**: This script demonstrates how to scrape project information from web pages using Python. It specifically targets the project information of a specific organisation from the Archaeology Data Service website. It uses the `requests` library to send HTTP GET requests to URLs of various organisations on the Archaeology Data Service website, extracts organisation names and project titles using `BeautifulSoup`, stores the data in a pandas DataFrame, and exports it to an Excel file.

### Usage

1. Ensure you have installed the required Python packages using `pip`:

   ```bash
   pip install requests beautifulsoup4 pandas
   pip install openpyxl pillow
