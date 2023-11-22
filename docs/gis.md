# GIS Metadata

<details>

<summary>Table of Contents</summary>

- [Prerequisites](#prerequisites)
- [Extractor script](#extractor-script)
- [Sample Output](#sample-output)
- [Compiler script](#compiler-script)
- [Usage](#usage)

</details>

### Prerequisites
1. Install the required library:

   ```bash
   pip install geopandas openpyxl

2. Clone or download the scripts to your local machine.

### [Extractor script](https://github.com/a-graziano/ADS_metadata_processing_scripts/blob/main/gis_metadata/gis_extractor.py) 

This Python script is designed to extract metadata from shapefiles located in a specified directory. It generates both general metadata about the shapefiles and attribute tables for each shapefile's features, 1 attributes.txt file for each shapefiles and 1 gis_metadata.txt for all the shapefiles inside the directory.
```bash
<shapefile_name>_attributes.txt
gis_metadata.txt
```
#### Sample Output
For each shapefile, the script will generate inside the gis_metadata.txt the information showed below:

Name<br>
CRS (Coordinate Reference System)<br> the script automatically assuming the value OSGB-36<br>
Geometry Type<br>
Number of Features<br>
Start Date (Creation Date)<br>
End Date (Last Modification Date)<br>
Software and Version<br>


### [Compiler script](https://github.com/a-graziano/ADS_metadata_processing_scripts/blob/main/gis_metadata/gis_compiler.py)

This script is designed to generate metadata from a text file and organize it into an Excel spreadsheet for GIS data. It also extracts specific information from shapefiles and populates the spreadsheet with relevant data.


## Usage

### Modify the script:

1. Place your shapefiles (with extensions .shp, .dbf, .shx) in your directory
2. Set the shapefile_directory variable to the path where your shapefiles are located.
3. Modify the `txt_file_path` variable in the script to point to your metadata text file.
4. Customize the metadata information, such as software and version, to match your environment.
5. Run the script:
```bash
python gis_extractory.py
```
4. Output
  General Metadata: The script will create a gis_metadata.txt file containing general information about the shapefiles in the specified directory.
7. Set the name of the organization using the `organisation` variable in the compiler script. For example: organisation = "Your organisation name"
8. Execute the script:
```bash
python gis_compiler.py
```
9. The script will generate an Excel file named gis_metadata.xlsx containing the organized metadata.
11. The Excel file is saved in the same directory as the shapefiles.