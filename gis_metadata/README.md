# GIS Metadata

## Table of Content
- [Extractor script](#extractor-script)
- [Attribute Tables](#attribute-tables)
- [Sample Output](#sample-output)
- [Compiler script](#compiler-script)
- [Usage](#usage)

  
## [Extractor script](gis_extractor.py) 

This Python script is designed to extract metadata from shapefiles located in a specified directory. It generates both general metadata about the shapefiles and attribute tables for each shapefile's features.

## Usage

1. Install the required library:

   ```bash
   pip install geopandas
   
2. Clone or download the script to your local machine.

### Modify the script:

1. Set the shapefile_directory variable to the path where your shapefiles are located.
2. Customize the metadata information, such as software and version, to match your environment.
3. Run the script:
  ```bash
python gis_extractory.py
  ```
4. Output
  General Metadata: The script will create a gis_metadata.txt file containing general information about the shapefiles in the specified directory.

### Attribute Tables: 
For each shapefile found, the script will create a separate <shapefile_name>_attributes.txt file containing attribute information for each feature.

#### Sample Output
For each shapefile, the script will generate metadata like:

Name<br>
CRS (Coordinate Reference System)<br>
Geometry Type<br>
Number of Features<br>
Start Date (Creation Date)<br>
End Date (Last Modification Date)<br>
Software and Version<br>


# [Compiler script](gis_compiler.py)

This script is designed to generate metadata from a text file and organize it into an Excel spreadsheet for GIS data. It also extracts specific information from shapefiles and populates the spreadsheet with relevant data.

## Usage

## Dependencies
1. Install the required library:

   ```bash
   pip install geopandas
   pip install openpyxl

2. Clone or download the scripts to your local machine.

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


### Contributing
Contributions are welcome! If you have suggestions, bug reports, or improvements, please open an issue or submit a pull request.

### License
This project is licensed under the MIT License.

### Acknowledgments
This script was inspired by the need to extract metadata from shapefiles for geospatial analysis.
Special thanks to the geopandas library for simplifying geospatial data manipulation.
