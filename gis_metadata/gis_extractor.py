import geopandas as gpd
import os
import datetime

# Directory where the shapefiles are located
shapefile_directory = ""

# Get a list of all .shp files in the directory
shapefile_paths = [os.path.join(shapefile_directory, filename) for filename in os.listdir(shapefile_directory) if filename.endswith(".shp")]

# Open the general output file
general_output_file = "gis_metadata.txt"
with open(general_output_file, 'w') as general_f:
    general_f.write("")

    for shapefile_path in shapefile_paths:
        # Read shapefile
        gdf = gpd.read_file(shapefile_path)

        # Extract shapefile name without extension
        shapefile_name = os.path.splitext(os.path.basename(shapefile_path))[0]

        # Output file path for attributes
        attribute_output_file = "{}_attributes.txt".format(shapefile_name)

        # Open the attribute output file
        with open(attribute_output_file, 'w') as attribute_f:
            attribute_f.write("Attribute Table for {}\n\n".format(shapefile_name))

            # Write attribute names and values
            for index, row in gdf.iterrows():
                attribute_f.write("----- Feature {} -----\n".format(index))
                for column, value in row.iteritems():
                    attribute_f.write("{}: {}\n".format(column, value))
                attribute_f.write("\n")

        # Write general metadata to the general output file
        general_f.write("File name(s): {}\n".format(shapefile_name))
        general_f.write("")

        # Extract metadata for the general output
        general_f.write("Name: {}\n".format(shapefile_name))
        general_f.write("CRS: {}\n".format(gdf.crs))
        general_f.write("Geometry Type: {}\n".format(gdf.geom_type.unique()))
        general_f.write("Number of Features: {}\n".format(len(gdf)))

        # Add Start Date (creation date) metadata
        creation_time = datetime.datetime.fromtimestamp(os.path.getctime(shapefile_path))
        general_f.write("Start Date: {}\n".format(creation_time.strftime('%Y-%m-%d %H:%M:%S')))

        # Add End Date (last modification date) metadata
        modification_time = datetime.datetime.fromtimestamp(os.path.getmtime(shapefile_path))
        general_f.write("End Date: {}\n".format(modification_time.strftime('%Y-%m-%d %H:%M:%S')))

        # Add Software and Version metadata (customize with actual information)
        general_f.write("Language: English\n")
        general_f.write("Software: {}\n".format("QGIS"))
        general_f.write("Software Version: {}\n".format("3.28"))

        if gdf.crs == {'init': 'epsg:27700'}:
            general_f.write("Coordinate Grid System: OSGB-36\n")

        general_f.write("\n")

print("General metadata has been written to {}".format(general_output_file))
print("Attribute table information files have been generated.")
