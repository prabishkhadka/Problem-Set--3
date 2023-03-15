@@ -0,0 +1,137 @@
#----------------------   Problem Set 3 -------
#Can change the country name and Player position as desired

country = "United States"
positionList = ["C","LW","RW"]
# Print the Selected information
print(f" The country selected is : {country}")
print(f" The postions of the player selected is {positionList[0]} , {positionList[1]} , {positionList[2]}")
# Import the necessary modules
import arcpy
import pandas as pd
import os
# Set the workspace to the directory containing the data
arcpy.env.workspace = r"C:\Geospatial Programming\Problem Set 3 Data"

# Set the overwrite output environment to True
arcpy.env.overwriteOutput = True

# Get a list of all the feature classes in the workspace
fcList = arcpy.ListFeatureClasses()

# Loop through each feature class and print its index and name
for index, fc in enumerate(fcList):
    print(f"Index: {index}, Feature Class: {fc}")

# Define the roster and world variables
roster = "nhlrosters.shp"
world = "Countries_WGS84.shp"

# Create feature layers for the roster and world shapefiles
arcpy.MakeFeatureLayer_management(roster, "rosterTemp")
arcpy.MakeFeatureLayer_management(world, "worldTemp")

# Define the output feature class name
rosterJoin = "rosterJoinCountries"

# Perform a spatial join operation between the rosterTemp and worldTemp feature layers
arcpy.SpatialJoin_analysis("rosterTemp", "worldTemp", rosterJoin, "JOIN_ONE_TO_ONE", "KEEP_ALL", "", "WITHIN")

# Create a feature layer for the joined data
arcpy.MakeFeatureLayer_management(rosterJoin, "rosterJoinTemp")

# Create an empty list to store the output shapefiles
outputList = []

# Use a try-except block to catch any execution errors
try:
    # Create a search cursor to iterate over the joined data
    with arcpy.da.SearchCursor("rosterJoinTemp", "CNTRY_NAME") as cursor:
        # Loop through each row in the cursor
        for row in cursor:
            # Check if the country name matches the desired country
            if row[0] == country:
                # Create a feature class for the desired country
                featureCountry = country
                query = """CNTRY_NAME = '{}'""".format(country)
                arcpy.SelectLayerByAttribute_management("rosterJoinTemp", "New_Selection", query)
                arcpy.CopyFeatures_management("rosterJoinTemp", featureCountry)
                arcpy.MakeFeatureLayer_management(featureCountry, "CountryTemp")
                
    # Add heightcm and weightkg fields to the feature class
    arcpy.AddField_management("CountryTemp", "heightcm", "DOUBLE")
    arcpy.AddField_management("CountryTemp", "weightkg", "DOUBLE")
    # Create an update cursor to iterate over the feature class rows
    with arcpy.da.UpdateCursor(featureCountry, ["position", "height", "weight", "heightcm", "weightkg"]) as cursor:
        # Loop through each row in the cursor
        for row in cursor:
            # Check if the player's position matches the desired positions
                for pos in positionList:
                    if row[0] == pos:
                         # Create a feature class for the desired position
                        featurePosition = country + pos
                        query1 = """position = '{}'""".format(pos)
                        arcpy.SelectLayerByAttribute_management("CountryTemp", "New_Selection", query1)
                        arcpy.CopyFeatures_management("CountryTemp", featurePosition)
                        # Add the postions wise shapefile to the output list
                        featureShapefile = featurePosition + ".shp"
                        if featureShapefile not in outputList:
                            outputList.append(featureShapefile)
                        # Convert the height from feet and inches to centimeters and update the heightcm field
                        height = row[1]
                        feet, inches = height.split("'")
                        heightcm = (int(feet)*12 + int(inches.replace('"',''))) * 2.54
                        row[3] = heightcm
                        # Convert the weight from pounds to kilograms and update the weightkg field
                        row[4] = row[2] * 0.453592
                        cursor.updateRow(row)


# Catch any execution errors and print the error message
except arcpy.ExecuteError:
    print(arcpy.GetMessages(2))

# Print the path to the geodatabase
print(f"All the shapefiles are stores in the workspace : {arcpy.env.workspace}")
print(outputList)     
# Ask if the user wants output in excel sheets
question = input("Do you want the output in excel sheets? \nIf yes, Press(Y/y). If no, press anything to end.\nPress::::::::::")
if question.lower() == 'y':
    #create a folder called "excel" within the workspace directory
    excFolder = os.path.join(arcpy.env.workspace,"excelSheets")
    if not os.path.exists(excFolder):
        os.mkdir(excFolder)
    for shapefiles in outputList:
            fields = arcpy.ListFields(shapefiles)
            header = [field.name for field in fields]
            # loop through the three output shapefiles
            for output in outputList:
                    # create a pandas dataframe from the attribute table of the output shapefile
                    data = [row for row in arcpy.da.SearchCursor(output, header)]
                    df = pd.DataFrame(data, columns=header)
                    # export the dataframe to an Excel file
                    output_path = f"{excFolder}/{output.split('.')[0]}.xlsx"
                    df.to_excel(output_path, index=False)
    print(f"The requested excel sheets are saved in : {excFolder}")
                    
                    
                    
print("Thank you. \nYou can change the country name at the top and run the program again if you want to process data for another country.")
      
# Clean the temporary files created
arcpy.Delete_management("rosterTemp")
arcpy.Delete_management("worldTemp")
arcpy.Delete_management("rosterTemp")
arcpy.Delete_management("rosterJoinTemp")
arcpy.Delete_management("CountryTemp")
                                        
                          
            
        
            





