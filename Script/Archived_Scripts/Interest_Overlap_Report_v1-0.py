'''
Created on 2014-03-05
Python v2.7.2
@author: mmacrae

This tool is used to run an Interest Report on lands covered by a shapefile extent or tenure number.
It will clip out data of the area of interest and report on the layer name, an identification field 
for each record (SITE_NAME, 
'''
# Import some modules to use
import arcpy, time, socket
from arcpy import env
import win32com.client
from time import strftime

tic = time.clock()

# Set the parameters for the ArcGIS GUI
clip_feature = arcpy.GetParameterAsText(0)
sqlQuery = arcpy.GetParameterAsText(1)
output_folder = arcpy.GetParameterAsText(2)

## Hard coded versions of parameters
#clip_feature = r"Database Connections\MEMPRD.sde\MTA_GEODB.MTA_ACQUIRED_TENURE_POLY"
#sqlQuery = "MTA_TENURE_TYPE_CODE = 'M' AND TENURE_NUMBER_ID = 200212"
#output_folder = r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\scripts\python\overlap_report\mike_mac\Geodatabases"

# Set some variables to the xls MASTER document, a clip feature for testing, an output location 
# for the clip function and a variable to an extent file name
xls = r"W:\em\vic\mtb\Local\scripts\python\overlap_report\mike_mac\Excel\Statusing_Report_05Mar14.xls\Statusing$"
scratchFolder = r"W:\em\vic\mtb\Local\scripts\python\MTB_Tools\Python_Scripts\Reporting_Tools\Interest_Overlap_Report\ScratchGDB"
clippedGDB = "Clipped_FeatureClasses.gdb"
output_workspace = scratchFolder + "\\" + clippedGDB
extentFile = "extentFile"

## Set the workspace environment of the MTA_GEODB.MTA_ACQUIRED_TENURE_POLY which can be used to query our  area of interest
#env.workspace = r"Database Connections\MEMPRD.sde"
#env.workspace = r"Database Connections\BCGW.sde"

# Set a variable to an empty excel instance
excel = win32com.client.Dispatch("Excel.Application")

# Initialize a workbook within excel
book = excel.Workbooks.Add()

# Create sheet in book
sheet = book.Worksheets(1)

# Set some header values in the sheet and do some formatting on those cells
sheet.Cells(1,1).Value = "Interest Report run on " + strftime('%d%b%y') + " @ " + strftime('%H:%M:%S')
sheet.Cells(1,1).Font.Size = 10

# Layer name Heading and formatting
sheet.Cells(2,1).Value = "Layer Name"
sheet.Cells(2,1).Font.Size = 12
sheet.Cells(2,1).Font.Bold = True

# Unique Field Name and formatting
sheet.Cells(2,2).Value = "Unique Field Name"
sheet.Cells(2,2).Font.Size = 12
sheet.Cells(2,2).Font.Bold = True

# Values in the unique field and formatting
sheet.Cells(2,3).Value = "Field Value"
sheet.Cells(2,3).Font.Size = 12
sheet.Cells(2,3).Font.Bold = True

# Area of each overlapping records
sheet.Cells(2,4).Value = "Area of Overlap (ha)"
sheet.Cells(2,4).Font.Size = 12
sheet.Cells(2,4).Font.Bold = True

# Percentage of each overlapping record
sheet.Cells(2,5).Value = "Percentage of Overlap (%)"
sheet.Cells(2,5).Font.Size = 12
sheet.Cells(2,5).Font.Bold = True

# Adjust columns widths
sheet.Columns(1).ColumnWidth = 35
sheet.Columns(2).ColumnWidth = 20
sheet.Columns(3).ColumnWidth = 20
sheet.Columns(4).ColumnWidth = 22
sheet.Columns(5).ColumnWidth = 30

# Delete the temp workspace if it exists
if arcpy.Exists(scratchFolder + "\\" + clippedGDB):
    arcpy.Delete_management(scratchFolder + "\\" + clippedGDB)        

# Create temp worksapce for clipped data
arcpy.CreateFileGDB_management(scratchFolder, clippedGDB)

# Test to make sure the SQL query exists. If it does, then we need to test to make sure a record gets returned.
# If 0 records return, the user will need to go back and make sure the query returns a record. This will stop 
# the script and prompt the user to update the query
if sqlQuery:
    if arcpy.GetCount_management(arcpy.MakeFeatureLayer_management(clip_feature,"extent",sqlQuery)) > 0:
        pass
    else:
        raise Exception("SQL Query returns empty results. Please redefine query, validate to see that more than one record returns and run tool again.")
else:
    pass

# Export the area of interest to our temp geodatabse. This will be used as a clip feature and to help calculate
# areas and percentages
arcpy.FeatureClassToFeatureClass_conversion(clip_feature, output_workspace, extentFile, sqlQuery)

# Set a variable to get the total area of our area of interest. This will be used to calculate percentage
# of overlap of the overlapping polygons
for row in arcpy.SearchCursor(output_workspace + "\\" + extentFile):
    extentArea = row.Shape_Area/10000

## Set workspace to geodatabase
#env.workspace = output_folder

# Set a counter to be used to increment rows in the excel sheet
excelrow = 4

# Cursor Search xls file
for row in arcpy.SearchCursor(xls):

    print row.getValue("Layer Type")
    arcpy.AddMessage("Processing Layer " + row.getValue("Layer Type"))
    
    # Conditional sentece to exclude a couple layers we aren't interested in at this point
    if row.getValue("Layer Type") not in ("Dominion_Coal_Block", "Freehold_Coal", None):
       
        # Set a variable to pull out the full path of the shapefile or FC and set a variable to the source name
        dataSourcePath = str(row.getValue("workspace_path")) + "\\" + str(row.getValue("dataSource"))
        dataSource = row.getValue("dataSource")
        
        # set workspace environemnt to match the workspace location detailed in the xls file
        env.workspace = row.getValue("workspace_path")      

        # Let's clip some data and send it to our scratch workspace         
        arcpy.Clip_analysis(dataSourcePath,output_workspace + "\\" + extentFile, output_workspace + "\\" + str(row.getValue("Layer Type")))

        # Get a count of the records in the clipped data
        result = int(arcpy.GetCount_management(output_workspace + "\\" + str(row.getValue("Layer Type"))).getOutput(0))
        
        # Test to see if there are any records within each clipped feature class
        # If it is zero, then let's output the layer name and a message indicating "No Overlap"
        # We'll also do some formatting on the cells
        if result == 0:
            #print str(row.getValue("Identifying_Field1"))                      
            excel.Cells(excelrow,1).Value = str(row.getValue("Layer Type"))
            excel.Cells(excelrow,2).Value = "No overlap"
            sheet.Cells(excelrow,1).Font.Size = 10
            sheet.Cells(excelrow,2).Font.Size = 10
            excelrow += 1

        # Otherwise, if there is data, let's output the Layer Name, Attribute Field Name, Attributes in that field,
        # are of each record and the percentage of overlap. 
        # Also, do some formatting on the cells
        else:
            layerTotalArea = 0
            for rowcell in arcpy.SearchCursor(output_workspace + "\\" + str(row.getValue("Layer Type"))):
                excel.Cells(excelrow,1).Value = str(row.getValue("Layer Type"))
                excel.Cells(excelrow,2).Value = (str(row.getValue("Identifying_Field1")))
                excel.Cells(excelrow,3).Value = rowcell.getValue(str(row.getValue("Identifying_Field1")))                
                excel.Cells(excelrow,4).Value = round(rowcell.getValue(','.join([str(field.name)for field in arcpy.ListFields(output_workspace + "\\" + str(row.getValue("Layer Type"))) if field.name in ["Shape_Area","GEOMETRY_Area", "SHAPE_Area"]]))/10000,2)
                excel.Cells(excelrow,5).Value = round(rowcell.getValue(','.join([str(field.name)for field in arcpy.ListFields(output_workspace + "\\" + str(row.getValue("Layer Type"))) if field.name in ["Shape_Area","GEOMETRY_Area", "SHAPE_Area"]]))/10000/extentArea*100,1)
                sheet.Cells(excelrow,1).Font.Size = 10
                sheet.Cells(excelrow,2).Font.Size = 10
                sheet.Cells(excelrow,3).Font.Size = 10
                sheet.Cells(excelrow,4).Font.Size = 10
                sheet.Cells(excelrow,5).Font.Size = 10
                excel.Cells(excelrow,2).Font.ColorIndex = 3
                excel.Cells(excelrow,2).HorizontalAlignment = win32com.client.constants.xlRight

                layerTotalArea += round(rowcell.getValue(','.join([str(field.name)for field in arcpy.ListFields(output_workspace + "\\" + str(row.getValue("Layer Type"))) if field.name in ["Shape_Area","GEOMETRY_Area", "SHAPE_Area"]])),2)

                excelrow += 1
                
            excel.Cells(excelrow,3).Value = "Total Area of Overlap"
            excel.Cells(excelrow,3).HorizontalAlignment = win32com.client.constants.xlRight
            excel.Cells(excelrow,3).Font.ColorIndex = 3
            sheet.Cells(excelrow,3).Font.Bold = True
            #sheet.Cells(excelrow,3).BorderAround()
            sheet.Cells(excelrow,3).Interior.ColorIndex = 15
            excel.Cells(excelrow,4).Value = layerTotalArea/10000
            excel.Cells(excelrow,4).HorizontalAlignment = win32com.client.constants.xlRight
            excel.Cells(excelrow,4).Font.ColorIndex = 3
            sheet.Cells(excelrow,4).Font.Bold = True
            #sheet.Cells(excelrow,4).BorderAround()
            sheet.Cells(excelrow,4).Interior.ColorIndex = 15
            excelrow += 1
        excelrow += 1          

# Save the excel spreadsheet as a .xlsx file. This may be useful to pull in macro calls later on.
book.SaveAs(output_folder + "\\" + "Interest_report_" + strftime('%d%b%y') + ".xlsx")

# Quit the instance of excel from win32com
excel.Quit()

# Delete some variables that were causing issues earlier
del excel, book, sheet, excelrow

# Output end time of process
toc = time.clock()

# Print out time informaton (total seconds/60 for rough amount of minutes
print (toc - tic)/60
