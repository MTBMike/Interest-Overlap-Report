'''
Created on 2014-03-05
Python v2.7.2
@author: mmacrae

This tool is used to run an Interest Report on lands covered by a shapefile extent, reserve or tenure.
It will clip out from the area of interest various layers outlined in a spreadsheet and report on the layer name and summarize chosen fields 
for each record (SITE_NAME, 
'''
# Import some modules to use
import arcpy, time, ntpath, datetime, itertools, win32com.client
from arcpy import env
from time import strftime

# Initiate a time counter
tic = time.clock()

# Set workspaces for login.
env.workspace = r"Database Connections\BCGW.sde"
env.workspace = r"Database Connections\MEMPRD.sde"

# Set the parameters for the ArcGIS GUI
clip_feature = arcpy.GetParameterAsText(0)
sqlQuery = arcpy.GetParameterAsText(1)
shFieldList = arcpy.GetParameter(2)
layerList = arcpy.GetParameterAsText(3)
output_folder = arcpy.GetParameterAsText(4)

#shFieldListDelim = ";".join(shFieldList)


## Hard coded versions of parameters
#clip_feature = r"W:\em\vic\mtb\Local\scripts\python\overlap_report\mike_mac\KML\REC191231.gdb\Placemarks\Polylines_Buffer"
##clip_feature = r"G:\Reserves\Reserve Requests\MARR - Kitselas_Kitsumkalum\Kitselas_AiP_2013-07-15.shp"
#sqlQuery = "OBJECTID = 1"
##sqlQuery = ""
#shFieldList = ""
#layerList = ['Agriculture Land Reserve', 'Archaeology Sites', 'Archaeology Sites within 50 metres', 'Crown Tenure - Transfer of Administrative Control', 'Forest District', 'Guide Outfitter Territoires', 'Historic Sites ', 'Land Act Subdivisions', 'Land Districts', 'Land Title Districts', 'Mountain Pine Beetle Salvage Area', 'Municipalities', 'Natural Resource Operations Admin Areas', 'Natural Resource Operations Operating Regions', 'Primary Survey Parcels', 'Regional Districts', 'Strategic Land Resource Planning Area', 'Surveyed Rights of Way', 'Crown Land Leases Tantalis ', 'Crown Rights of Way', 'Crown Tenure Inventory', 'Crown Tenure PreTantalis', 'Crown Tenure Tantalis Application', 'Crown Tenure Tantalis Tenure', 'Surveyed Rights of Way', 'Consultative Areas', 'First Nation Traditional  Use Study Area', 'First Nation Treaty Areas', 'First Nation Treaty Lands', 'First Nation Treaty Related Lands', 'Indian Reserves', 'MFR First Nation Agreement Boundaries', 'Traditional Use Lines', 'Traditional Use Points', 'Traditional Use Polygons', 'Cutblocks', 'Tree Farm Licence', 'BCGS Grid', 'NTS Grid', 'Coal Bed', 'Coal Grid Units', 'Coal Land Reserves', 'Coal Titles', 'Crown Granted 2 Post Mineral Claims', 'Freehold Coal - Dominion', 'Freehold Coal - Elk Valley Coal Partnerships', 'Freehold Coal - Fording, Elk Valley', 'Mineral Land Reserves - Mineral', 'Mineral Land Reserves - Placer', 'Mineral Tenures', 'Mining Division', 'MTO Grid Cells', 'Notice of Work', 'Placer Tenures', 'Oil and Gas Pipeline  Rights of Way', 'Petroleum & Natural Gas Tenure', 'Community Watersheds', 'Drinking Water Extraction Points', 'Drinking Water Points of Diversion', 'Environmental Remediation Sites', 'FTEN Real Property Projects', 'Ground Water Aquifers', 'Legal OGMA', 'Map Notation Lines', 'Map Notation Points', 'Map Notation Polys', 'Recreation Sites', 'Recreation Trails (1)', 'Recreation Trails (2)', 'Reservoirs Permits over Crown Land', 'Water Licenced Linear Features', 'Water Licenced Work Point Features', 'Water Reserves', 'Water Wells', 'Conservancy Areas', 'Conservation Lands including Wildlife Managment Areas', 'National Park', 'Park and Protected Areas', 'Regional Parks', 'Integrated Cadastral Fabric Restricted Access', 'Trapline', 'Ungulate Winter Range Legal', 'Wildlife Habitat Area_Legal', 'Wildlife Management Areas - Tantalis']
#output_folder = r"W:\em\vic\mtb\Local\scripts\python\overlap_report\mike_mac\Test_Excel\Sample_Reports_April3_2014"

arcpy.AddMessage(shFieldList)
arcpy.AddMessage(len(shFieldList))

# Set some variables to the xls MASTER document, a clip feature for testing, an output location 
# for the clip function and a variable to an extent file name
xls = r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\ArcGIS_Tools\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Excel_Spreadsheets\InterestReport_Layer_List_MASTER.xls\'MineralTitles Dataset selection$'
scratchFolder = r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\ArcGIS_Tools\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\ScratchGDB"

# set some variables as shortcuts to avoid long paths names
clippedGDB = "Clipped_FeatureClasses.gdb"
output_workspace = scratchFolder + "\\" + clippedGDB

# create a variable to a lsit of land types that are considered alienated lands
alienatedLandsList = ("MTA_SPATIAL.MTA_SITE_POLY","WHSE_ADMIN_BOUNDARIES.CLAB_INDIAN_RESERVES",
					  "WHSE_ADMIN_BOUNDARIES.CLAB_NATIONAL_PARKS","WHSE_LEGAL_ADMIN_BOUNDARIES.FNT_TREATY_LAND_SP",
					  "WHSE_LEGAL_ADMIN_BOUNDARIES.FNT_TREATY_RELATED_LAND_SP","WHSE_TANTALIS.TA_CONSERVANCY_AREAS_SVW",
					  "WHSE_TANTALIS.TA_INTEREST_PARCEL_SHAPES","WHSE_TANTALIS.TA_PARK_ECORES_PA_SVW","Dominion.shp",
					  "Elk_Valley_Coal_Partnership.shp","freehold_fording_elk_valley.shp")

# initiate a message to start creation of the scratch workspace
arcpy.AddMessage("Setting scratch workspace")

# Delete the temp workspace if it exists
if arcpy.Exists(scratchFolder + "\\" + clippedGDB):
	arcpy.Delete_management(scratchFolder + "\\" + clippedGDB)        

# Create temp workspace for clipped data
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
arcpy.AddMessage("Exporting area of interest extent to scratch workspace")
arcpy.FeatureClassToFeatureClass_conversion(clip_feature, output_workspace, "extentFile", sqlQuery)
extentFile = output_workspace + "\\" + "extentFile"

# Set a variable to get the total area of our area of interest. This will be used to calculate percentage
# of overlap of the overlapping polygons
for row in arcpy.SearchCursor(extentFile):
	arcpy.AddMessage("Inside for loop")
	extentArea = row.getValue(str(arcpy.Describe(extentFile).AreaFieldName))/10000
	arcpy.AddMessage("Extent Area = " + str(extentArea))

# Set a variable to an empty excel instance
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True

# Initialize a workbook within excel
book = excel.Workbooks.Add()

# Set first sheet in book and rename it for the report
sheet = book.Worksheets(1)

# Set sheet name
sheet.Name = "Interest Report"

# Hide Gridlines in sheet
excel.ActiveWindow.DisplayGridlines = False

# Add logo to report and set size and location
pic = sheet.Pictures().Insert(r"\\zipline\S6203\rddTB.shr\Mineral-Titles-Br\Admin\logos\approved_logos\BCID+ENER\2013_BC_ENER\ENER\BC_ENER_H_CMYK_pos.jpg")

# Format size of the logo
pic.Height = 418
pic.Width = 218

# Format location of the logo
cell = sheet.Cells(2,1)
pic.Left = cell.Left
pic.Top = cell.Top

# Initiate some variables for the worksheets rows and columns
excelrow = 1
excelcol = 1

# Set the "REPORT FOR INTERNAL USE ONLY" comment in the sheet and do some formatting
sheet.Cells(excelrow,excelcol + 2).Value = "REPORT FOR INTERNAL USE ONLY"
sheet.Cells(excelrow,excelcol + 2).Font.Size = 10
sheet.Cells(excelrow,excelcol + 2).Font.Bold = True
sheet.Cells(excelrow,excelcol + 2).Font.ColorIndex = 3

# Set a comment to indicate when the report was run and do some formatting
sheet.Cells(excelrow + 1,excelcol + 2).Value = "Report run on " + strftime('%d%b%y') + " @ " + strftime('%I:%M:%S')
sheet.Cells(excelrow + 1,excelcol + 2).Font.Size = 10
sheet.Cells(excelrow + 1,excelcol + 2).Font.Italic = True

# Set a conditional sentence to determine what type of input is used for the area of interest
# In the first conditional sentence, determine if the area of interest is not the tenure or reserve files
if ntpath.basename(clip_feature) not in [r'MTA_SPATIAL.MTA_ACQUIRED_TENURE_SVW', r'MTA_SPATIAL.MTA_SITE_SVW']:
	arcpy.AddMessage(ntpath.basename(clip_feature))
	
	# Set a conditional sentence to determine if the user chose any fields to summarize in the area of interest parameter
	# In the first condition, determine if there are no fields to summarize
	if len(shFieldList) == 0:
		
		# Increment the counter by 8. This will push the rows paste the title header and the image
		# Also, set another variable to be used as a static row count which will be used to set the position of the legend
		# so that the legend lies next to the header information
		excelrow += 8
		staticexcelrow = excelrow - 4
		
#        sheet.Cells(excelrow,excelcol).Value = "Area (ha):"
#        sheet.Cells(excelrow,excelcol).Font.Bold = True
#        sheet.Cells(excelrow,excelcol + 1).Value = cliprow.getValue("AREA_IN_HECTARES")       
		
		# Set the header above the list of layers of overlaps and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "INTEREST OVERLAPS"
		sheet.Cells(excelrow,excelcol).Font.Size = 12
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol).Font.Italic = True  
		
		# Increment the row count by 2 for the position of the first overlapping layer
		excelrow += 2
		
	# End of conditional sentence. If the user chose some fields from the area of interest, update the header
	else:
		
		# Increment the counter by 8. This will push the rows paste the title header and the image
		# Also, set another variable to be used as a static row count which will be used to set the position of the legend
		# so that the legend lies next to the header information
		excelrow += 8
		staticexcelrow = excelrow
		
		# Loop over the field list from the fields chosen by the user
		for shField in shFieldList:

			# Add each field name in the first column under the image and do some formatting
			sheet.Cells(excelrow,excelcol).Value = shField.replace("_"," ")
			sheet.Cells(excelrow,excelcol).Font.Bold = True
			sheet.Cells(excelrow,excelcol).Font.Italic = True
			
			# Loop through the area of interest file and add the value that corresponds to the field name
			for shrow in arcpy.SearchCursor(arcpy.MakeFeatureLayer_management(clip_feature, "clipAOI", sqlQuery)):
				sheet.Cells(excelrow,excelcol + 1).Value = str(shrow.getValue(str(shField)))
			
			if arcpy.Exists("clipAOI"):
				arcpy.Delete_management("clipAOI")
			# increment the row count for the next field to be added
			excelrow += 1
		
		# Do some formatting on the area of interest header information. Set the font size, border, fill color and alignment
		sheet.Range(sheet.Cells(staticexcelrow,excelcol),sheet.Cells(excelrow - 1,2)).Font.Size = 10
		sheet.Range(sheet.Cells(staticexcelrow,excelcol),sheet.Cells(excelrow - 1,2)).BorderAround()
		sheet.Range(sheet.Cells(staticexcelrow,excelcol),sheet.Cells(excelrow - 1,2)).Interior.ColorIndex = 36
		sheet.Range(sheet.Cells(staticexcelrow,excelcol + 1),sheet.Cells(excelrow - 1,2)).HorizontalAlignment = win32com.client.constants.xlRight                            
		
		# Increase the row count by 2 to buffer between the area of interest header information and the next header
		excelrow += 2
		
		# Set the header above the list of layers of overlaps and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "INTEREST OVERLAPS"
		sheet.Cells(excelrow,excelcol).Font.Size = 12
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol).Font.Italic = True
		
		# Increment the row count by 2 for the position of the first overlapping layer
		excelrow += 2     
		
	# set a varaible to the input file name to be used later in naming the filename for the report.
	# This will help identify where the information came from for the report
	if ntpath.basename(clip_feature).endswith(".shp") and sqlQuery <> "":
		
		# if the area of interest file is a shapefile, truncate the file name to remove the .shp extension
		inputSourceName = ntpath.basename(clip_feature)[:-4]  + "_" + sqlQuery
		
	elif ntpath.basename(clip_feature).endswith(".shp") and sqlQuery == "":
		inputSourceName = ntpath.basename(clip_feature)[:-4] 
	else:
		if sqlQuery == "":
		# otherwise, just use the file name for the area of interest
			inputSourceName = ntpath.basename(clip_feature)
		else:
			inputSourceName = ntpath.basename(clip_feature) + "_" + sqlQuery
	arcpy.AddMessage("inputSourceName = " + inputSourceName) 
# next conditional sentence is to determine of the user chose the tenure feature class
elif ntpath.basename(clip_feature) == 'MTA_SPATIAL.MTA_ACQUIRED_TENURE_SVW':
	
	arcpy.AddMessage(ntpath.basename(clip_feature))
	
	# Loop over the area of interest extent file that was exported to the scratch workspace earlier in the script
	# to pull out some header information on the tenure file
	for cliprow in arcpy.SearchCursor(extentFile):
		
		
		# Increment the counter by 8. This will push the rows paste the title header and the image
		# Also, set another variable to be used as a static row count which will be used to set the position of the legend
		# so that the legend lies next to the header information        
		excelrow += 8
		staticexcelrow = excelrow
		
		# Set the tenure number into the header area and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "Tenure Number:"
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol + 1).Value = str(cliprow.getValue("TENURE_NUMBER_ID"))
		
		# increment to the next row
		excelrow += 1
		
		# Set the Title Type Description into the header area and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "Title Type Description:"
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol + 1).Value = cliprow.getValue("TITLE_TYPE_DESCRIPTION")
		
		# increment to the next row
		excelrow += 1
		
		# Set the Issue Date into the header area and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "Issue Date:"
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol + 1).Value = cliprow.getValue("ISSUE_DATE")
		
		# increment to the next row
		excelrow += 1
		
		# Set the Good To Date into the header area and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "Good To Date:"
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol + 1).Value = cliprow.getValue("GOOD_TO_DATE")
		
		# increment to the next row
		excelrow += 1
		
		# Set the Owner information into the header area and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "Owner:"
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol + 1).Value = cliprow.getValue("OWNER_NAME")
		sheet.Cells(excelrow,excelcol + 1).WrapText = True
		
		# increment to the next row
		excelrow += 1
		
		# Set the area information from the AREA_IN_HECTARES field into the header area and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "Area (ha):"
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol + 1).Value = cliprow.getValue("AREA_IN_HECTARES")
		
		# increment to the next row
		excelrow += 1
		  
		# Set the area information as determined by the AreaFieldName spatial field into the header area and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "True Geometry Area (ha):"
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol + 1).Value = round(cliprow.getValue("GEOMETRY_AREA")/10000,2)
		
		# Do some formatting on the area of interest header information. Set the font size, border, fill color and alignment
		sheet.Range(sheet.Cells(staticexcelrow,excelcol),sheet.Cells(excelrow,2)).Font.Size = 10
		sheet.Range(sheet.Cells(staticexcelrow,excelcol),sheet.Cells(excelrow,2)).BorderAround()
		sheet.Range(sheet.Cells(staticexcelrow,excelcol),sheet.Cells(excelrow,2)).Interior.ColorIndex = 36
		sheet.Range(sheet.Cells(staticexcelrow,excelcol + 1),sheet.Cells(excelrow,2)).HorizontalAlignment = win32com.client.constants.xlRight        
		
		# Increase the row count by 2 to buffer between the area of interest header information and the next header
		excelrow += 2
		
		# Set the header above the list of layers of overlaps and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "INTEREST OVERLAPS"
		sheet.Cells(excelrow,excelcol).Font.Size = 12
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol).Font.Italic = True
		
		# Create variable to be used in file path name
		inputSourceName = sqlQuery 
		
		# Increment the row count by 2 for the position of the first overlapping layer
		excelrow += 2

# Last conditional sentence is to determine of the user chose the reserves feature class
elif ntpath.basename(clip_feature) == 'MTA_SPATIAL.MTA_SITE_SVW':
	
	arcpy.AddMessage(ntpath.basename(clip_feature))
	
	# Loop over the area of interest extent file that was exported to the scratch workspace earlier in the script
	# to pull out some header information on the reserves file
	for cliprow in arcpy.SearchCursor(extentFile):
		
		# Increment the counter by 8. This will push the rows paste the title header and the image
		# Also, set another variable to be used as a static row count which will be used to set the position of the legend
		# so that the legend lies next to the header information        
		excelrow += 8
		staticexcelrow = excelrow        
		
		# Set the reserve site ID number into the header area and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "Site Number ID:"
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol + 1).Value = str(cliprow.getValue("SITE_NUMBER_ID"))
		
		# increment to the next row
		excelrow += 1
		
		# Set the reserve stype into the header area and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "Reserve Type:"
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol + 1).Value = cliprow.getValue("RESERVE_TYPE")
		sheet.Cells(excelrow,excelcol + 1).WrapText = True
		
		# increment to the next row
		excelrow += 1
		
		# Set the site order restriction description into the header area and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "Site order restriction description:"
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol + 1).Value = cliprow.getValue("MTA_SITE_ORDER_RESTR_DESC")
		
		# increment to the next row
		excelrow += 1
		
		# Set the area information from the AREA_IN_HECTARES field into the header area and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "Area (ha):"
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol + 1).Value = cliprow.getValue("TOTAL_AREA")
		
		# increment to the next row
		excelrow += 1
		
		# Set the area information as determined by the AreaFieldName spatial field into the header area and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "True Geometry Area (ha):"
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol + 1).Value = round(cliprow.getValue("Shape_Area")/10000,2)
		
		# Do some formatting on the area of interest header information. Set the font size, border, fill color and alignment
		sheet.Range(sheet.Cells(staticexcelrow,excelcol),sheet.Cells(excelrow,2)).Font.Size = 10
		sheet.Range(sheet.Cells(staticexcelrow,excelcol),sheet.Cells(excelrow,2)).BorderAround()
		sheet.Range(sheet.Cells(staticexcelrow,excelcol),sheet.Cells(excelrow,2)).Interior.ColorIndex = 36
		sheet.Range(sheet.Cells(staticexcelrow,excelcol + 1),sheet.Cells(excelrow,2)).HorizontalAlignment = win32com.client.constants.xlRight
		
		# Increase the row count by 2 to buffer between the area of interest header information and the next header
		excelrow += 2
		
		# Set the header above the list of layers of overlaps and do some formatting
		sheet.Cells(excelrow,excelcol).Value = "INTEREST OVERLAPS"
		sheet.Cells(excelrow,excelcol).Font.Size = 12
		sheet.Cells(excelrow,excelcol).Font.Bold = True
		sheet.Cells(excelrow,excelcol).Font.Italic = True
		
		# Create variable to be used in file path name
		inputSourceName = "RESERVE_" + sqlQuery

		# Increment the row count by 2 for the position of the first overlapping layer
		excelrow += 2
		
		
# Build the legend of the report

# Set the Legend Header and do some formatting
sheet.Cells(staticexcelrow - 1,excelcol + 3).Value = "Legend"
sheet.Cells(staticexcelrow - 1,excelcol + 3).Font.Size = 12
sheet.Cells(staticexcelrow - 1,excelcol + 3).Font.Bold = True
sheet.Cells(staticexcelrow - 1,excelcol + 3).Font.Underline = True

# Set color (grey) for the first item in the legend
sheet.Cells(staticexcelrow,excelcol + 3).Interior.ColorIndex = 15

# Set "Legend Category" next to the grey color and do some formatting
sheet.Cells(staticexcelrow,excelcol + 4).Value = "Layer Category"
sheet.Cells(staticexcelrow,excelcol + 4).Font.Size = 10
sheet.Cells(staticexcelrow,excelcol + 4).Font.Bold = True

# Set color (pale blue) for the second item in the legend
sheet.Cells(staticexcelrow + 1,excelcol + 3).Interior.ColorIndex = 37

# Set "Indicates Overlapping Layer" next to the pale blue color and do some formatting
sheet.Cells(staticexcelrow + 1,excelcol + 4).Value = "Indicates Overlapping Layer"
sheet.Cells(staticexcelrow + 1,excelcol + 4).Font.Size = 10
sheet.Cells(staticexcelrow + 1,excelcol + 4).Font.Bold = True

# Set an example or an alientated land and the orange color for the third item in the legend
sheet.Cells(staticexcelrow + 2,excelcol + 3).Value = "National Park"
sheet.Cells(staticexcelrow + 2,excelcol + 3).Font.Size = 10
sheet.Cells(staticexcelrow + 2,excelcol + 3).Font.ColorIndex = 46 

# Set "Indicates Alienated Lands" next to the orange color and do some formatting
sheet.Cells(staticexcelrow + 2,excelcol + 4).Value = "Indicates Alienated Lands"
sheet.Cells(staticexcelrow + 2,excelcol + 4).Font.Size = 10
sheet.Cells(staticexcelrow + 2,excelcol + 4).Font.Bold = True

# Place a border around the legend
sheet.Range(sheet.Cells(staticexcelrow,excelcol + 3),sheet.Cells(staticexcelrow + 2,excelcol + 4)).BorderAround()

# Initiate an empty python dictionary container. This will build Category/Feature Class rdictionary using the category
# as the key and the feature classes as the values. This will help dynamically sort and organize the layers 
layerListDict = dict()

# Cursor Search xls file
for row in arcpy.SearchCursor(xls):  
	
	# Set condition to pull out the Featureclass_Name field in the xls file
	if str(row.getValue("Featureclass_Name")).rstrip() in layerList:
			
		# Determine if the category item is in the dictionary as a key already. If so, then append the Featureclass_Name to the list of values associated with the category 
		if row.getValue("Category") in layerListDict:
			layerListDict[row.getValue("Category")].append(str(row.getValue("Featureclass_Name")))
		# if not, create a new category key and add the associated Featureclass_Name value to it
		else:
			layerListDict[row.getValue("Category")] = [str(row.getValue("Featureclass_Name"))]

# Build a list of fields in the xls spreadsheet that contain the fields to summarize
fieldList = [str(field.name) for field in arcpy.ListFields(xls) if "Summarize" in field.name]

# Set a variable to help determine how many columns of information the report will contain. This will help determine how to 
# set the outline border extent at the end of the script
maxcolCount = 0

mineral_Coal_first = 'Mineral/Coal'

# iterate over the layerListDict python dictionary, sorted by categories
#for key,valueList in sorted(layerListDict.iteritems()):
for key, valueList in itertools.chain([(mineral_Coal_first, layerListDict[mineral_Coal_first])],((key,valueList) for (key,valueList) in sorted(layerListDict.iteritems()) if key != mineral_Coal_first)):

	print key, valueList
	print excelrow
	# Reset the column count to 1
	excelcol = 1
	
	# Set the first category to the first key from the dictionary and do some formatting
	sheet.Cells(excelrow,excelcol).Value = str(key)
	sheet.Cells(excelrow,excelcol).Font.Size = 12
	sheet.Cells(excelrow,excelcol).Font.Bold = True
	sheet.Cells(excelrow,excelcol).Interior.ColorIndex = 15
	
	# increment the row count by 2
	excelrow += 2
	
	# Loop over each value item from the value list related to the key
	for valueItem in sorted(valueList):       
		print valueItem
		# set a cursor object to cursor search the xls file
		cursor = arcpy.da.SearchCursor(xls,"*")
		c = 0
		d = {}        
		
		# Loop over the field names in the xls and Build a python dictionary of the field headings
		for fieldda in cursor.fields:
			d[fieldda] = c
			c += 1
		
		# Loop over the xls file to pull out data for analysis
		for row in cursor:
			
			# Because we do not have access to archaeology data, this is a temporary condition to exclude those layers for now
			if row[d["Featureclass_Name"]] == valueItem:# and row[d["Featureclass_Name"]] not in ("Archaeology Sites","Archaeology Sites within 50 metres", None):
				
				# Add a message to ArcGIS GUI to notify the user of which layer is being processed
				arcpy.AddMessage("Processing Layer " + str(row[d["Featureclass_Name"]]).replace("_"," "))
					 
				# Set a variable to pull out the full path of the shapefile or FC and set a variable to the source name
				dataSourcePath = str(row[d["workspace_path"]]) + "\\" + str(row[d["dataSource"]])
				dataSource = row[d["dataSource"]]
				
				# Translate the value of the Featureclass_Name item to repalce any non characters into underscores. We do this because
				# the name of the clipped feature below cannot have non characters (ie the periods '.' used in the sde feature class names)
				featureclassNameTrans = ''.join(chr(c) if chr(c).isupper() or chr(c).islower() else '_' for c in range(256))
				
				# Set a variable to the output workspace pathname and the translated feature class name
				outClipFC = output_workspace + "\\" + str(row[d["Featureclass_Name"]]).translate(featureclassNameTrans)
	
				# Create feature layer in order to apply a definition query to the dataset
				arcpy.MakeFeatureLayer_management(dataSourcePath, "dsPath", row[d["Definition_Query"]])
				
				# Create a select by location to test for overlap.								
				selFeats = arcpy.SelectLayerByLocation_management("dsPath", "intersect", extentFile)
				
				# Test to see if there are any records within each selected feature class
				# If it is zero, then let's output the layer name and a message indicating "No Overlap"
				# We'll also do some formatting on the cells								
				if int(arcpy.GetCount_management(selFeats).getOutput(0)) == 0:
						
					# Add a message to to the user to indicate the feature class returned no selected values, therefore no overlap
					arcpy.AddMessage(str(row[d["Featureclass_Name"]]).replace("_"," ") + ' Layer is empty')
					
					excelcol = 1
					
					sheet.Cells(excelrow,excelcol).Value = str(row[d["Featureclass_Name"]])
					sheet.Cells(excelrow,excelcol).Font.Size = 8
					
					# Set font color for alienated lands datasets
					if str(row[d["dataSource"]]) in alienatedLandsList:
							
						sheet.Cells(excelrow,excelcol).Font.ColorIndex = 46			 
					
					sheet.Cells(excelrow,excelcol + 1).Value = "No overlap"
					sheet.Cells(excelrow,excelcol + 1).Font.Size = 8
					sheet.Cells(excelrow,excelcol + 1).HorizontalAlignment = win32com.client.constants.xlRight
					
					excelrow += 1					 
	
				# Otherwise, if there is data, let's clip the data, output the Layer Name, Attribute Field Name, Attributes in that field,
				# are of each record and the percentage of overlap. 
				# Also, do some formatting on the cells
				else:
						
					arcpy.AddMessage(str(row[d["Featureclass_Name"]]).replace("_"," ") + ' Layer is being clipped')
					# test if a Join is needed. If so, add join table
					if row[d["Join_Table"]]:
						arcpy.AddJoin_management("dsPath", str(row[d["dataSource_Join_Field"]]), str(row[d["Join_Table"]]), str(row[d["Join_Table_Field"]]))
										 
					# Let's clip some data and send it to our scratch workspace				 
					arcpy.Clip_analysis("dsPath",extentFile, outClipFC)
	
					fieldValueList = [str(row[d[fieldName]]).replace(".","_") for fieldName in fieldList if str(row[d[fieldName]]) <> 'None']
					
					if maxcolCount < len(fieldValueList):
						maxcolCount = len(fieldValueList)									
					
					if arcpy.Describe(outClipFC).ShapeType == "Polygon":
						areaFieldName = arcpy.Describe(outClipFC).AreaFieldName
						fieldValueList.append(str(areaFieldName))
					elif arcpy.Describe(outClipFC).ShapeType == "Polyline":
						lengthFieldName = arcpy.Describe(outClipFC).LengthFieldName
						fieldValueList.append(str(lengthFieldName))
					else:
						pass
					
					with arcpy.da.SearchCursor(outClipFC,fieldValueList) as rows:
	
						excelcol = 2												
						for field in rows.fields:
								
							sheet.Cells(excelrow,excelcol).Value = field.replace("_"," ")
							sheet.Cells(excelrow,excelcol).Font.Size = 10
							sheet.Cells(excelrow,excelcol).Font.Bold = True
							sheet.Cells(excelrow,excelcol).Interior.ColorIndex = 37								
							sheet.Cells(excelrow,excelcol).HorizontalAlignment = win32com.client.constants.xlRight
							
							sheet.Cells(excelrow,1).Interior.ColorIndex = 37														 
							
							if 'areaFieldName' in locals():
									
								if field == str(areaFieldName):
						
									areacol = len(fieldValueList) + 1
									
									sheet.Cells(excelrow,areacol).Value = "AREA OF OVERLAP (HA)"
									sheet.Cells(excelrow,areacol).Font.Size = 10
									sheet.Cells(excelrow,areacol).Font.Bold = True
									sheet.Cells(excelrow,areacol).WrapText = True
									sheet.Cells(excelrow,areacol).Interior.ColorIndex = 37																	 
									sheet.Cells(excelrow,areacol).HorizontalAlignment = win32com.client.constants.xlCenter
	
									
									percentcol = len(fieldValueList) + 2
																								 
									sheet.Cells(excelrow,percentcol).Value = "PERCENTAGE OF OVERLAP"
									sheet.Cells(excelrow,percentcol).Font.Size = 10
									sheet.Cells(excelrow,percentcol).Font.Bold = True
									sheet.Cells(excelrow,percentcol).WrapText = True
									sheet.Cells(excelrow,percentcol).Interior.ColorIndex = 37																		
									sheet.Cells(excelrow,percentcol).HorizontalAlignment = win32com.client.constants.xlCenter
											
											
							elif 'lengthFieldName' in locals():
									
								if field == str(lengthFieldName):
								
									lengthcol = len(fieldValueList) + 1
									
									sheet.Cells(excelrow,lengthcol).Value = "LENGTH OF OVERLAP (m)"
									sheet.Cells(excelrow,lengthcol).Font.Size = 10
									sheet.Cells(excelrow,lengthcol).Font.Bold = True
									sheet.Cells(excelrow,lengthcol).WrapText = True
									sheet.Cells(excelrow,lengthcol).Interior.ColorIndex = 37																		
									sheet.Cells(excelrow,lengthcol).HorizontalAlignment = win32com.client.constants.xlCenter
											
							else:
								pass
																									
							excelcol += 1
								
						excelrow += 1
						
						for rowoutClipFC in rows:
								
							fieldnum = 0
							excelcol = 1
							
							sheet.Cells(excelrow,excelcol).Value = str(row[d["Featureclass_Name"]])
							sheet.Cells(excelrow,excelcol).Font.Size = 8
							
							if str(row[d["dataSource"]]) in alienatedLandsList:
									
								sheet.Cells(excelrow,excelcol).Font.ColorIndex = 46
								
	#                        else:
	#                            pass											
							
							while fieldnum <= len(rows.fields) - 1:																
	
								if 'areaFieldName' in locals():
	
									for field in rows.fields:
													
										if field == areaFieldName:
	
											sheet.Cells(excelrow,areacol).Value = round(rowoutClipFC[rows.fields.index(areaFieldName)]/10000,2)
											sheet.Cells(excelrow,areacol).Font.Size = 8
											sheet.Cells(excelrow,areacol).HorizontalAlignment = win32com.client.constants.xlRight
											
											excel.Cells(excelrow,percentcol).Value = round(rowoutClipFC[rows.fields.index(areaFieldName)]/10000/extentArea*100,1)
											sheet.Cells(excelrow,percentcol).Font.Size = 8
										
										else:
										
											sheet.Cells(excelrow,fieldnum + 2).Value = rowoutClipFC[fieldnum]																
											sheet.Cells(excelrow,fieldnum + 2).Font.Size = 8
											sheet.Cells(excelrow,fieldnum + 2).HorizontalAlignment = win32com.client.constants.xlRight
											
											if rowoutClipFC[fieldnum] == None:
												pass
											elif type(rowoutClipFC[fieldnum]) in (int, float) or isinstance(rowoutClipFC[fieldnum], datetime.datetime):
												lenrowoutClipFC = len(str(rowoutClipFC[fieldnum]))
											else:
												lenrowoutClipFC = len(rowoutClipFC[fieldnum])
											if lenrowoutClipFC >= 38:
												sheet.Cells(excelrow,fieldnum + 2).WrapText = True
											else:
												pass
																	
								elif 'lengthFieldName' in locals():
										
									for field in rows.fields:																	 
									
										if field == lengthFieldName:
	
											sheet.Cells(excelrow,lengthcol).Value = round(rowoutClipFC[rows.fields.index(lengthFieldName)],2)
											sheet.Cells(excelrow,lengthcol).Font.Size = 8
											sheet.Cells(excelrow,lengthcol).HorizontalAlignment = win32com.client.constants.xlRight
												
										else:
										
											sheet.Cells(excelrow,fieldnum + 2).Value = rowoutClipFC[fieldnum]														 
											sheet.Cells(excelrow,fieldnum + 2).Font.Size = 8
											sheet.Cells(excelrow,fieldnum + 2).HorizontalAlignment = win32com.client.constants.xlRight
											
											if rowoutClipFC[fieldnum] == None:
												pass
											elif type(rowoutClipFC[fieldnum]) in (int, float) or isinstance(rowoutClipFC[fieldnum], datetime.datetime):
												lenrowoutClipFC = len(str(rowoutClipFC[fieldnum]))
											else:
												lenrowoutClipFC = len(rowoutClipFC[fieldnum])
											if lenrowoutClipFC >= 38:
												sheet.Cells(excelrow,fieldnum + 2).WrapText = True
											else:
												pass                                        
																																										 
																																							 
								else:
										
									sheet.Cells(excelrow,fieldnum + 2).Value = rowoutClipFC[fieldnum]																
									sheet.Cells(excelrow,fieldnum + 2).Font.Size = 8
									sheet.Cells(excelrow,fieldnum + 2).HorizontalAlignment = win32com.client.constants.xlRight
									
									if rowoutClipFC[fieldnum] == None:
										pass
									elif type(rowoutClipFC[fieldnum]) in (int, float) or isinstance(rowoutClipFC[fieldnum], datetime.datetime):
										lenrowoutClipFC = len(str(rowoutClipFC[fieldnum]))
									else:
										lenrowoutClipFC = len(rowoutClipFC[fieldnum])
									if lenrowoutClipFC >= 38:
										sheet.Cells(excelrow,fieldnum + 2).WrapText = True
									else:
										pass																	
																																											
								fieldnum += 1															 
	
							excelrow += 1
								
					if 'areaFieldName' in locals():
							del areaFieldName
					if 'lengthFieldName' in locals():
							del lengthFieldName
	
				
				# Delete feature Layer. Need to do this because it will hang on the next loop.
				if arcpy.Exists("dsPath"):
					arcpy.Delete_management("dsPath")
						
				excelrow += 1

	excelrow += 1

# Adjust columns widths
sheet.Columns.AutoFit()
sheet.Rows.AutoFit()

if sheet.Columns(1).ColumnWidth > 40:
	sheet.Columns(1).ColumnWidth = 40
	
if sheet.Columns(2).ColumnWidth > 40:
	sheet.Columns(2).ColumnWidth = 40   

#Set Report Header Title
sheet.Cells(1,1).Value = "INTEREST OVERLAP REPORT"
sheet.Cells(1,1).Font.Size = 16
sheet.Cells(1,1).Font.Bold = True

# Set Border around report
sheet.Range(sheet.Cells(1,1),sheet.Cells(excelrow - 1,maxcolCount + 3)).BorderAround()

# Set the Area of Interest path
sheet.Cells(8,1).Value = "Interest Area = " + str(clip_feature)
sheet.Cells(8,1).Font.Size = 8
sheet.Cells(8,1).Font.Bold = True
excel.Cells(8,1).Font.ColorIndex = 3

if len(str(clip_feature)) > 100:
	excel.Cells(8,1).WrapText = True

# Set second sheet in book and rename it for the parameters
sheet2 = book.Worksheets(2)
sheet2.Name = "Input Information"

sheet2.Cells(1,1).Value = "Input Feature Class for Area of Interest:"
sheet2.Cells(1,1).Font.Bold = True

sheet2.Cells(1,2).Value =  clip_feature

sheet2.Cells(2,1).Value = "SQL Query Used on Input Feature Class:"
sheet2.Cells(2,1).Font.Bold = True

sheet2.Cells(2,2).Value =  sqlQuery

sheet2.Cells(3,1).Value = "Output Folder Location of Excel Spreadsheet:"
sheet2.Cells(3,1).Font.Bold = True

sheet2.Cells(3,2).Value =  output_folder

sheet2.Cells(4,1).Value = "Configuration Excel Spreadsheet Location"
sheet2.Cells(4,1).Font.Bold = True

sheet2.Cells(4,2).Value =  xls

sheet2.Cells(5,1).Value = "Scratch Geodatabase Location"
sheet2.Cells(5,1).Font.Bold = True

sheet2.Cells(5,2).Value = scratchFolder

sheet2.Columns.AutoFit()
sheet2.Rows.AutoFit()

# Save the excel spreadsheet as a .xlsx file. This may be useful to pull in macro calls later on.
book.SaveAs(output_folder + "\\" + "Interest_report_" + inputSourceName + "_" + strftime('%d%b%y') + ".xlsx")

# Quit the instance of excel from win32com
excel.Quit()

# Output end time of process
toc = time.clock()

# Print out time informaton (total seconds/60 for rough amount of minutes
timeLapse = toc - tic

m,s = divmod(timeLapse,60)
h,m = divmod(m,60)

arcpy.AddMessage("Report run in %d hours %02d minutes %02d seconds" % (h, m, s))

# Delete some variables that were causing issues earlier
del excel, book, sheet, excelrow, row
