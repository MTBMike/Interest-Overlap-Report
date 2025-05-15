'''
Tool name: Interest Overlap Report (IOR)
Developer: Mike MacRae for Ministry of Energy, Mines and Low Carbon Innovation
Contact: michael.macrae@gov.bc.ca or mineral.titles@gov.bc.ca
Date: Developed March 2014. Updated April 2023
'''

# Import required modules
import arcpy
import os
import win32com.client
import itertools
import datetime
import urllib
import hmac
import json
import requests
import hashlib
import base64
import time
from arcpy import env
from getpass import getuser
from collections import OrderedDict

## Set the parameters for the ArcGIS GUI
AOI = arcpy.GetParameterAsText(0)
sqlQuery = arcpy.GetParameterAsText(1)
shFieldList = arcpy.GetParameter(2)
pre_defined_layer_list_choice = arcpy.GetParameterAsText(3)
layerList = [x.strip("'") for x in arcpy.GetParameterAsText(4).split(";")]
output_GDB = arcpy.GetParameterAsText(5)
output_excel = arcpy.GetParameterAsText(6)
output_name = arcpy.GetParameterAsText(7)
username = arcpy.GetParameterAsText(8)
mtoprodpassword = arcpy.GetParameter(9)
bcgwpassword = arcpy.GetParameter(10)
#createGeomark = arcpy.GetParameter(11)


def login(username, mtoprodpassword, bcgwpassword):
    ''' 
    A login prompt to get the users username and password for both MTOPROD and BCGW,
    create database connections for each and log into the databases
    '''
    arcpy.AddMessage("Checking to see if directory exists")
    if os.path.exists(os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser())):
        arcpy.AddMessage("Directory exists")
        pass        

    else:
        arcpy.AddMessage("Directory didn't exist")
        os.makedirs(os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser()))        

    arcpy.AddMessage("Passed directory check")
    
    
    try:
        arcpy.AddMessage("Logging into MTOPROD...")
        arcpy.CreateDatabaseConnection_management(os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser()),
                                                  "MTOPROD.sde",
                                                  "ORACLE",
                                                  "nrkdb02.bcgov/mtoprod.nrs.bcgov",
                                                  "DATABASE_AUTH",
                                                  username,
                                                  mtoprodpassword,
                                                  "DO_NOT_SAVE_USERNAME")       
        
        arcpy.AddMessage("Logging into BCGW...")
        arcpy.CreateDatabaseConnection_management(os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser()),
                                                  "BCGW.sde",
                                                  "ORACLE",
                                                  "bcgw.bcgov/idwprod1.bcgov",
                                                  "DATABASE_AUTH",
                                                  username,
                                                  bcgwpassword,
                                                  "DO_NOT_SAVE_USERNAME")
    except arcpy.ExecuteError:
        print(arcpy.GetMessages())
        arcpy.AddMessage("Entered Exception")
        if os.path.exists(os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser())):    
            arcpy.Delete_management(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files" + '\\' + getuser())
        else:
            pass
        
    arcpy.env.overwriteOutput = True
    
    env.workspace = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "MTOPROD.sde")
    env.workspace = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "BCGW.sde")
        
    
def logout():
    '''
    A logout prompt to delete the database connections created in the 'login' function
    '''
    
    arcpy.AddMessage("Logging out of MTOPROD and BCGW...")
    if os.path.exists(os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser())):    
        arcpy.Delete_management(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files" + '\\' + getuser())
    else:
        pass

def delFeatLayer(featLayer):
    '''
    A function to test and delete feature layers
    '''
    if arcpy.Exists(featLayer):
        arcpy.Delete_management(featLayer)
    else:
        pass
    

def createScratchGDB(output_folder):
    '''
    A function that creates a scratch geodatabase
    '''
    
    arcpy.AddMessage("Setting Output Geodatabase Location...")
    
    arcpy.env.overwriteOutput = True
    
    if output_folder == '':
    
        output_folder = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser())
        lockedGDBs=[]
        
        env.workspace = output_folder
        
        for gdb in arcpy.ListWorkspaces("*", "FileGDB"):

            basename = os.path.basename(gdb)
            inFeatures, file_extension = os.path.splitext(os.path.basename(gdb))

            try:
                delFeatLayer(gdb)
            except: 
                lockedGDBs.append(basename)
        
        gdbIndex = 1
        scratchGDB = "scratch_1.gdb"

        while gdbIndex:

            if scratchGDB not in lockedGDBs:
                arcpy.CreateFileGDB_management(output_folder, scratchGDB)
                print("Created " + scratchGDB)
                break
            else:
                gdbIndex+=1
                scratchGDB = "scratch_" + str(gdbIndex) + ".gdb"
        
    else:
                
        scratchGDB = "IOR_Clipped_FeatureClasses_" + output_name + "_" + time.strftime('%d%b%Y') + ".gdb"
        scratchLoc = output_folder + "\\" + scratchGDB
        
        delFeatLayer(scratchLoc)
        
        arcpy.CreateFileGDB_management(output_folder, scratchGDB)       

    return output_folder, scratchGDB


def getXLSData(excel):
    '''
    A function to set which Database to pull Mineral Titles data from and
    to create a python dictionary to store the dataset name and buffer distance for
    layers in the MASTER spreadsheet that ask for buffering
    '''
    arcpy.AddMessage("Getting XLS Data...")
    xls = r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\Testing_Area\mmacrae\IOR\2022Nov_updates\Excel_Spreadsheets\InterestReport_Layer_List_MASTER.xlsx\IOR_Data$"

    xlsFields = [field.name for field in arcpy.ListFields(xls)]
    
    return xls, xlsFields


def processAOI(AOI):
    '''
    A function to determine the number of features in the AOI, set an exception if it's not only one feature
    and to compute the area of the AOI from its spatial file
    '''

    arcpy.AddMessage("Processing AOI...")
    delFeatLayer("AOIFeat")
    
    if sqlQuery:
        
        AOIFeats = arcpy.GetCount_management(arcpy.MakeFeatureLayer_management(AOI,"AOIFeat",sqlQuery))
        AOICount = int(AOIFeats.getOutput(0))
        
        arcpy.AddMessage("Count of AOI features = " + str(AOICount))
        
        if AOICount != 1:
            if AOICount == 0:
                raise Exception("SQL Query returns empty result. Please redefine query, validate to see one record returns and rerun IOR.")
            elif AOICount > 1:
                raise Exception("SQL Query returns more than one feature. Please redefine query, validate to see one record returns and rerun IOR.")
        else:
            arcpy.FeatureClassToFeatureClass_conversion("AOIFeat", output_folder + "\\" + scratchGDB, "AOI")
            processedAOI = output_folder + "\\" + scratchGDB + "\\" + "AOI"
            
    else:
        arcpy.MakeFeatureLayer_management(AOI,"AOIFeat")
        arcpy.FeatureClassToFeatureClass_conversion("AOIFeat", output_folder + "\\" + scratchGDB, "AOI")
        processedAOI = output_folder + "\\" + scratchGDB + "\\" + "AOI"

    AOI_Area = 0
    for row in arcpy.da.SearchCursor(processedAOI, ["SHAPE@AREA"]):
        AOI_Area += row[0]
        
    processedAOI_Hectares = round(AOI_Area/10000, 2)
        
    out_coordinate_system = arcpy.SpatialReference(3857)
    projected_AOI = output_folder + "\\" + scratchGDB + "\\" + "AOI_projected"
    arcpy.Project_management(processedAOI, projected_AOI, out_coordinate_system)
    
    for row in arcpy.da.SearchCursor(projected_AOI, ["SHAPE@TRUECENTROID"]):
        urlX, urlY = row[0]
        coords = str(urlX) + ',' + str(urlY)
        
    mxd = arcpy.mapping.MapDocument(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\templates\getScale_template.mxd")
    df = arcpy.mapping.ListDataFrames(mxd)[0]
    
    arcpy.MakeFeatureLayer_management(projected_AOI, 'projected_AOI')
    layer = arcpy.mapping.Layer('projected_AOI')
    arcpy.mapping.AddLayer(df, layer)

    df.scale = int(math.ceil(df.scale/500)*500)

    iMapBCBaseURL = 'https://arcmaps.gov.bc.ca/ess/hm/imap4m/?scale={0}&center={1}'.format(df.scale, coords)
    
    return processedAOI, processedAOI_Hectares, iMapBCBaseURL


def initializeSpreadsheet():
    '''
    A function to create an empty spreadsheet, add 4 required sheets and name them
    '''
    arcpy.AddMessage("Initializing spreadsheet...")
    
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = True
    
    # Initialize a workbook within excel
    book = excel.Workbooks.Add()
    
    book.Sheets.Add()
    
    # Set first sheet in book and rename it for the report
    book.Worksheets(1).Name = 'Summary'
    book.Worksheets(2).Name = 'Interest_Report'
    book.Worksheets(3).Name = 'Districts_and_BCGS-NTS_Location'
    book.Worksheets(4).Name = 'Input_Information'

    return book, excel


def sheetCells(sheet, excelrow, excelcol, value="", size=10, bold=False, italic=False, underline=False, fontcolor=0, fillcolor=0, wrap=False, numFormat=None):
    '''
    A function to format cells in each sheet
    '''

    if isinstance(value, datetime.datetime) == True and value < datetime.datetime.strptime('01 01 1900', '%m %d %Y'):
        sheet.Cells(excelrow, excelcol).Value = str(value)
    else:
        sheet.Cells(excelrow, excelcol).Value = value
            
    sheet.Cells(excelrow, excelcol).Font.Size = size
    sheet.Cells(excelrow, excelcol).Font.Bold = bold
    sheet.Cells(excelrow, excelcol).Font.Italic = italic
    sheet.Cells(excelrow, excelcol).Font.Underline = underline
    sheet.Cells(excelrow, excelcol).Font.ColorIndex = fontcolor
    sheet.Cells(excelrow, excelcol).Interior.ColorIndex = fillcolor
    sheet.Cells(excelrow, excelcol).WrapText = wrap
#     sheet.Cells(excelrow, excelcol).Style = style
    sheet.Cells(excelrow, excelcol).NumberFormat = numFormat


def mappingFields(lyr, fieldList):
    ''' 
    A function to map the fields of the output (clipped) feature class. This will limit the output 
    fields to the fields chosen in the configuration spreadsheet.
    '''    
    
    fms = arcpy.FieldMappings()         
    for field in arcpy.ListFields(lyr):
        if field.name in fieldList:
            fm = arcpy.FieldMap()
            fm.addInputField(lyr, field.name)
            fms.addFieldMap(fm)
                
    return fms


def processData(processedAOI, processedAOI_Hectares, xls, xlsFields):
    ''' 
    A function to process layers to determine if there is an overlap and subsequently clips and overlaps.
    The process data is used further on in the script to report on a spreadsheet.
    '''    

    arcpy.AddMessage("Processing Layers...")

    layerListDict = {}
    collectFeatsCountDict = OrderedDict()
    originalprocessAOI = processedAOI
    
    catList = list(sorted(set([row[1] for row in arcpy.da.SearchCursor(xls, xlsFields) if row[1] not in [u'District', u'Location']])))
    mineral_coal = u'Mineral/Coal'
    catList.insert(0, catList.pop(catList.index(mineral_coal)))
    
    for row in arcpy.da.SearchCursor(xls, xlsFields):

        if row[2] in layerList:        
        
            fcName = row[2].replace(' ', '_')
            if row[4] == 'BCGW':
                fc = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "BCGW.sde", row[5])
            elif row[4] == 'MTOPROD':
                fc = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "MTOPROD.sde", row[5])
            else:
                fc = os.path.join(row[4], row[5])
                
            arcpy.AddMessage("Processing Layer: " + row[2])

            if row[11] is not None:
                arcpy.AddMessage("    " + "Buffering AOI for " + "'" + str(row[2]) + "' layer by " + str(int(float(row[11]))))
                processedAOI = originalprocessAOI
                outBuffer = os.path.join(os.path.dirname(processedAOI), str(row[0]) + "_" + str(row[11]).replace('.','_') + "m_buffer")
                arcpy.Buffer_analysis(processedAOI, outBuffer, row[11])
                processedAOI = outBuffer
            else:
                processedAOI = originalprocessAOI
            
            delFeatLayer("lyr")
            arcpy.MakeFeatureLayer_management(fc, "lyr")
            
            # Test for table joins
            if row[8] is not None:
                arcpy.AddJoin_management("lyr", str(row[9]), str(row[8]), str(row[10]))
                arcpy.AddMessage("    " + row[5] + " joined with " + os.path.basename(row[8]))
            else:
                pass
            
            lyr = arcpy.mapping.Layer("lyr")
            
            # Test to see if a Definition Query is needed
            if row[7] is not None:
                if lyr.supports("DEFINITIONQUERY"):
                    lyr.definitionQuery = row[7]
                    arcpy.AddMessage("    Def. Query applied")
                else:
                    arcpy.AddMessage("    Does not Support Definition Queries")
            else:
                arcpy.AddMessage("    No Definition Query")
                pass
            
            fieldList = [field.name for field in arcpy.ListFields(lyr) if field.name in [str(row[i]) for i in range (12, 24) if row[i] is not None]]
            
            fms = mappingFields(lyr, fieldList)
        
            selectresult = arcpy.GetCount_management(lyr)
             
            selectcount = int(selectresult.getOutput(0))
            arcpy.AddMessage("    Count before select: " + str(selectcount))
            
            arcpy.AddMessage("    Processing Select by Location")
            arcpy.SelectLayerByLocation_management(lyr, "intersect", processedAOI)

            selectresult = arcpy.GetCount_management(lyr)
            selectcount = int(selectresult.getOutput(0))       
 
            arcpy.AddMessage("    Count after selection: " + str(selectcount))

            # Test to see if records were selected during select by location
            if selectcount != 0:

                arcpy.AddMessage("    Exporting Selected Features")
                
                arcpy.FeatureClassToFeatureClass_conversion(lyr, output_folder + "\\" + scratchGDB, row[0], '', field_mapping=fms)
                
                # Describe the shapetype of each layer
                desc = arcpy.Describe(lyr)
                if desc.shapeType == "Polygon":

                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0], "ORIGINAL_HECTARES", "FLOAT", "", "", "", "Original Area (Ha)")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0], "ORIGINAL_HECTARES", "round(!shape.area!/10000,6)", "PYTHON_9.3") 
                    
                    arcpy.AddMessage("    Clipping Select Features with AOI")
                    arcpy.Clip_analysis(output_folder + "\\" + scratchGDB + "\\" + row[0], processedAOI, output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip")
                    
                    arcpy.Delete_management(output_folder + "\\" + scratchGDB + "\\" + row[0])
                    
                    if row[1] in layerListDict:
                        layerListDict[row[1]].update({row[2]: [str(row[0] + "_clip"), str(row[26])]})
                    else:
                        layerListDict[row[1]] = {}
                        layerListDict[row[1]].update({row[2]: [str(row[0] + "_clip"), str(row[26])]})
                    
                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "OVERLAPPING_HECTARES", "FLOAT", "", "", "", "Overlapping Area (Ha)")
                                                       
                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "OVERLAPPING_HECTARES", "round(!shape.area!/10000,6)", "PYTHON_9.3")

                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "PERCENT_OF_LAYER_BEING_OVERLAPPED_BY_AOI", "FLOAT", "", "", "", "% Layer being Overlapped by AOI")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "PERCENT_OF_LAYER_BEING_OVERLAPPED_BY_AOI", "round(!OVERLAPPING_HECTARES!/!ORIGINAL_HECTARES!*100,6)", "PYTHON_9.3")

                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "PERCENT_OF_AOI_BEING_OVERLAPPED_BY_LAYER", "FLOAT", "", "", "", "% AOI being Overlapped by Layer")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "PERCENT_OF_AOI_BEING_OVERLAPPED_BY_LAYER", "round(!OVERLAPPING_HECTARES!/{0}*100, 12)".format(processedAOI_Hectares), "PYTHON_9.3")

                                                   
                elif desc.shapeType == "Polyline":
                    
                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0], "ORIGINAL_LENGTH", "FLOAT", "", "", "", "Original Length")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0], "ORIGINAL_LENGTH", "round(!shape.length!/1000,6)", "PYTHON_9.3") 

                    arcpy.AddMessage("    Clipping Select Features with AOI")
                    arcpy.Clip_analysis(output_folder + "\\" + scratchGDB + "\\" + row[0], processedAOI, output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip")
                    
                    arcpy.Delete_management(output_folder + "\\" + scratchGDB + "\\" + row[0])

                    if row[1] in layerListDict:
                        layerListDict[row[1]].update({row[2]: [str(row[0] + "_clip"), str(row[26])]})
                    else:
                        layerListDict[row[1]] = {}
                        layerListDict[row[1]].update({row[2]: [str(row[0] + "_clip"), str(row[26])]})                    
                          
                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "OVERLAPPING_LENGTH", "FLOAT", "", "", "", "Overlapping Length")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "OVERLAPPING_LENGTH", "round(!shape.length!/1000,6)", "PYTHON_9.3") 

                elif desc.shapeType == "Point":
                    
                    arcpy.AddMessage("    Clipping Select Features with AOI")
                    arcpy.Clip_analysis(output_folder + "\\" + scratchGDB + "\\" + row[0], processedAOI, output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip")
                    
                    arcpy.Delete_management(output_folder + "\\" + scratchGDB + "\\" + row[0])

                    if row[1] in layerListDict:
                        layerListDict[row[1]].update({row[2]: [str(row[0] + "_clip"), str(row[26])]})
                    else:
                        layerListDict[row[1]] = {}
                        layerListDict[row[1]].update({row[2]: [str(row[0] + "_clip"), str(row[26])]})

                    arcpy.management.AddXY(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip")

                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "EASTING", "DOUBLE", "", 2)
                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "NORTHING", "DOUBLE", "", 2)

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "EASTING", "[POINT_X]")
                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "NORTHING", "[POINT_Y]")

                    arcpy.DeleteField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", ["POINT_X", "POINT_Y"])
                    
                elif desc.shapeType == "Multipoint":
                    
                    arcpy.AddMessage("    Clipping selected features with AOI")
                    arcpy.Clip_analysis(output_folder + "\\" + scratchGDB + "\\" + row[0], processedAOI, output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip")

                    arcpy.Delete_management(output_folder + "\\" + scratchGDB + "\\" + row[0])

                    if row[1] in layerListDict:
                        layerListDict[row[1]].update({row[2]: [str(row[0] + "_clip"), str(row[26])]})
                    else:
                        layerListDict[row[1]] = {}
                        layerListDict[row[1]].update({row[2]: [str(row[0] + "_clip"), str(row[26])]})

                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "POINT_LOCATION", "TEXT", "", "", 20, "Point Location")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "POINT_LOCATION", '"Point Location"', "PYTHON_9.3")
                  
            else:

                if row[1] in layerListDict:
                    layerListDict[row[1]].update({row[2]: ["No Overlap Found", str(row[26])]})
                else:
                    layerListDict[row[1]] = {}
                    layerListDict[row[1]].update({row[2]: ["No Overlap Found", str(row[26])]})

                arcpy.AddMessage("    No Overlap Found")
       
            if arcpy.Exists(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip"):              
                newDict = {}
                newDict[row[2]] = [int(arcpy.GetCount_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip").getOutput(0)), row[27]]
                
                if row[1] not in collectFeatsCountDict.keys():
                    collectFeatsCountDict[row[1]] = {}
                    collectFeatsCountDict[row[1]].update(newDict)
                else:
                    collectFeatsCountDict[row[1]].update(newDict)
      

    tenureDict = {}
    reserveDict = {}
    remainingMiningDict = {}
    remaningLayers = {}
    
    for category, values in sorted(collectFeatsCountDict.iteritems()):
        if category == 'Mineral/Coal':
            for fClass, fItems in values.items():
                if 'Tenure - ' in fClass:
                    tenureDict[fClass] = fItems
                elif 'Reserve - ' in fClass:
                    reserveDict[fClass] = fItems
                else:
                    remainingMiningDict[fClass] = fItems

    featsCountDict = OrderedDict(itertools.chain(sorted(tenureDict.items()), sorted(reserveDict.items()), sorted(remainingMiningDict.items())))
    
    for key, value in collectFeatsCountDict.iteritems():
        if key == 'Mineral/Coal':
            collectFeatsCountDict[key] = featsCountDict
        else:
            collectFeatsCountDict[key] = OrderedDict(sorted(value.items()))   
    
    arcpy.AddMessage('')
    arcpy.AddMessage('===============================================================================')

    return layerListDict, collectFeatsCountDict


def createInterestReportSheet(book, layerListDict):
    '''
    A function to process data and populate an interest report sheet that provides details of each feature
    in each overlapping layer
    '''    

    arcpy.AddMessage("Creating Interest Report Detail sheet...")
    
    sheet = book.Worksheets("Interest_Report")
    
    excelrow = 1
    excelcol = 1
    
    sheetCells(sheet, excelrow, excelcol, "INTEREST OVERLAP REPORT", 16, True)
    
    excelrow += 2
    
    sheetCells(sheet, excelrow, excelcol, "Legend", 12, True, False, True)
    
    excelrow += 1
    
    sheetCells(sheet, excelrow, excelcol, "", None, False, False, False, None, 15)
    sheetCells(sheet, excelrow, excelcol + 1, "Layer Category", 10, True, False, False, None, None)
    
    excelrow += 1
    
    sheetCells(sheet, excelrow, excelcol, "", None, False, False, False, None, 37)
    sheetCells(sheet, excelrow, excelcol + 1, "Overlapping Layer", 10, True)
    
    sheet.Range(sheet.Cells(excelrow - 1, excelcol),sheet.Cells(excelrow,excelcol + 1)).BorderAround() 
    
    excelrow += 2
    
    legendrow = excelrow
    
    excelrow += len(layerListDict.keys())
    
    excelrow += 2
    
    sheetCells(sheet, excelrow, excelcol, "INTEREST OVERLAPS (click to view Data BC Record)", 12, True, True)
    
    excelrow += 2
    legendsearchrow = excelrow

    mineralcoal = {}
    tenureDict = {}
    reserveDict = {}
    remainingMiningDict = {}
    otherLayersDict = {}
    
    crossReferenceDict = {}
    IRCatDict = {}
    
    for category, values in layerListDict.items():

        if category == 'Mineral/Coal':
            for Fclass, listItems in sorted(values.items()):
                if 'Tenure - ' in Fclass:
                    tenureDict[Fclass] = listItems
                elif 'Reserve - ' in Fclass:
                    reserveDict[Fclass] = listItems
                else:
                    remainingMiningDict[Fclass] = listItems      
            mineralcoal[category] =  itertools.chain(sorted(tenureDict.items()), sorted(reserveDict.items()), sorted(remainingMiningDict.items()))
    
        else:       
            otherLayers = {}
            for Fclass, listItems in sorted(values.items()):  
                otherLayers[Fclass] = listItems
            otherLayersDict[category] = itertools.chain(sorted(otherLayers.items()))
         
    processedlayerListDict = itertools.chain(mineralcoal.items(), sorted(otherLayersDict.items()))
    
    env.workspace = output_folder + "\\" + scratchGDB
        
    cat = []
    for category, values in processedlayerListDict:
        
        cat.append(category)
        
        sheetCells(sheet, excelrow, excelcol, category, 12, True, False, False, None, 15)

        excelrow +=2
        
        for Fclass, listItems in values:
            
            if listItems[1] != 'None':
                sheetCells(sheet, excelrow, excelcol, '=HYPERLINK("{0}","{1}")'.format(listItems[1], str(Fclass)), 10, False, False, True, None)
                crossReferenceDict[Fclass] = excelrow
            else:
                sheetCells(sheet, excelrow, excelcol, Fclass, 10, True, False, False, None)
                crossReferenceDict[Fclass] = excelrow
            
            if listItems[0] != 'No Overlap Found':
                for fc in arcpy.ListFeatureClasses():
                    if fc == listItems[0]:
                        fields = [field for field in arcpy.ListFields(fc) if not field.required]
                        excelcol = 2
                        
                        for field in fields:
                            sheetCells(sheet, excelrow, excelcol, field.aliasName, 8, True, False, False, None, 37)
                            excelcol += 1
                        excelrow += 1
                        excelcol = 2
                        
                        for row in arcpy.da.SearchCursor(fc, [field.name for field in fields]):
                            for item in row:                                
                                if isinstance(item, int):
                                    sheetCells(sheet, excelrow, excelcol, int(item), 8, False, False, False, None, numFormat=0)
                                elif isinstance(item, float):
                                    sheetCells(sheet, excelrow, excelcol, float(item), 8, False, False, False, None)
                                elif isinstance(item, (str, unicode)):
                                    
                                    try:
                                        sheetCells(sheet, excelrow, excelcol, float(item), 8, False, False, False, None)
                                    except:
                                        pass                                     
                                    
                                    if item.isdigit():
                                        sheetCells(sheet, excelrow, excelcol, int(item), 8, False, False, False, None, numFormat=0)
                                    else:
                                        sheetCells(sheet, excelrow, excelcol, item, 8, False, False, False, None)

                                excelcol += 1
                            excelcol = 2
                            excelrow += 1

            else:
                if listItems[1] != 'None':
                    sheetCells(sheet, excelrow, excelcol, '=HYPERLINK("{0}","{1}")'.format(listItems[1], str(Fclass)), 10, False, False, True, None)
                else:
                    sheetCells(sheet, excelrow, excelcol, Fclass, 10, True, False, False, None)
                    
                sheetCells(sheet, excelrow, excelcol + 1, "No Overlaps", 10, True, False, False)
                
            excelcol = 1
            excelrow += 2
        excelrow += 1
    
    excelrow -= 3

    for category in cat:
        sheetCells(sheet, legendrow, excelcol, '=HYPERLINK("#"&CELL("address",INDEX(A{0}:A{1},MATCH("{2}",A{0}:A{1},0),1)), "{3}")'.format(legendsearchrow, excelrow, category, "Go to " + category + " category"), 10, False, False, True, None)
        legendrow += 1

    sheet.Columns.AutoFit()
    
    return crossReferenceDict


def createSummarySheet(book, excel, processedAOI, processedAOI_Hectares, collectFeatsCountDict, iMapBCBaseURL, crossReferenceDict, geoMark_URL):
    ''' 
    A function to process data and create a count summary of the mining layer overlaps
    '''
    
    arcpy.AddMessage("Creating summary sheet...")
    
    sheet = book.Worksheets('Summary')
    
    excel.ActiveWindow.DisplayGridlines = False
    
    excelrow = 1
    excelcol = 1
    
    #Set Report Header Title   
    sheetCells(sheet, excelrow, excelcol, "INTEREST OVERLAP REPORT - SUMMARY", 16, True)
    
    # Set the "REPORT FOR INTERNAL USE ONLY" comment in the sheet and do some formatting    
    sheetCells(sheet, excelrow + 1, excelcol, "REPORT FOR INTERNAL USE ONLY", 10, True, False, False, 3)
    
    # Set a comment to indicate when the report was run and do some formatting    
    sheetCells(sheet, excelrow + 2, excelcol, "Report run on " + time.strftime('%d%b%y') + " @ " + time.strftime('%I:%M:%S'), 10)   
        
    # Add BC Government logo
    pic = sheet.Pictures().Insert(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Image\BC_EMLC_H_RGB_pos.jpg")
    
    # Format size of the logo    
    pic.Height = 65
    pic.Width = 130
    
    # Format location of the logo
    cell = sheet.Cells(excelrow, excelcol + 1)
    
    pic.Left = cell.Left
    pic.Top = cell.Top
       
    excelrow = 1
                    # Add logo to report and set size and location
    pic = sheet.Pictures().Insert(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Image\Overlap_Example.JPG")
    
    # Format size of the logo
    pic.Height = 390
    pic.Width = 190
    
    # Format location of the logo
    cell = sheet.Cells(excelrow,3)
    
    pic.Left = cell.Left + 40
    pic.Top = cell.Top
    
    excelrow = 5
    
    if os.path.basename(AOI) not in [r'MTA_SPATIAL.MTA_ACQUIRED_TENURE_SVW', r'MTA_SPATIAL.MTA_SITE_SVW', r'WHSE_MINERAL_TENURE.MTA_SITE_SP',r'WHSE_MINERAL_TENURE.MTA_ACQUIRED_TENURE_SVW']:
        
        sheetCells(sheet, excelrow, excelcol, "AREA OF INTEREST INFORMATION", 10, bold=True, underline=True)
        
        if geoMark_URL <> '':
            sheetCells(sheet, excelrow, excelcol + 1, '=HYPERLINK("{0}","{1}")'.format(geoMark_URL, 'Area of Interest GeoMark Link'), 10, False, False, True, None)

        staticexcelrow = excelrow
        
        # Set a conditional sentence to determine if the user chose any fields to summarize in the area of interest parameter
        # In the first condition, determine if there are no fields to summarize
        if len(shFieldList) == 0:
            
            # Increment the counter by 8. This will push the rows paste the title header and the image
            # Also, set another variable to be used as a static row count which will be used to set the position of the legend
            # so that the legend lies next to the header information
            
            excelrow += 1
            
            sheetCells(sheet, excelrow, excelcol, "Area (ha)", 10, False)
                
            sheetCells(sheet, excelrow, excelcol + 1, processedAOI_Hectares)

            # Do some formatting on the area of interest header information. Set the font size, border, fill color and alignment
            sheet.Range(sheet.Cells(staticexcelrow, excelcol),sheet.Cells(excelrow, excelcol + 1)).Font.Size = 10
            sheet.Range(sheet.Cells(staticexcelrow, excelcol),sheet.Cells(excelrow, excelcol + 1)).BorderAround()
            sheet.Range(sheet.Cells(staticexcelrow, excelcol),sheet.Cells(excelrow, excelcol + 1)).Interior.ColorIndex = 36
            sheet.Range(sheet.Cells(staticexcelrow, excelcol + 1),sheet.Cells(excelrow, excelcol + 1)).HorizontalAlignment = win32com.client.constants.xlRight                            
            
            # Increment the row count by 2 for the position of the first overlapping layer
            excelrow += 1
            
        # End of conditional sentence. If the user chose some fields from the area of interest, update the header
        else:
    
            # Loop over the field list from the fields chosen by the user
            for shField in shFieldList:
                
                excelrow += 1
    
                # Add each field name in the first column under the image and do some formatting               
                sheetCells(sheet, excelrow, excelcol, shField.replace("_"," "), 10, False, True)
                
                # Loop through the area of interest file and add the value that corresponds to the field name
    
                for shrow in arcpy.da.SearchCursor(AOI, [shField], sqlQuery):
                    sheetCells(sheet, excelrow, excelcol + 1, str(shrow[0]), 10)
            
            excelrow += 1       
            sheetCells(sheet, excelrow, excelcol, "Area (ha)", 10, False)
                
            sheetCells(sheet, excelrow, excelcol + 1, processedAOI_Hectares)            

            # Do some formatting on the area of interest header information. Set the font size, border, fill color and alignment
            sheet.Range(sheet.Cells(staticexcelrow, excelcol),sheet.Cells(excelrow, excelcol + 1)).Font.Size = 10
            sheet.Range(sheet.Cells(staticexcelrow, excelcol),sheet.Cells(excelrow, excelcol + 1)).BorderAround()
            sheet.Range(sheet.Cells(staticexcelrow, excelcol),sheet.Cells(excelrow, excelcol + 1)).Interior.ColorIndex = 36
            sheet.Range(sheet.Cells(staticexcelrow, excelcol + 1),sheet.Cells(excelrow, excelcol + 1)).HorizontalAlignment = win32com.client.constants.xlRight                            
                    
            # Increment the row count by 2 for the position of the first overlapping layer
            excelrow += 1

        
    # next conditional sentence is to determine of the user chose the tenure feature class
    elif os.path.basename(AOI) in ('MTA_SPATIAL.MTA_ACQUIRED_TENURE_SVW', 'WHSE_MINERAL_TENURE.MTA_ACQUIRED_TENURE_SVW'):
        
        sheetCells(sheet, excelrow, excelcol, "Tenure Information", 10, True)
        
        if geoMark_URL <> '':
            sheetCells(sheet, excelrow, excelcol + 1, '=HYPERLINK("{0}","{1}")'.format(geoMark_URL, 'Area of Interest GeoMark Link'), 10, False, False, True, None)        
        
        excelrow += 1
        staticexcelrow = excelrow
        
        fieldItemDict = {}
        with arcpy.da.SearchCursor(AOI, ["TENURE_NUMBER_ID", "TITLE_TYPE_DESCRIPTION", "ISSUE_DATE", "GOOD_TO_DATE", "OWNER_NAME", "AREA_IN_HECTARES", "SHAPE@AREA"], sqlQuery) as cursor:
            for row in cursor:
                fieldCount = 0

                for field in cursor.fields:
                    if field.title() == "Shape@Area":
                        fieldItemDict[field] = row[fieldCount]
                        sheetCells(sheet, excelrow, excelcol, "True Geometry Size (ha)", 10, False)
                        sheetCells(sheet, excelrow, excelcol + 1, str(round(row[fieldCount]/10000, 2)))
                    else:
                        fieldItemDict[field] = row[fieldCount]
                        sheetCells(sheet, excelrow, excelcol, str(field.title()), 10, False)
                        sheetCells(sheet, excelrow, excelcol + 1, str(row[fieldCount]))

                    fieldCount += 1
                    excelrow += 1   
            
            # Do some formatting on the area of interest header information. Set the font size, border, fill color and alignment
            sheet.Range(sheet.Cells(staticexcelrow, excelcol),sheet.Cells(excelrow - 1, excelcol + 1)).Font.Size = 10
            sheet.Range(sheet.Cells(staticexcelrow, excelcol),sheet.Cells(excelrow - 1, excelcol + 1)).BorderAround()
            sheet.Range(sheet.Cells(staticexcelrow, excelcol),sheet.Cells(excelrow - 1, excelcol + 1)).Interior.ColorIndex = 36
            sheet.Range(sheet.Cells(staticexcelrow, excelcol + 1),sheet.Cells(excelrow - 1, excelcol + 1)).HorizontalAlignment = win32com.client.constants.xlRight      
            
            # Increase the row count by 2 to buffer between the area of interest header information and the next header
            excelrow += 1


    # Last conditional sentence is to determine of the user chose the reserves feature class
    elif os.path.basename(AOI) in ('MTA_SPATIAL.MTA_SITE_SVW', 'WHSE_MINERAL_TENURE.MTA_SITE_SP'):
        
        sheetCells(sheet, excelrow, excelcol, "Reserve Information", 10, True)
        
        if geoMark_URL <> '':
            sheetCells(sheet, excelrow, excelcol + 1, '=HYPERLINK("{0}","{1}")'.format(geoMark_URL, 'Area of Interest GeoMark Link'), 10, False, False, True, None)        
        
        excelrow += 1
        staticexcelrow = excelrow
        
        fieldItemDict = {}
        with arcpy.da.SearchCursor(AOI, ["SITE_NUMBER_ID", "RESERVE_TYPE", "MTA_SITE_ORDER_RESTR_DESC", "TOTAL_AREA", "SHAPE@AREA"], sqlQuery) as cursor:
            for row in cursor:

                fieldCount = 0

                for field in cursor.fields:
                    if field.title() == "Shape@Area":
                        fieldItemDict[field] = row[fieldCount]
                        sheetCells(sheet, excelrow, excelcol, "True Geometry Area (ha)", 10, False)
                        sheetCells(sheet, excelrow, excelcol + 1, str(round(row[fieldCount]/10000, 2)))
                    else:
                        fieldItemDict[field] = row[fieldCount]
                        sheetCells(sheet, excelrow, excelcol, str(field.title()), 10, False)
                        sheetCells(sheet, excelrow, excelcol + 1, str(row[fieldCount]))

                    fieldCount += 1
                    excelrow += 1         

            # Do some formatting on the area of interest header information. Set the font size, border, fill color and alignment
            sheet.Range(sheet.Cells(staticexcelrow, excelcol),sheet.Cells(excelrow - 1, excelcol + 1)).Font.Size = 10
            sheet.Range(sheet.Cells(staticexcelrow, excelcol),sheet.Cells(excelrow - 1, excelcol + 1)).BorderAround()
            sheet.Range(sheet.Cells(staticexcelrow, excelcol),sheet.Cells(excelrow - 1, excelcol + 1)).Interior.ColorIndex = 36
            sheet.Range(sheet.Cells(staticexcelrow, excelcol + 1),sheet.Cells(excelrow - 1, excelcol + 1)).HorizontalAlignment = win32com.client.constants.xlRight      
            
            # Increase the row count by 2 to buffer between the area of interest header information and the next header
            excelrow += 1

    
    excelrow += 1
    
    sheetCells(sheet, excelrow, excelcol, '="Counts of Interest Overlaps" & CHAR(10) & "(See Interest Report Sheet for Details)"', 12, True, wrap=True)

    excelrow += 2
    
    sheetCells(sheet, excelrow, excelcol, "Category", 11, True, fillcolor=15)
    sheetCells(sheet, excelrow, excelcol + 1, "Layer Name  (click to view Interest Report details)", 11, True, fillcolor=15)
    sheetCells(sheet, excelrow, excelcol + 2, "Count of Overlapping Features", 11, True, fillcolor=15)
    sheetCells(sheet, excelrow, excelcol + 3, "iMapBC Link", 11, True, fillcolor=15)
    sheetCells(sheet, excelrow, excelcol + 4, "Reviewer Comments", 11, True, fillcolor=15)
    
    excelrow += 2

        
    for category, catList in collectFeatsCountDict.iteritems():
        sheetCells(sheet, excelrow, excelcol, category, 10, True)
        excelrow += 1
        for Fclass, FclassValues in catList.iteritems():
            if FclassValues[1] is not None:
                sheetCells(sheet, excelrow, excelcol + 1, '=HYPERLINK(CELL("address",Interest_Report!A{0}),"{1}")'.format(crossReferenceDict[Fclass], [k for k in crossReferenceDict.keys() if k == Fclass][0]), 10, False, False, True, None)
                sheetCells(sheet, excelrow, excelcol + 3, '=HYPERLINK("{0}","{1}")'.format(iMapBCBaseURL + '&catalogLayers=' + FclassValues[1], 'View in iMapBC'), 10, False, False, True, None)
            else:
                sheetCells(sheet, excelrow, excelcol + 1, str(Fclass), 10, False, False, False, None)
            sheetCells(sheet, excelrow, excelcol + 2, FclassValues[0], 10, True)
            excelrow += 1

        excelrow +=1    
    
    sheet.Columns.AutoFit()
    sheet.Rows.AutoFit()


def createDistrictSheet(book, xls, xlsFields, processedAOI):
    ''' 
    A function to analyze various district types and maps sheets that overlaps the AOI
    A separate sheet is created and populated with the overlaps.
    '''    
     
    arcpy.AddMessage("Creating district information...")    
         
    sheet = book.Worksheets("Districts_and_BCGS-NTS_Location")
     
    excelrow = 1
    excelcol = 1
    excelrowCount = 0
     
    sheetCells(sheet, excelrow, excelcol, "District Information", 13, True, False, True)
     
    excelrow += 1
     
    for row in arcpy.da.SearchCursor(xls, xlsFields):
        if row[1] == 'District':
                 
            sheetCells(sheet, excelrow, excelcol, '=HYPERLINK("{0}","{1}")'.format(str(row[26]), str(row[2])), 10, False, False, True, None)
              
            excelrow += 1

            # Set a variable to pull out the full path of the shapefile or FC and set a variable to the source name         
            if row[4] == 'BCGW':
                dataSourcePath = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "BCGW.sde", row[5])
            elif row[4] == 'MTOPROD':
                dataSourcePath = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "MTOPROD.sde", row[5])
            else:
                dataSourcePath = os.path.join(row[4], row[5])    
             
            # Delete feature Layer. Need to do this because it will hang on the next loop.
            if arcpy.Exists("district"):
                delFeatLayer("district")
            else:
                pass       

            # Create feature layer in order to apply a definition query to the dataset
            arcpy.MakeFeatureLayer_management(dataSourcePath, "district")
              
            # Create a select by location to test for overlap.                                
            selFeats = arcpy.SelectLayerByLocation_management("district", "intersect", processedAOI)
              
            # Test to see if there are any records within each selected feature class
            # If it is zero, then let's output the layer name and a message indicating "No Overlap Found"
            # We'll also do some formatting on the cells                                
            if int(arcpy.GetCount_management(selFeats).getOutput(0)) == 0:
                sheetCells(sheet, excelrow, excelcol, "NA")
            else:
                with arcpy.da.SearchCursor("district", row[12]) as cursor:
                    for rowDistrict in cursor:
                        sheetCells(sheet, excelrow, excelcol, str(rowDistrict[0]))
                        excelrow += 1
     
            if excelrow > excelrowCount:
                excelrowCount = excelrow
 
            excelcol += 1
                        
            # Delete feature Layer. Need to do this because it will hang on the next loop.
            delFeatLayer("district")
                 
            excelrow = 2
                    
    excelrow = excelrowCount
    excelrow += 2
     
    excelcol = 1  
     
    sheetCells(sheet, excelrow, excelcol, "Location", 13, True, False, True)
     
    locationList = []
      
    for row in arcpy.da.SearchCursor(xls, xlsFields):
        if row[1] == 'Location':
            locationList.append(row[2])
      
    excelrow += 1
      
    staticexcelrow = excelrow
      
    for row in arcpy.da.SearchCursor(xls, xlsFields):
        if row[2] in sorted(locationList):
             
            excelrow = staticexcelrow
             
#             sheetCells(sheet, excelrow, excelcol, row[2], 12, True)
            sheetCells(sheet, excelrow, excelcol, '=HYPERLINK("{0}","{1}")'.format(str(row[26]), str(row[2])), 10, False, False, True, None)
       
            excelrow += 1
               
            # Set a variable to pull out the full path of the shapefile or FC and set a variable to the source name
            if row[4] == 'BCGW':
                dataSourcePath = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "BCGW.sde", row[5])
            elif row[4] == 'MTOPROD':
                dataSourcePath = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "MTOPROD.sde", row[5])
            else:
                dataSourcePath = os.path.join(row[4], row[5])            
 
            # Delete feature layer
            delFeatLayer("locale")       
       
            # Create feature layer in order to apply a definition query to the dataset
            arcpy.MakeFeatureLayer_management(dataSourcePath, "locale")
               
            # Create a select by location to test for overlap.                                
            selFeats = arcpy.SelectLayerByLocation_management("locale", "intersect", processedAOI)

            with arcpy.da.SearchCursor("locale", row[12]) as cursor:
                for rowDistrict in cursor:
                    sheetCells(sheet, excelrow, excelcol, str(rowDistrict[0]))
                    excelrow += 1
                   
            excelcol += 1
            excelrow += 1
             
    sheet.Columns.AutoFit()
        
    
def createMetadataSheet(book, scratchFolder, scratchGDB):
    '''
    A function to update the 'Input Information' sheet with parameter input information as set by the user
    '''
    
    arcpy.AddMessage("Creating metadata sheet")
    
    sheet = book.Worksheets("Input_Information")
    
    arcInstall = arcpy.GetInstallInfo()

    for key, value in list(arcInstall.items()):
        sheetCells(sheet, 1, 1, "Run on ArcGIS version: ", 10, True)
        sheetCells(sheet, 1, 2, str(arcInstall['ProductName'] + ': ' + arcInstall['Version']))

    sheetCells(sheet, 2, 1, "Input Feature Class for Area of Interest:", 10, True)
    sheetCells(sheet, 2, 2, AOI)

    sheetCells(sheet, 3, 1, "SQL Query Used on Input Feature Class:", 10, True)
    if sqlQuery:
        sheetCells(sheet, 3, 2, sqlQuery)
    else:
        sheetCells(sheet, 3, 2, "No AOI Query applied")
    
    sheetCells(sheet, 4, 1, "Output Folder Location of Excel Spreadsheet:", 10, True)
    sheetCells(sheet, 4, 2, output_excel) 
    
    sheetCells(sheet, 5, 1, "Configuration Excel Spreadsheet Location:", 10, True)
    sheetCells(sheet, 5, 2, xls)
    
    sheetCells(sheet, 6, 1, "Scratch Geodatabase Location:", 10, True)
    sheetCells(sheet, 6, 2, scratchFolder + "\\" + scratchGDB)

    sheetCells(sheet, 7, 1, "Pre-defined Layer List:", 10, True)
    sheetCells(sheet, 7, 2, pre_defined_layer_list_choice)

    sheetCells(sheet, 8, 1, "Layers used in Report:", 10, True)
    row = 8
    for lyr in layerList:
        sheetCells(sheet, row, 2, str(lyr))
        row += 1
    
    sheet.Columns.AutoFit()
    sheet.Rows.AutoFit()
    
def check_geomark(output_folder):
    '''
    This function takes input feature from a shapefile or
    featureclass and, using arcpy and the geomark api,
    creates a geomark shortcut in the chosen output location
    Details on how to use Geomark rest api can be found at:
    https://apps.gov.bc.ca/pub/geomark/docs/rest-api/
    '''
    arcpy.AddMessage("Generating Geomark...")
        
    permit_folders = ['mmd', 'Workarea', 'Victoria', 'Reclamation & Permitting', '_Statusing']
    normalPath = os.path.normpath(output_folder)
    folder_check = normalPath.split(os.sep)
    
    if set(permit_folders).issubset(folder_check) == True:
        
        for outpath, folder in os.path.split(output_folder):
        
            geomark_path = os.path.join(outpath, 'geomark')
            
            if os.path.exists(geomark_path):
                geomark_file = os.path.join(geomark_path, 'geomak_link.url')
                
                if os.path.exists(geomark_file):
                    with open(geomark_path, "r") as infile:
                         for line in infile:
                             if (line.startswith('URL')):
                                 geomarkInfoPage = line[4:]
                                 break
            else:
                os.mkdir(geomark_path)
                create_geomark(inFeatures, output_folder)

    else:   
        create_geomark(inFeatures, output_folder)

def create_geomark(inFeatures, output_folder):
    '''
    A function to create a geomark
    '''
    
    arcpy.FeatureClassToFeatureClass_conversion(inFeatures, output_folder, "AOI_geomark")
     
    # Geomark request URL.
    geomarkEnv = "https://apps.gov.bc.ca/pub/geomark/geomarks/new"
     
    headers = {"Accept": "*/*"}
    
    arcpy.env.overwriteOutput = True
    
    shp = os.path.join(output_folder, "AOI_geomark.shp")
    
    # Input file geometries will be submitted in the body of a POST request
    files = {"body": open(shp, "rb")}
    
    fname, file_extension = os.path.splitext(shp)
    
    fileFormat = file_extension.replace(".", "")
     
#     AOI_geomark = os.path.join(output_folder, "AOI_geomark")
     
    # Geomark Web Service request parameter values
    fields = {
        "allowOverlap": "false",
        "bufferCap": "ROUND",
        "bufferJoin": "ROUND",
        "bufferMetres": "",
        "bufferMitreLimit": "5",
        "bufferSegments": "8",
        "callback": "",
        "failureRedirectUrl": "",
        "format": fileFormat,
        "geometryType": "Polygon",
        "multiple": False,
        "redirectUrl": "",
        "resultFormat": "json",
        "srid": 3005
    }
     
    # PROCESSING
    arcpy.AddMessage("    Sending request to: " + geomarkEnv)
     
    # Submit request to the Geomark Web Service and parse response
     
    try:
        geomarkRequest = requests.post(
            geomarkEnv, files=files, headers=headers, data=fields)
        geomarkResponse = (
            str(geomarkRequest.text).replace("(", "").
            replace(")", "").replace(";", ""))
        data = json.loads(geomarkResponse)
        geomarkID = data["id"]
        print geomarkID
        geomarkInfoPage = data["url"]
    except (NameError, TypeError, KeyError, ValueError) as error:
        arcpy.AddMessage("    *****************************************************************")
        arcpy.AddMessage("    Error processing Geomark request for " + AOI_geomark)
        arcpy.AddMessage("    " + str(data["error"]))
        arcpy.AddMessage("*       ****************************************************************")
        exit()
     
    # DESCRIPTION
    """
    This script will add a geomark to a group with a secret key to remove expiry
    dates on the created geomark.
    """
    # GROUP and SECRET_KEY values are used to post the geomark to identified group, provided from DataBC 
    URL_BASE = "https://apps.gov.bc.ca/pub/geomark/"
    GROUP = "gg-43E47F2612574B119C535AF8CDCF2E2A"
    SECRET_KEY = "kg-9F97DC2E179541B7855EA388925CCBDF"
     
    TIMESTAMP = str(int(time.time() * 1000))
       
    addtogroup(geomarkID) 
    
    arcpy.AddMessage("    *****************************************************************")
    arcpy.AddMessage("    Geomark info page URL: " + geomarkInfoPage)
    arcpy.AddMessage("    *****************************************************************")
    
    with open(os.path.join(output_folder, "geomark_link.url"), "w") as text_file:
        text_file.write(geomarkInfoPage)
        text_file.close()

    return geomarkInfoPage

# Define the Geomark Group Functions
def url_encode(s):
    s1 = urllib.quote(s)
    s2 = str.replace(s1, '_', '%2F')
    return str.replace(s2, '-', '%2B')

def sign(message, key):
    print key
    key = bytes(key.encode('UTF-8'))
    message = bytes(message.encode('UTF-8'))
    digester = hmac.new(key, message, hashlib.sha1)
    signature = digester.digest()
    signature64 = base64.urlsafe_b64encode(signature)
    return str(signature64.encode('UTF-8'))

def addtogroup(geomark_id):
    SIGNATURE = sign("/geomarkGroups/" + GROUP + "/geomarks/add:" + TIMESTAMP + ":geomarkId=" + geomark_id, SECRET_KEY)
    SIGNATURE_ENCODED = url_encode(SIGNATURE)
    URL = URL_BASE + "geomarkGroups/" + GROUP + "/geomarks/add?geomarkId=" + geomark_id + "&signature=" + SIGNATURE_ENCODED + "&time=" + TIMESTAMP
    response = requests.post(URL, headers = {'Accept': 'application/json'})
    print("response is: ", response.json())
  
def removefromgroup(geomark_id):
    SIGNATURE = sign("/geomarkGroups/" + GROUP + "/geomarks/delete:" + TIMESTAMP + ":geomarkId=" + geomark_id, SECRET_KEY)
    SIGNATURE_ENCODED = url_encode(SIGNATURE)
    URL = URL_BASE + "geomarkGroups/" + GROUP + "/geomarks/delete?geomarkId=" + geomark_id + "&signature=" + SIGNATURE_ENCODED + "&time=" + TIMESTAMP
    response = requests.post(URL, headers = {'Accept': 'application/json'})
    print("response is: ", response.json()) 


# Log into BCGW and MTOPROD Oracle databases    
login(username, mtoprodpassword, bcgwpassword)

# Set scratch geodatabase
output_folder, scratchGDB = createScratchGDB(output_GDB)

# Process AOI to determine feature count and area in hectares
processedAOI, processedAOI_Hectares, iMapBCBaseURL = processAOI(AOI)

# Initialize an excel worksheet for the report
book, excel = initializeSpreadsheet()

# Set the database option for mineral titles datasets (BCGW or MTOPROD)
xls, xlsFields = getXLSData(excel)

# Process layers against AOI
layerListDict, collectFeatsCountDict = processData(processedAOI, processedAOI_Hectares, xls, xlsFields)

# Create the detailed Interest Report Sheet
crossReferenceDict = createInterestReportSheet(book, layerListDict)

#==================================================================================================================
'''
Commented out geomark until we decide how to best utilize this functionality
'''
# # Run Geomark tool if enabled
# if createGeomark == True:
#     geoMark_URL = check_geomark(output_folder)
# else:
#     geoMark_URL = ''
# # Create a summary sheet for the IOR
# createSummarySheet(book, excel, processedAOI, processedAOI_Hectares, collectFeatsCountDict, iMapBCBaseURL, crossReferenceDict, geoMark_URL)

#===================================================================================================================

# Create a summary sheet for the IOR
createSummarySheet(book, excel, processedAOI, processedAOI_Hectares, collectFeatsCountDict, iMapBCBaseURL, crossReferenceDict, '')

# Create a sheet that contains information about districts the AOI lies within
createDistrictSheet(book, xls, xlsFields, processedAOI)

# Create a metadata sheet to record user input information
createMetadataSheet(book, output_excel, scratchGDB)


# Save and close the workbook
book.SaveAs(output_excel + "\\" + "Interest_report_" + output_name + "_" + time.strftime('%Y%b%d') + ".xlsx")

# Quit the instance of excel from the process list in Task Manager
excel.Quit()

# Logout and remove connection Files
logout()
