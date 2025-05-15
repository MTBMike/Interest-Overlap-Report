'''
Tool name: Interest Overlap Report (IOR)
Developer: Mike MacRae for the Ministry of Mines and Critical Minerals
Contact: michael.macrae@gov.bc.ca or mineral.titles@gov.bc.ca
Date: Developed March 2014. Updated May 2025
'''

# Import required modules
import arcpy
import sys
import re
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
createGeomark = arcpy.GetParameter(11)



def login(username, mtoprodpassword, bcgwpassword):
    ''' 
    A login prompt to get the users username and password for both MTOPROD and BCGW,
    create database connections for each and log into the databases
    '''

    if os.path.exists(os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser())):
        pass        

    else:
        os.makedirs(os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser()))

    arcpy.AddMessage("Passed directory check...")
    
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


def getXLSData():
    '''
    A function to set which Database to pull Mineral Titles data from and
    to create a python dictionary to store the dataset name and buffer distance for
    layers in the MASTER spreadsheet that ask for buffering
    '''
    arcpy.AddMessage("Getting XLS Data...")

    xls = r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Excel_Spreadsheets\InterestReport_Layer_List_MASTER.xlsx"
    
    IOR_Data = os.path.join(xls, 'IOR_Data$')
    apps_Data = os.path.join(xls, 'Apps$')

    IORData_Fields = [field.name for field in arcpy.ListFields(IOR_Data)]
    
    appDict={}
    for row in arcpy.da.SearchCursor(apps_Data,['App', 'URL']):
        if row[0] is not None:
            appDict[str(row[0])] = row[1]  

    return IOR_Data, IORData_Fields, appDict


def processAOI(AOI, output_folder, scratchGDB):
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
    
    if arcpy.Describe(processedAOI).spatialReference.factoryCode != 3005:
        
        arcpy.AddMessage("Re-projecting to BC Albers (WKID: 3005)...")
        
        out_coordinate_system = arcpy.SpatialReference(3005)
        projected_AOI = output_folder + "\\" + scratchGDB + "\\" + "AOI_projected"
        arcpy.Project_management(processedAOI, projected_AOI, out_coordinate_system)
        processedAOI = projected_AOI

    else:
        pass
    
    for row in arcpy.da.SearchCursor(processedAOI, ["SHAPE@TRUECENTROID"]):
        urlX, urlY = row[0]
        coords = str(urlX) + ',' + str(urlY)
    
    arcpy.MakeFeatureLayer_management(processedAOI, 'processedAOI')
    layer = arcpy.mapping.Layer('processedAOI')
    
    mxd = arcpy.mapping.MapDocument(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\templates\getScale_template.mxd")
    df = arcpy.mapping.ListDataFrames(mxd)[0]    
    
    arcpy.mapping.AddLayer(df, layer)

    df.scale = int(math.ceil(df.scale/500)*500)

    iMapBCBaseURL = 'https://arcmaps.gov.bc.ca/ess/hm/imap4m/?scale={0}&center={1}'.format(df.scale, coords + ',' + '3005')
    
    return processedAOI, processedAOI_Hectares, iMapBCBaseURL


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


def getLayerInfo(lyrDict, xlsRow, overlap):
    
    if overlap == True:
        if xlsRow[1] in lyrDict:
            lyrDict[xlsRow[1]].update({xlsRow[2]: [str(xlsRow[0] + "_clip"), str(xlsRow[28]), str(xlsRow[36]), str(xlsRow[37])]})
        else:
            lyrDict[xlsRow[1]] = {}
            lyrDict[xlsRow[1]].update({xlsRow[2]: [str(xlsRow[0] + "_clip"), str(xlsRow[28]), str(xlsRow[36]), str(xlsRow[37])]})
        
    else:
        if xlsRow[1] in lyrDict:
            lyrDict[xlsRow[1]].update({xlsRow[2]: ["No Overlap Found", str(xlsRow[28]), str(xlsRow[36]), str(xlsRow[37])]})
        else:
            lyrDict[xlsRow[1]] = {}
            lyrDict[xlsRow[1]].update({xlsRow[2]: ["No Overlap Found", str(xlsRow[27]), str(xlsRow[36]), str(xlsRow[37])]})
    
    return lyrDict


def chunks(lst, n):
    """
    Yield successive n-sized chunks from a list of items to be used to split up lists that have more than 1000 items into individual lsists of 1000 itmes.
    This will be used to aid query layers where there ar emore than 1000 overlapping items.
    """
    for i in xrange(0, len(lst), n):
        yield lst[i:i + n]
        

def processData(processedAOI, processedAOI_Hectares, IOR_Data, IORData_Fields, output_folder):
    ''' 
    A function to process layers to determine if there is an overlap and subsequently clips and overlaps.
    The process data is used further on in the script to report on a spreadsheet.
    '''    
    arcpy.AddMessage("    ")
    arcpy.AddMessage("Processing Layers...")

    layerListDict = {}
    collectFeatsCountDict = OrderedDict()
    originalprocessAOI = processedAOI
    mineral_coal = u'Mineral/Coal'
    
#     catList = list(sorted(set([row[1] for row in arcpy.da.SearchCursor(IOR_Data, IORData_Fields) if row[1] not in [u'District', u'Location']])))
#     catList.insert(0, catList.pop(catList.index(mineral_coal)))
    
    for row in arcpy.da.SearchCursor(IOR_Data, IORData_Fields):

        if row[2] in layerList:        
        
            fcName = row[2].replace(' ', '_')
            if row[4] == 'BCGW':
                fc = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "BCGW.sde", row[5])
            elif row[4] == 'MTOPROD':
                fc = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "MTOPROD.sde", row[5])
            else:
                fc = os.path.join(row[4], row[5])
                
            arcpy.AddMessage("  Processing Layer: " + row[2])

            if row[12] is not None:
                
                arcpy.AddMessage("    " + "Buffering AOI for " + "'" + str(row[2]) + "' layer by " + str(int(float(row[12]))))
                
                processedAOI = originalprocessAOI
                outBuffer = os.path.join(os.path.dirname(processedAOI), str(row[0]) + "_" + str(row[12]).replace('.','_') + "m_buffer")
                arcpy.Buffer_analysis(processedAOI, outBuffer, row[12])
                processedAOI = outBuffer
            else:
                processedAOI = originalprocessAOI
                
            delFeatLayer("lyr")
            
            arcpy.MakeFeatureLayer_management(fc, "lyr")
            
            # Test for table joins
            if row[9] is not None:
                arcpy.AddJoin_management("lyr", str(row[10]), str(row[9]), str(row[11]))
                arcpy.AddMessage("    " + row[5] + " joined with " + os.path.basename(row[9]))
            else:
                pass
            
            lyr = arcpy.mapping.Layer("lyr")

            # Test to see if a Definition Query is needed
            if row[7] is not None:
                if lyr.supports("DEFINITIONQUERY"):
                    lyr.definitionQuery = row[7]
                    arcpy.AddMessage("    Definition Query applied")
                else:
                    arcpy.AddMessage("    Does not Support Definition Queries")
            else:
                arcpy.AddMessage("    No Definition Query")
                pass
            
            if row[8] is None:
                fieldList = [field.name for field in arcpy.ListFields(lyr) if field.name in [str(row[i]) for i in range (14, 27) if row[i] is not None]]
                fms = mappingFields(lyr, fieldList)
            else:
                pass              
        
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
                
                # Process: Make Query Layer
                if row[8] is not None:
                    
                    toolPath = r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report"
                    
                    sql, whereColumn = row[8].split(';')

                    whereColumnAlias = 't.' + whereColumn
                    
                    whereList = [whereRow[0] for whereRow in arcpy.da.SearchCursor(lyr,whereColumn)]
                    
                    if len(whereList) == 1:
                        whereList = whereColumnAlias + ' in ' + str(tuple(whereList)).replace(",)", ")")

                    elif len(whereList) > 1000:
                        where_conditions = []
                        for chunk in chunks(whereList, 1000):
                            where_conditions.append(whereColumnAlias + " IN " + str(tuple(chunk)))

                        # Combine conditions using OR
                        whereList = " OR\n".join(where_conditions)
                        
                    else:
                        whereList = whereColumnAlias + ' in '  + str(tuple(whereList))

                    
                    for filename in os.listdir(os.path.join(toolPath, "sqls")):
                        
                        if filename.endswith('.sql'):
                            
                            if filename == sql:
                                
                                filepath = os.path.join(os.path.join(toolPath, "sqls"), filename)
                                
                                fd = open(filepath, 'r')
                                sqlFile = fd.read()
                                fd.close()

                    sqlFile = sqlFile.replace("update_query", whereList)
                    
                    arcpy.MakeQueryLayer_management(os.path.join(output_folder, row[4] + '.sde'), "queryOutput", sqlFile, "OBJECTID")
                
                    # Process: Feature Class to Feature Class
                    arcpy.FeatureClassToFeatureClass_conversion("queryOutput", output_folder + "\\" + scratchGDB, row[0])
                
                else:
                                      
                    arcpy.FeatureClassToFeatureClass_conversion(lyr, output_folder + "\\" + scratchGDB, row[0], '', field_mapping=fms)
                
                # Describe the shapetype of each layer
                desc = arcpy.Describe(lyr)
                if desc.shapeType == "Polygon":

                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0], "ORIGINAL_HECTARES", "DOUBLE", "", "", "", "Original Area (Ha)")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0], "ORIGINAL_HECTARES", "round(!shape.area!/10000,6)", "PYTHON_9.3") 
                    
                    arcpy.AddMessage("    Clipping Select Features with AOI")
                    arcpy.Clip_analysis(output_folder + "\\" + scratchGDB + "\\" + row[0], processedAOI, output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip")
                    
                    arcpy.Delete_management(output_folder + "\\" + scratchGDB + "\\" + row[0])
                    
                    layerListDict = getLayerInfo(layerListDict, row, True)
                    
                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "OVERLAPPING_HECTARES", "DOUBLE", "", "", "", "Overlapping Area (Ha)")
                                                       
                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "OVERLAPPING_HECTARES", "round(!shape.area!/10000,6)", "PYTHON_9.3")

                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "PERCENT_OF_LAYER_BEING_OVERLAPPED_BY_AOI", "DOUBLE", "", "", "", "% Layer being Overlapped by AOI")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "PERCENT_OF_LAYER_BEING_OVERLAPPED_BY_AOI", "round(!OVERLAPPING_HECTARES!/!ORIGINAL_HECTARES!*100,6)", "PYTHON_9.3")

                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "PERCENT_OF_AOI_BEING_OVERLAPPED_BY_LAYER", "DOUBLE", "", "", "", "% AOI being Overlapped by Layer")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "PERCENT_OF_AOI_BEING_OVERLAPPED_BY_LAYER", "round(!OVERLAPPING_HECTARES!/{0}*100, 12)".format(processedAOI_Hectares), "PYTHON_9.3")

                                                   
                elif desc.shapeType == "Polyline":
                    
                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0], "ORIGINAL_LENGTH", "FLOAT", "", "", "", "Original Length")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0], "ORIGINAL_LENGTH", "round(!shape.length!/1000,6)", "PYTHON_9.3") 

                    arcpy.AddMessage("    Clipping Select Features with AOI")
                    arcpy.Clip_analysis(output_folder + "\\" + scratchGDB + "\\" + row[0], processedAOI, output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip")
                    
                    arcpy.Delete_management(output_folder + "\\" + scratchGDB + "\\" + row[0])
                        
                    layerListDict = getLayerInfo(layerListDict, row, True)                   
                          
                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "OVERLAPPING_LENGTH", "FLOAT", "", "", "", "Overlapping Length")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "OVERLAPPING_LENGTH", "round(!shape.length!/1000,6)", "PYTHON_9.3") 

                elif desc.shapeType == "Point":
                    
                    arcpy.AddMessage("    Clipping Select Features with AOI")
                    arcpy.Clip_analysis(output_folder + "\\" + scratchGDB + "\\" + row[0], processedAOI, output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip")
                    
                    arcpy.Delete_management(output_folder + "\\" + scratchGDB + "\\" + row[0])

                    layerListDict = getLayerInfo(layerListDict, row, True)

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
                    
                    layerListDict = getLayerInfo(layerListDict, row, True)

                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "POINT_LOCATION", "TEXT", "", "", 20, "Point Location")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", "POINT_LOCATION", '"Point Location"', "PYTHON_9.3")
                    
                if row[13] is not None:
                    arcpy.AddMessage("    Sorting rows...")
                    
                    fieldsorted = [str(pair).split(',') for pair in row[13].split(';')]

                    arcpy.Sort_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip", output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip_sorted", fieldsorted)
                    arcpy.Delete_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip")
                    arcpy.Rename_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip_sorted", output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip")   
                  
            else:
                layerListDict = getLayerInfo(layerListDict, row, False)

                arcpy.AddMessage("    No Overlap Found")
       
            if arcpy.Exists(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip"):
                newDict = {}
                newDict[row[2]] = [int(arcpy.GetCount_management(output_folder + "\\" + scratchGDB + "\\" + row[0] + "_clip").getOutput(0)), row[29]]
                
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
              
    arcpy.AddMessage('    ')
    arcpy.AddMessage('Data Processed...')
    arcpy.AddMessage('===============================================================================')
    
    return layerListDict, collectFeatsCountDict


def sheetCells(sheet, excelrow, excelcol, value="", size=10, bold=False, italic=False, underline=False, fontcolor=0, fillcolor=0, wrap=False, numFormat=None):
    '''
    A function to format cells in each sheet
    '''

    if isinstance(value, int):
        value = int(value)
        numFormat = 0
    elif isinstance(value, float):
        value = float(value)
        numFormat = '0.' + '0' * 3
    elif isinstance(value, (str, unicode)):
        if value.isdigit():
            value = int(value)
            numFormat = 0
        else:
            try:
                float(value)
                value = float(value)
            except:
                value = value.encode('utf-8')
    elif isinstance(value, datetime.datetime):
        if value < datetime.datetime.strptime('01 01 1900', '%m %d %Y'):
            dateSplit = str(datetime.datetime.strptime(str(value.date()), '%Y-%m-%d').date()).split('-')
            value  = dateSplit[0] + '-' + datetime.date(1900, datetime.datetime.strptime(dateSplit[1], '%m').date().month, 1).strftime('%b') + '-' + dateSplit[2]
        else:
            numFormat = "yyyy-mmm-dd"

    sheet.Cells(excelrow, excelcol).Value = value
    sheet.Cells(excelrow, excelcol).Font.Name = 'BC Sans'
    sheet.Cells(excelrow, excelcol).Font.Size = size
    sheet.Cells(excelrow, excelcol).Font.Bold = bold
    sheet.Cells(excelrow, excelcol).Font.Italic = italic
    sheet.Cells(excelrow, excelcol).Font.Underline = underline
    sheet.Cells(excelrow, excelcol).Font.ColorIndex = fontcolor
    sheet.Cells(excelrow, excelcol).Interior.ColorIndex = fillcolor
    sheet.Cells(excelrow, excelcol).WrapText = wrap
    sheet.Cells(excelrow, excelcol).NumberFormat = numFormat
    sheet.Cells(excelrow, excelcol).HorizontalAlignment = win32com.client.constants.xlRight

##    numFormat=None
    
    
def initializeSpreadsheet():
    '''
    A function to create an empty spreadsheet, add 4 required sheets and name them
    '''
    arcpy.AddMessage("Initializing spreadsheet...")
    
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = True
    
    # Initialize a workbook within excel
    book = excel.Workbooks.Add()
    
    if book.Sheets.Count != 4:
        for i in range(book.Sheets.Count, 4):
            arcpy.AddMessage(i)
            book.Sheets.Add()
    
    # Set first sheet in book and rename it for the report
    book.Worksheets(1).Name = 'Summary'
    book.Worksheets(2).Name = 'Interest_Report'
    book.Worksheets(3).Name = 'Districts_and_BCGS-NTS_Location'
    book.Worksheets(4).Name = 'Input_Information'

    return book, excel    

def createInterestReportSheet(book, layerListDict, appsDict, output_folder):
    '''
    A function to process data and populate an interest report sheet that provides details of each feature
    in each overlapping layer
    '''    

    arcpy.AddMessage("Creating Interest Report Detail sheet...")
    
    sheet = book.Worksheets("Interest_Report")
    sheet.Activate()
    
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
    
    sheetCells(sheet, excelrow, excelcol, "INTEREST OVERLAPS (click to view Data BC Catalogue Record)", 12, True, True)
    
    excelrow += 2
    legendsearchrow = excelrow

    mineralcoal = {}
    tenureDict = {}
    reserveDict = {}
    remainingMiningDict = {}
    otherLayersDict = {}
    
    crossReferenceDict = {}
    
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
        
    catList = []

    for category, values in processedlayerListDict:
        
        arcpy.AddMessage("    Writing detail information from " + str(category) + " category...")

        catList.append(category)
        
        sheetCells(sheet, excelrow, excelcol, category, 12, True, False, False, None, 15)

        excelrow +=2

        for Fclass, listItems in values:
            
            arcpy.AddMessage("      Writing detail information from " + str(Fclass) + " layer...")

            arcpy.AddMessage("      FClass: " + str(Fclass))
            arcpy.AddMessage("      Listitems: " + str(listItems))
            arcpy.AddMessage("      Listitems[1]: " + str(listItems[1]))
            arcpy.AddMessage("      Excelrow: " + str(excelrow))
            
            if listItems[1] != 'None':
                sheetCells(sheet, excelrow, excelcol, '=HYPERLINK("{0}","{1}")'.format(listItems[1], Fclass), 10, False, False, True, None)
                sheetCells(sheet, excelrow+1, excelcol, '=HYPERLINK("#"&CELL("address",INDEX(A1:A555,MATCH("INTEREST OVERLAP REPORT",A1:A555,0),1)), "Go to Top")', 8, False, False, False, 53)
                crossReferenceDict[Fclass] = excelrow
            else:
                sheetCells(sheet, excelrow, excelcol, Fclass, 10, True, False, False, None)
                crossReferenceDict[Fclass] = excelrow
            
            arcpy.AddMessage("crossReferenceDict: " + str(crossReferenceDict))
            
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
                        grouprow = excelrow
                        fieldNames = [str(field.name) for field in fields]

                        for row in arcpy.da.SearchCursor(fc, fieldNames):

                            for index, item in enumerate(row):
                                if listItems[3] != 'None':
                                    ind = fieldNames.index(listItems[3])
                                    appURL = [v for k,v in appDict.iteritems() if k == listItems[2]][0]
                                    if index == ind:                                 
                                        sheetCells(sheet, excelrow, excelcol, '=HYPERLINK("{0}","{1}")'.format(appURL.format(str(item)), item), 8, False, False, False, None, numFormat=0)
                                    else:
                                        sheetCells(sheet, excelrow, excelcol, item, 8, False, False, False, None, numFormat=0)
                                else:
                                    sheetCells(sheet, excelrow, excelcol, item, 8, False, False, False, None, numFormat=0)
                                excelcol += 1                               
                                
                            excelcol = 2
                            excelrow += 1
                            
                        sheet.Range(sheet.Rows(grouprow), sheet.Rows(excelrow-1)).Rows.Group()
                        sheet.Outline.SummaryRow = win32com.client.constants.xlAbove
                        sheet.Outline.ShowLevels(RowLevels=2)  


            else:
                if listItems[1] != 'None':
                    sheetCells(sheet, excelrow, excelcol, '=HYPERLINK("{0}","{1}")'.format(listItems[1], Fclass), 10, False, False, True, None)
                    sheetCells(sheet, excelrow+1, excelcol, '=HYPERLINK("#"&CELL("address",INDEX(A1:A555,MATCH("INTEREST OVERLAP REPORT",A1:A555,0),1)), "Go to Top")', 8, False, False, False, 53)
                else:
                    sheetCells(sheet, excelrow, excelcol, Fclass, 10, True, False, False, None)
                    
                sheetCells(sheet, excelrow, excelcol + 1, "No Overlaps", 10, True, False, False)            

            excelcol = 1
            excelrow += 2

        excelrow += 1
    
    excelrow -= 3

    for category in catList:
        sheetCells(sheet, legendrow, excelcol, '=HYPERLINK("#"&CELL("address",INDEX(A{0}:A{1},MATCH("{2}",A{0}:A{1},0),1)), "{3}")'.format(legendsearchrow, excelrow, category, "Go to " + category + " category"), 10, False, False, True, None)
        legendrow += 1

    sheet.Columns.AutoFit()
    sheet.Cells(1, 1).Select()
    
    return crossReferenceDict


def createSummarySheet(book, excel, processedAOI, processedAOI_Hectares, collectFeatsCountDict, iMapBCBaseURL, crossReferenceDict, geoMark_URL):
    ''' 
    A function to process data and create a count summary of the mining layer overlaps
    '''
    
    arcpy.AddMessage("Creating summary sheet...")
    
    sheet = book.Worksheets('Summary')
    sheet.Activate()
    
    excel.ActiveWindow.DisplayGridlines = False
    
    excelrow = 1
    excelcol = 1
    
    arcpy.AddMessage("  Adding summary sheet headers and inserting imagery...")
    
    #Set Report Header Title   
    sheetCells(sheet, excelrow, excelcol, "INTEREST OVERLAP REPORT - SUMMARY", 16, True)
    
    # Set the "REPORT FOR INTERNAL USE ONLY" comment in the sheet and do some formatting    
    sheetCells(sheet, excelrow + 1, excelcol, "REPORT FOR INTERNAL USE ONLY", 10, True, False, False, 3)
    
    # Set a comment to indicate when the report was run and do some formatting    
    sheetCells(sheet, excelrow + 2, excelcol, "Report run on and data current as of: " + time.strftime('%b %d, %Y') + " @ " + time.strftime('%I:%M:%S'), 10)   
        
    # Add BC Government logo
    pic = sheet.Pictures().Insert(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Image\BC_MCM_H_RGB_pos.png")
    
    # Format size of the logo    
    pic.Height = 65
    pic.Width = 130
    
    # Format location of the logo
    cell = sheet.Cells(excelrow, excelcol + 1)
    pic.Left = cell.Left
    pic.Top = cell.Top
       
    excelrow = 1

    # Add overlap example image to the report and set size and location
    pic = sheet.Pictures().Insert(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Image\Overlap_Example.JPG")
    
    # Format size of the logo
    pic.Height = 390
    pic.Width = 190
    
    # Format location of the logo
    cell = sheet.Cells(excelrow,3)
    pic.Left = cell.Left + 40
    pic.Top = cell.Top
    
    excelrow = 5
    
    arcpy.AddMessage("  Writing Area of Interest (AOI) Details...")
        
    sheetCells(sheet, excelrow, excelcol, "AREA OF INTEREST INFORMATION", 10, bold=True, underline=True)
    
    if geoMark_URL != '':
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
            
        sheetCells(sheet, excelrow, excelcol + 1, format(processedAOI_Hectares, ","))

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
            sheetCells(sheet, excelrow, excelcol, shField.replace("_"," "), 10, False, False)
            
            # Loop through the area of interest file and add the value that corresponds to the field name

            for shrow in arcpy.da.SearchCursor(processedAOI, [shField], sqlQuery):
                sheetCells(sheet, excelrow, excelcol + 1, shrow[0], 10)
        
        excelrow += 1  
             
        sheetCells(sheet, excelrow, excelcol, "Area (ha)", 10, False)
            
        sheetCells(sheet, excelrow, excelcol + 1, format(processedAOI_Hectares, ","))

        # Do some formatting on the area of interest header information. Set the font size, border, fill color and alignment
        sheet.Range(sheet.Cells(staticexcelrow, excelcol),sheet.Cells(excelrow, excelcol + 1)).Font.Size = 10
        sheet.Range(sheet.Cells(staticexcelrow, excelcol),sheet.Cells(excelrow, excelcol + 1)).BorderAround()
        sheet.Range(sheet.Cells(staticexcelrow, excelcol),sheet.Cells(excelrow, excelcol + 1)).Interior.ColorIndex = 36
        sheet.Range(sheet.Cells(staticexcelrow, excelcol + 1),sheet.Cells(excelrow, excelcol + 1)).HorizontalAlignment = win32com.client.constants.xlRight                            
                
        # Increment the row count by 2 for the position of the first overlapping layer
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
        
        arcpy.AddMessage("  Writing summary information from " + str(category) + " category...")
        
        sheetCells(sheet, excelrow, excelcol, category, 10, True)
        excelrow += 1
        
        for Fclass, FclassValues in catList.iteritems():
            
            arcpy.AddMessage("    Writing summary information from " + str(Fclass) + " layer...")
            
            if FclassValues[1] is not None:
                sheetCells(sheet, excelrow, excelcol + 1, '=HYPERLINK(CELL("address",Interest_Report!A{0}),"{1}")'.format(crossReferenceDict[Fclass], [k for k in crossReferenceDict.keys() if k == Fclass][0]), 10, False, False, True, None)
                sheetCells(sheet, excelrow, excelcol + 3, '=HYPERLINK("{0}","{1}")'.format(iMapBCBaseURL + '&catalogLayers=' + FclassValues[1], 'View in iMapBC'), 10, False, False, True, None)
            else:
                sheetCells(sheet, excelrow, excelcol + 1, '=HYPERLINK(CELL("address",Interest_Report!A{0}),"{1}")'.format(crossReferenceDict[Fclass], [k for k in crossReferenceDict.keys() if k == Fclass][0]), 10, False, False, True, None)
            sheetCells(sheet, excelrow, excelcol + 2, FclassValues[0], 10, True)
            excelrow += 1

        excelrow +=1    
    
    sheet.Columns.AutoFit()
    sheet.Rows.AutoFit()
    sheet.Cells(1, 1).Select()


def createDistrictSheet(book, IOR_Data, IORData_Fields, processedAOI):
    ''' 
    A function to analyze various district types and maps sheets that overlaps the AOI
    A separate sheet is created and populated with the overlaps.
    '''    
     
    arcpy.AddMessage("Creating district information...")    
         
    sheet = book.Worksheets("Districts_and_BCGS-NTS_Location")
    sheet.Activate()
     
    excelrow = 1
    excelcol = 1
    excelrowCount = 0
     
    sheetCells(sheet, excelrow, excelcol, "District Information", 13, True, False, True)
     
    excelrow += 1
     
    for row in arcpy.da.SearchCursor(IOR_Data, IORData_Fields):
        if row[1] == 'District':
                 
            sheetCells(sheet, excelrow, excelcol, '=HYPERLINK("{0}","{1}")'.format(str(row[28]), str(row[2])), 10, False, False, True, None)
              
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
                with arcpy.da.SearchCursor("district", row[14]) as cursor:
                    for rowDistrict in cursor:
                        sheetCells(sheet, excelrow, excelcol, rowDistrict[0])
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
      
    for row in arcpy.da.SearchCursor(IOR_Data, IORData_Fields):
        if row[1] == 'Location':
            locationList.append(row[2])
      
    excelrow += 1
      
    staticexcelrow = excelrow
      
    for row in arcpy.da.SearchCursor(IOR_Data, IORData_Fields):
        if row[2] in sorted(locationList):
             
            excelrow = staticexcelrow
             
            sheetCells(sheet, excelrow, excelcol, '=HYPERLINK("{0}","{1}")'.format(str(row[28]), str(row[2])), 10, False, False, True, None)
       
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

            with arcpy.da.SearchCursor("locale", row[14]) as cursor:
                for rowDistrict in cursor:
                    sheetCells(sheet, excelrow, excelcol, rowDistrict[0])
                    excelrow += 1
                   
            excelcol += 1
            excelrow += 1
             
    sheet.Columns.AutoFit()
    sheet.Cells(1, 1).Select()
        
    
def createMetadataSheet(book, scratchFolder, scratchGDB):
    '''
    A function to update the 'Input Information' sheet with parameter input information as set by the user
    '''
    
    arcpy.AddMessage("Creating metadata sheet...")
    
    sheet = book.Worksheets("Input_Information")
    sheet.Activate()
    
    arcInstall = arcpy.GetInstallInfo()

    for key, value in list(arcInstall.items()):
        sheetCells(sheet, 1, 1, "Run on ArcGIS version: ", 10, True)
        sheetCells(sheet, 1, 2, arcInstall['ProductName'] + ': ' + arcInstall['Version'])

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
    sheetCells(sheet, 5, 2, IOR_Data)
    
    sheetCells(sheet, 6, 1, "Scratch Geodatabase Location:", 10, True)
    sheetCells(sheet, 6, 2, scratchFolder + "\\" + scratchGDB)

    sheetCells(sheet, 7, 1, "Pre-defined Layer List:", 10, True)
    sheetCells(sheet, 7, 2, pre_defined_layer_list_choice)

    sheetCells(sheet, 8, 1, "Layers used in Report:", 10, True)
    row = 8
    for lyr in layerList:
        sheetCells(sheet, row, 2, lyr)
        row += 1
    
    sheet.Columns.AutoFit()
    sheet.Rows.AutoFit()
    sheet.Cells(1, 1).Select()
    
def check_geomark(inFeatures, output_folder):
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
                geomark_file = os.path.join(geomark_path, 'geomark_link.url')
                
                if os.path.exists(geomark_file):
                    with open(geomark_path, "r") as infile:
                         for line in infile:
                             if (line.startswith('URL')):
                                 geomarkInfoPage = line[4:]
                                 break
            else:
                os.mkdir(geomark_path)
                geoMark = create_geomark(inFeatures, output_folder)

    else:   
        geoMark = create_geomark(inFeatures, output_folder)
        
    return geoMark

def create_geomark(inFeatures, output_folder):
    '''
    A function to create a geomark
    '''
    arcpy.env.overwriteOutput = True

    lockedSHPs=[]
    
    env.workspace = output_folder
    
    for shp in arcpy.ListFeatureClasses():

        basename = os.path.basename(shp)
#             inFeatures, file_extension = os.path.splitext(os.path.basename(gdb))

        try:
            delFeatLayer(shp)
        except: 
            lockedSHPs.append(basename)
    
    shpIndex = 1
    AOI_geomark = "AOI_geomark.shp"

    while shpIndex:

        if AOI_geomark not in lockedSHPs:
            arcpy.FeatureClassToFeatureClass_conversion(inFeatures, output_folder, AOI_geomark)
            break
        else:
            shpIndex+=1
            AOI_geomark = "scratch_" + str(shpIndex) + ".gdb"
        

    # Geomark request URL.
    geomarkEnv = "https://apps.gov.bc.ca/pub/geomark/geomarks/new"
    
    # Input file geometries will be submitted in the body of a POST request
    files = {"body": open(os.path.join(output_folder, AOI_geomark), "rb")}
        
    # Set headers
    headers = {"Accept": "*/*"}
    
    # check if statusing feature class is multi part
    with arcpy.da.SearchCursor(os.path.join(output_folder, AOI_geomark), ['SHAPE@']) as cursor:
        for row in cursor:
            if row[0].isMultipart:
                multi_part = "true"
            else:
                multi_part = "false"
    arcpy.AddMessage(multi_part)
    
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
        "format": "shp",
        "geometryType": "Polygon",
        "multiple": multi_part,
        "redirectUrl": "",
        "resultFormat": "json",
        "srid": 3005
    }
     
    # Submit request to the Geomark Web Service and parse response
    arcpy.AddMessage("    Sending request to: " + geomarkEnv)
    
    arcpy.AddMessage("    Files: " + str(files))
    arcpy.AddMessage("    Headers: " + str(headers))
    arcpy.AddMessage("    Fields: " + str(fields))
    
    if geomarkEnv is None:
        arcpy.AddMessage("geomarkEnv is None")
    if files is None:
        arcpy.AddMessage("files is None")
    if headers is None:
        arcpy.AddMessage("headers is None")
    if fields is None:
        arcpy.AddMessage("fields is None")               
     
    try:
        geomarkRequest = requests.post(geomarkEnv, files=files, headers=headers, data=fields,verify=False)
        geomarkResponse = (str(geomarkRequest.text).replace("(", "").replace(")", "").replace(";", ""))
        data = json.loads(geomarkResponse)
        geomarkID = data["id"]
        geomarkInfoPage = data["url"]
    except (NameError, TypeError, KeyError, ValueError) as error:
        arcpy.AddMessage("    *****************************************************************")
        arcpy.AddMessage("    Error processing Geomark request for " + str(AOI_geomark))
        arcpy.AddMessage("    " + str(data["error"]))
        arcpy.AddMessage("*       ****************************************************************")
        arcpy.Delete_management(AOI_geomark)

     
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
       
    addtogroup(geomarkID, URL_BASE, GROUP, SECRET_KEY, TIMESTAMP)
    
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

def addtogroup(geomark_id, URL_BASE, GROUP, SECRET_KEY, TIMESTAMP):
    SIGNATURE = sign("/geomarkGroups/" + GROUP + "/geomarks/add:" + TIMESTAMP + ":geomarkId=" + geomark_id, SECRET_KEY)
    SIGNATURE_ENCODED = url_encode(SIGNATURE)
    URL = URL_BASE + "geomarkGroups/" + GROUP + "/geomarks/add?geomarkId=" + geomark_id + "&signature=" + SIGNATURE_ENCODED + "&time=" + TIMESTAMP
    response = requests.post(URL, headers = {'Accept': 'application/json'})
    print("response is: ", response.json())
  
def removefromgroup(geomark_id, URL_BASE, GROUP, SECRET_KEY, TIMESTAMP):
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
processedAOI, processedAOI_Hectares, iMapBCBaseURL = processAOI(AOI, output_folder, scratchGDB)

#==================================================================================================================
'''
Commented out geomark until we decide how to best utilize this functionality
'''
# Run Geomark tool if enabled
if createGeomark == True:
    geoMark_URL = check_geomark(processedAOI, output_folder)
else:
    geoMark_URL = ''

#===================================================================================================================


# Set the database option for mineral titles datasets (BCGW or MTOPROD)
IOR_Data, IORData_Fields, appDict = getXLSData()

# Process layers against AOI
layerListDict, collectFeatsCountDict = processData(processedAOI, processedAOI_Hectares, IOR_Data, IORData_Fields, output_folder)

# #==================================================================================================================
# '''
# Commented out geomark until we decide how to best utilize this functionality
# '''
# # Run Geomark tool if enabled
# if createGeomark == True:
#     geoMark_URL = check_geomark(processedAOI, output_folder)
# else:
#     geoMark_URL = ''
# 
# #===================================================================================================================

# Initialize an excel worksheet for the report
book, excel = initializeSpreadsheet()

# Create the detailed Interest Report Sheet
crossReferenceDict = createInterestReportSheet(book, layerListDict, appDict, output_folder)

# Create a summary sheet for the IOR
createSummarySheet(book, excel, processedAOI, processedAOI_Hectares, collectFeatsCountDict, iMapBCBaseURL, crossReferenceDict, geoMark_URL)

# Create a sheet that contains information about districts the AOI lies within
createDistrictSheet(book, IOR_Data, IORData_Fields, processedAOI)

# Create a metadata sheet to record user input information
createMetadataSheet(book, output_excel, scratchGDB)

# Activate the Summary Sheet so when the sheet is initially opened, it opens on the Summary sheet
book.Worksheets("Summary").Activate()

# Save and close the workbook
book.SaveAs(output_excel + "\\" + "Interest_report_" + output_name + "_" + time.strftime('%Y%b%d') + ".xlsx")

# Quit the instance of excel from the process list in Task Manager
excel.Quit()

# Logout and remove connection Files
logout()
