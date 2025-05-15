'''
Tool name: Interest Overlap Report (IOR)
Developer: Mike MacRae for Ministry of Energy, Mines and Low Carbon Innovation
Contact: michael.macrae@gov.bc.ca or mineral.titles@gov.bc.ca
Date: Developed March 2014. Updated January 2021
'''

# Import required modules
import arcpy, win32com.client, os, itertools, gc, datetime, sys
from arcpy import env
from time import strftime
from getpass import getuser
from shutil import rmtree
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


def login(username, mtoprodpassword, bcgwpassword):
    ''' 
    A login prompt to get the users username and password for both MTOPROD and BCGW,
    create database connections for each and log into the databases
    '''
    arcpy.AddMessage("Checking to see if directory exists")
    if os.path.exists(os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser())):
        arcpy.AddMessage("Directory exists")
        pass
        #arcpy.Delete_management(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files" + '\\' + getuser())

    else:
        arcpy.AddMessage("Directory didn't exist")
        os.makedirs(os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser()))
        arcpy.AddMessage(os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser()))

    arcpy.AddMessage("Passed directory check")
    arcpy.AddMessage(getuser())
    arcpy.AddMessage(os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser()))
    
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
            #rmtree(os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser()))
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
        #rmtree(os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser()))
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
    #arcpy.env.outputCoordinateSystem = arcpy.SpatialReference("NAD_1983_BC_Environment_Albers")
    
    if output_folder == '':
    
        output_folder = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser())
        scratchGDB = "scratch.gdb"
        scratchLoc = output_folder + "\\" + scratchGDB
    
        delFeatLayer(scratchLoc)
        
        arcpy.CreateFileGDB_management(output_folder, scratchGDB)
        
    else:
                
        scratchGDB = "IOR_Clipped_FeatureClasses_" + output_name + "_" + strftime('%d%b%Y') + ".gdb"
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
    xls = r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Excel_Spreadsheets\InterestReport_Layer_List_MASTER.xlsx\MineralTitles_Dataset_selection$"

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

    return processedAOI, processedAOI_Hectares


def initializeSpreadsheet():
    '''
    A function to create an empty spreadsheet, add 4 required sheets and name them
    '''
    arcpy.AddMessage("Initializing spreadsheet...")
    
    #excel = win32com.client.Dispatch("Excel.Application")
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    #excel.Interactive = False
    excel.Visible = True
    
    # Initialize a workbook within excel
    book = excel.Workbooks.Add()
    
    book.Sheets.Add()
    
    # Set first sheet in book and rename it for the report
    book.Worksheets(1).Name = 'Summary'
    book.Worksheets(2).Name = 'Interest Report'
    book.Worksheets(3).Name = 'Districts and BCGS-NTS Location'
    book.Worksheets(4).Name = 'Input Information'

    return book, excel


def sheetCells(sheet, excelrow, excelcol, value="", size=10, bold=False, italic=False, underline=False, fontcolor=0, fillcolor=0, wrap=False):
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
    collectFeatsCountDict = {}
    featsCountDict = {}
    originalprocessAOI = processedAOI
    
    for row in arcpy.da.SearchCursor(xls, xlsFields): 

        if row[1] in layerList:

            featureclassNameTrans = ''.join(chr(c) if chr(c).isupper() or chr(c).islower() or chr(c).isdigit() else '_' for c in range(256))            
        
            fcName = row[1].replace(' ', '_')
            if row[5] == 'BCGW':
                fc = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "BCGW.sde", row[6])
            elif row[5] == 'MTOPROD':
                fc = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "MTOPROD.sde", row[6])
            else:
                fc = os.path.join(row[5], row[6])
                
            arcpy.AddMessage("Processing Layer: " + row[1])

            if row[11] is not None:
                arcpy.AddMessage("    " + "Buffering AOI for " + "'" + str(row[1]) + "' layer by " + str(int(row[11])))
                processedAOI = originalprocessAOI
                outBuffer = os.path.join(os.path.dirname(processedAOI), str(row[1]).translate(featureclassNameTrans) + "_" + str(int(row[11])) + "m_buffer")
                arcpy.Buffer_analysis(processedAOI, outBuffer, row[11])
                processedAOI = outBuffer
            else:
                processedAOI = originalprocessAOI           
            
            delFeatLayer("lyr")
            arcpy.MakeFeatureLayer_management(fc, "lyr")
            
            # Test for table joins
            if row[8] is not None:
                arcpy.AddJoin_management("lyr", str(row[9]), str(row[8]), str(row[10]))
                arcpy.AddMessage("    " + row[6] + " joined with " + os.path.basename(row[8]))
            else:
                pass
            
            lyr = arcpy.mapping.Layer("lyr")
            
            # Test to see if a Definition Query is needed
            if row[3] is not None:
                if lyr.supports("DEFINITIONQUERY"):
                    lyr.definitionQuery = row[3]
                    arcpy.AddMessage("    Def. Query applied")
                else:
                    arcpy.AddMessage("    Does not Support Definition Queries")
            else:
                arcpy.AddMessage("    No Definition Query")
                pass
            
            fieldList = [field.name for field in arcpy.ListFields(lyr) if field.name in [str(row[i]) for i in range (12, 25) if row[i] is not None]]
            
            fms = mappingFields(lyr, fieldList)
        
            selectresult = arcpy.GetCount_management(lyr)
             
            selectcount = int(selectresult.getOutput(0))
            arcpy.AddMessage("    Count before select: " + str(selectcount))

            arcpy.AddMessage("    Processing Select by Location")
            arcpy.SelectLayerByLocation_management(lyr, "intersect", processedAOI)

            selectresult = arcpy.GetCount_management(lyr)
            selectcount = int(selectresult.getOutput(0))       
 
            arcpy.AddMessage("    Count after selection: " + str(selectcount))
            
            outFC = str(row[1]).translate(featureclassNameTrans) + "_export"            

            # Test to see if records were selected during select by location
            if selectcount != 0:

                arcpy.AddMessage("    Exporting Selected Features")
                
                arcpy.FeatureClassToFeatureClass_conversion(lyr, output_folder + "\\" + scratchGDB, outFC, '', field_mapping=fms)
                
                # Describe the shapetype of each layer
                desc = arcpy.Describe(lyr)
                if desc.shapeType == "Polygon":

                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + outFC, "ORIGINAL_HECTARES", "FLOAT", "", "", "", "Original Area (Ha)")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + outFC, "ORIGINAL_HECTARES", "round(!shape.area!/10000,6)", "PYTHON_9.3") 
                    
                    arcpy.AddMessage("    Clipping Select Features with AOI")
                    arcpy.Clip_analysis(output_folder + "\\" + scratchGDB + "\\" + outFC, processedAOI, output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip")
                    
                    arcpy.Delete_management(output_folder + "\\" + scratchGDB + "\\" + outFC)
                    
                    if row[0] in layerListDict:
                        layerListDict[row[0]].update({row[1]: [str(outFC + "_clip"), str(row[26])]})
                    else:
                        layerListDict[row[0]] = {}
                        layerListDict[row[0]].update({row[1]: [str(outFC + "_clip"), str(row[26])]})
                    
                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip", "OVERLAPPING_HECTARES", "FLOAT", "", "", "", "Overlapping Area (Ha)")
                                                       
                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip", "OVERLAPPING_HECTARES", "round(!shape.area!/10000,6)", "PYTHON_9.3")

                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip", "PERCENT_OF_LAYER_BEING_OVERLAPPED_BY_AOI", "FLOAT", "", "", "", "% Layer being Overlapped by AOI")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip", "PERCENT_OF_LAYER_BEING_OVERLAPPED_BY_AOI", "round(!OVERLAPPING_HECTARES!/!ORIGINAL_HECTARES!*100,6)", "PYTHON_9.3")

                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip", "PERCENT_OF_AOI_BEING_OVERLAPPED_BY_LAYER", "FLOAT", "", "", "", "% AOI being Overlapped by Layer")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip", "PERCENT_OF_AOI_BEING_OVERLAPPED_BY_LAYER", "round(!OVERLAPPING_HECTARES!/{0}*100, 12)".format(processedAOI_Hectares), "PYTHON_9.3")

                                                   
                elif desc.shapeType == "Polyline":
                    
                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + outFC, "ORIGINAL_LENGTH", "FLOAT", "", "", "", "Original Length")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + outFC, "ORIGINAL_LENGTH", "round(!shape.length!/1000,6)", "PYTHON_9.3") 

                    arcpy.AddMessage("    Clipping Select Features with AOI")
                    arcpy.Clip_analysis(output_folder + "\\" + scratchGDB + "\\" + outFC, processedAOI, output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip")
                    
                    arcpy.Delete_management(output_folder + "\\" + scratchGDB + "\\" + outFC)

                    if row[0] in layerListDict:
                        layerListDict[row[0]].update({row[1]: [str(outFC + "_clip"), str(row[26])]})
                    else:
                        layerListDict[row[0]] = {}
                        layerListDict[row[0]].update({row[1]: [str(outFC + "_clip"), str(row[26])]})                    
                          
                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip", "OVERLAPPING_LENGTH", "FLOAT", "", "", "", "Overlapping Length")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip", "OVERLAPPING_LENGTH", "round(!shape.length!/1000,6)", "PYTHON_9.3") 

                elif desc.shapeType == "Point":
                    
                    arcpy.AddMessage("    Clipping Select Features with AOI")
                    arcpy.Clip_analysis(output_folder + "\\" + scratchGDB + "\\" + outFC, processedAOI, output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip")
                    
                    arcpy.Delete_management(output_folder + "\\" + scratchGDB + "\\" + outFC)

                    if row[0] in layerListDict:
                        layerListDict[row[0]].update({row[1]: [str(outFC + "_clip"), str(row[26])]})
                    else:
                        layerListDict[row[0]] = {}
                        layerListDict[row[0]].update({row[1]: [str(outFC + "_clip"), str(row[26])]})

                    arcpy.management.AddXY(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip")

                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip", "EASTING", "DOUBLE", "", 2)
                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip", "NORTHING", "DOUBLE", "", 2)

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip", "EASTING", "[POINT_X]")
                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip", "NORTHING", "[POINT_Y]")

                    arcpy.DeleteField_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip", ["POINT_X", "POINT_Y"])
                    
                elif desc.shapeType == "Multipoint":
                    
                    arcpy.AddMessage("    Clipping selected features with AOI")
                    arcpy.Clip_analysis(output_folder + "\\" + scratchGDB + "\\" + outFC, processedAOI, output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip")

                    arcpy.Delete_management(output_folder + "\\" + scratchGDB + "\\" + outFC)

                    if row[0] in layerListDict:
                        layerListDict[row[0]].update({row[1]: [str(outFC + "_clip"), str(row[26])]})
                    else:
                        layerListDict[row[0]] = {}
                        layerListDict[row[0]].update({row[1]: [str(outFC + "_clip"), str(row[26])]})

                    arcpy.AddField_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip", "POINT_LOCATION", "TEXT", "", "", 20, "Point Location")

                    arcpy.CalculateField_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip", "POINT_LOCATION", '"Point Location"', "PYTHON_9.3")
                  
            else:

                if row[0] in layerListDict:
                    layerListDict[row[0]].update({row[1]: ["No Overlap Found", str(row[26])]})
                else:
                    layerListDict[row[0]] = {}
                    layerListDict[row[0]].update({row[1]: ["No Overlap Found", str(row[26])]})

                arcpy.AddMessage("    No Overlap Found")
                
            if arcpy.Exists(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip"):
                countClippedFeats = int(arcpy.GetCount_management(output_folder + "\\" + scratchGDB + "\\" + outFC + "_clip").getOutput(0))

                if row[0] in collectFeatsCountDict:
                    collectFeatsCountDict[row[0]].update({row[1]: countClippedFeats})
                else:
                    collectFeatsCountDict[row[0]] = {}
                    collectFeatsCountDict[row[0]].update({row[1]: countClippedFeats})       
                
    tenureDict = {}
    reserveDict = {}
    remainingMiningDict = {}

    for category,value in sorted(collectFeatsCountDict.items()):
        if category == 'Mineral/Coal':
            for fClass, fCount in value.items():
                if 'Tenure - ' in fClass:
                    tenureDict[fClass] = fCount
                elif 'Reserve - ' in fClass:
                    reserveDict[fClass] = fCount
                else:
                    remainingMiningDict[fClass] = fCount            
            
    featsCountDict = itertools.chain(tenureDict.items(), reserveDict.items(), sorted(remainingMiningDict.items()))

    arcpy.AddMessage('===============================================================================')
    
    return layerListDict, featsCountDict


def createInterestReportSheet(book, layerListDict):
    '''
    A function to process data and populate an interest report sheet that provides details of each feature
    in each overlapping layer
    '''    

    arcpy.AddMessage("Creating Interest Report Detail sheet...")
    
    sheet = book.Worksheets("Interest Report")
    
    excelrow = 1
    excelcol = 1
    
    sheetCells(sheet, excelrow, excelcol, "INTEREST OVERLAP REPORT", 16, True)
    
    excelrow += 3

    sheetCells(sheet, excelrow, excelcol, "Legend", 12, True, False, True)
    
    excelrow += 1
    
    sheetCells(sheet, excelrow, excelcol, "", None, False, False, False, None, 15)
    sheetCells(sheet, excelrow, excelcol + 1, "Layer Category", 10, True, False, False, None, None)
    sheetCells(sheet, excelrow + 1, excelcol, "", None, False, False, False, None, 37)
    sheetCells(sheet, excelrow + 1, excelcol + 1, "Overlapping Layer", 10, True)
    sheetCells(sheet, excelrow + 2, excelcol, "National Park", 10, True, False, False, 46)
    sheetCells(sheet, excelrow + 2, excelcol + 1, "Alienated Lands", 10, True)
    
    # Place a border around the legend
    sheet.Range(sheet.Cells(excelrow,excelcol),sheet.Cells(excelrow + 2,excelcol + 1)).BorderAround()
    
    sheetCells(sheet, excelrow + 4, excelcol, "INTEREST OVERLAPS", 12, True, True)
    
    excelrow = 12      

    mineralcoal = {}
    tenureDict = {}
    reserveDict = {}
    remainingMiningDict = {}
    otherLayersDict = {}
    
    for key1, value1 in layerListDict.items():

        if key1 == 'Mineral/Coal':
            for key2, value2 in sorted(value1.items()):
                if 'Tenure' in key2:
                    tenureDict[key2] = value2
                elif 'Reserve' in key2 and key2 != 'Mineral Reserve':
                    reserveDict[key2] = value2
                else:
                    remainingMiningDict[key2] = value2      
            mineralcoal[key1] =  itertools.chain(sorted(tenureDict.items()), sorted(reserveDict.items()), sorted(remainingMiningDict.items()))
    
        else:       
            otherLayers = {}
            for key2, value2 in sorted(value1.items()):     
                otherLayers[key2] = value2
            otherLayersDict[key1] = itertools.chain(sorted(otherLayers.items()))
         
        processedlayerListDict = itertools.chain(mineralcoal.items(), sorted(otherLayersDict.items()))
        
        env.workspace = output_folder + "\\" + scratchGDB
        

    for key, value in processedlayerListDict:
        
        sheetCells(sheet, excelrow, excelcol, key, 12, True, False, False, None, 15)
        
        excelrow +=2
        
        for key2, value2 in value:
            
            if value2[1] != 'None':
                sheetCells(sheet, excelrow, excelcol, '=HYPERLINK("{0}","{1}")'.format(value2[1], str(key2)), 10, False, False, True, None)
            else:
                sheetCells(sheet, excelrow, excelcol, key2, 10, True, False, False, None)
            
            if value2[0] != 'No Overlap Found':
                for fc in arcpy.ListFeatureClasses():
                    if fc == value2[0]:
                        excelrow += 1
                        fields = [field for field in arcpy.ListFields(fc) if not field.required]
                        excelcol = 2
                        
                        for field in fields:
                            sheetCells(sheet, excelrow, excelcol, field.aliasName, 8, True, False, False, None, 37)
                            excelcol += 1
                        excelrow += 1
                        excelcol = 2
                        
                        for row in arcpy.da.SearchCursor(fc, [field.name for field in fields]):
                            for item in row:
                                sheetCells(sheet, excelrow, excelcol, item, 8, False, False, False, None)
                                excelcol += 1
                            excelcol = 2
                            excelrow += 1

            else:
                if value2[1] != 'None':
                    sheetCells(sheet, excelrow, excelcol, '=HYPERLINK("{0}","{1}")'.format(value2[1], str(key2)), 10, False, False, True, None)
                else:
                    sheetCells(sheet, excelrow, excelcol, key2, 10, True, False, False, None)
                    
                sheetCells(sheet, excelrow, excelcol + 1, "No Overlaps", 10, True, False, False)       
                
            excelcol = 1
            excelrow += 2
        excelrow += 1
    
    sheet.Columns.AutoFit()


def createSummarySheet(book, excel, processedAOI, processedAOI_Hectares, featsCountDict):
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
    
    # Add BC Government logo
    pic = sheet.Pictures().Insert(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Image\BC_EMLC_H_RGB_pos.jpg")
    
    # Format size of the logo
    pic.Height = 418
    pic.Width = 218
    
    # Format location of the logo
    cell = sheet.Cells(excelrow + 2, excelcol)
    
    pic.Left = cell.Left
    pic.Top = cell.Top
    
    # Add logo to report and set size and location
    pic = sheet.Pictures().Insert(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Image\Overlap_Example.JPG")
    
    # Format size of the logo
    pic.Height = 418
    pic.Width = 218
    
    # Format location of the logo
    cell = sheet.Cells(2,7)
    
    pic.Left = cell.Left
    pic.Top = cell.Top
    
    # Set the "REPORT FOR INTERNAL USE ONLY" comment in the sheet and do some formatting    
    sheetCells(sheet, excelrow, excelcol + 2, "REPORT FOR INTERNAL USE ONLY", 10, True, False, False, 3)
    
      # Set a comment to indicate when the report was run and do some formatting    
    sheetCells(sheet, excelrow + 1, excelcol + 2, "Report run on " + strftime('%d%b%y') + " @ " + strftime('%I:%M:%S'), 10)   
    
    excelrow += 9                 
    
    if os.path.basename(AOI) not in [r'MTA_SPATIAL.MTA_ACQUIRED_TENURE_SVW', r'MTA_SPATIAL.MTA_SITE_SVW', r'WHSE_MINERAL_TENURE.MTA_SITE_SP',r'WHSE_MINERAL_TENURE.MTA_ACQUIRED_TENURE_SVW']:
        
        sheetCells(sheet, excelrow, excelcol, "Area of Interest Information", 10)

        staticexcelrow = excelrow
        
        # Set a conditional sentence to determine if the user chose any fields to summarize in the area of interest parameter
        # In the first condition, determine if there are no fields to summarize
        if len(shFieldList) == 0:
            
            # Increment the counter by 8. This will push the rows paste the title header and the image
            # Also, set another variable to be used as a static row count which will be used to set the position of the legend
            # so that the legend lies next to the header information
            
            excelrow += 1
            
            sheetCells(sheet, excelrow, excelcol, "Area (ha):", 10, True)
                
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
                sheetCells(sheet, excelrow, excelcol, shField.replace("_"," "), 10, True, True)
                
                # Loop through the area of interest file and add the value that corresponds to the field name
    
                for shrow in arcpy.da.SearchCursor(AOI, [shField], sqlQuery):
                    sheetCells(sheet, excelrow, excelcol + 1, str(shrow[0]), 10)

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
        
        excelrow += 1
        staticexcelrow = excelrow
        
        fieldItemDict = {}
        with arcpy.da.SearchCursor(AOI, ["TENURE_NUMBER_ID", "TITLE_TYPE_DESCRIPTION", "ISSUE_DATE", "GOOD_TO_DATE", "OWNER_NAME", "AREA_IN_HECTARES", "SHAPE@AREA"], sqlQuery) as cursor:
            for row in cursor:
                fieldCount = 0

                for field in cursor.fields:
                    if field.title() == "Shape@Area":
                        fieldItemDict[field] = row[fieldCount]
                        sheetCells(sheet, excelrow, excelcol, "True Geometry Size (ha)", 10, True)
                        sheetCells(sheet, excelrow, excelcol + 1, str(round(row[fieldCount]/10000, 2)))
                    else:
                        fieldItemDict[field] = row[fieldCount]
                        sheetCells(sheet, excelrow, excelcol, str(field.title()), 10, True)
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
        
        excelrow += 1
        staticexcelrow = excelrow
        
        fieldItemDict = {}
        with arcpy.da.SearchCursor(AOI, ["SITE_NUMBER_ID", "RESERVE_TYPE", "MTA_SITE_ORDER_RESTR_DESC", "TOTAL_AREA", "SHAPE@AREA"], sqlQuery) as cursor:
            for row in cursor:

                fieldCount = 0

                for field in cursor.fields:
                    if field.title() == "Shape@Area":
                        fieldItemDict[field] = row[fieldCount]
                        sheetCells(sheet, excelrow, excelcol, "True Geometry Area (ha)", 10, True)
                        sheetCells(sheet, excelrow, excelcol + 1, str(round(row[fieldCount]/10000, 2)))
                    else:
                        fieldItemDict[field] = row[fieldCount]
                        sheetCells(sheet, excelrow, excelcol, str(field.title()), 10, True)
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

     
    #Set Report Header Title

    excelrow += 1
    
    sheetCells(sheet, excelrow, excelcol, "Counts of Mining Related Interest Overlaps (See Interest Report Sheet for Details)", 12, True)

    excelrow += 2
    
    for key, value in featsCountDict:
        sheetCells(sheet, excelrow, excelcol, key, 10)
        if value != 0:
            sheetCells(sheet, excelrow, excelcol + 1, value, 10, True)
        else:
            sheetCells(sheet, excelrow, excelcol + 1, value, 10)
        excelrow +=2
    
    sheet.Columns.AutoFit()
    sheet.Rows.AutoFit()


def createDistrictSheet(book, xls, xlsFields, processedAOI):
    ''' 
    A function to analyze various district types and maps sheets that overlaps the AOI
    A separate sheet is created and populated with the overlaps.
    '''    
     
    arcpy.AddMessage("Creating district information...")    
         
    sheet = book.Worksheets("Districts and BCGS-NTS Location")
     
    excelrow = 1
    excelcol = 1
    excelrowCount = 0
     
    sheetCells(sheet, excelrow, excelcol, "District Information", 16, True, False, True)
     
    excelrow += 1
     
    for row in arcpy.da.SearchCursor(xls, xlsFields):
        if row[0] == 'District':
                 
            sheetCells(sheet, excelrow, excelcol, row[1], 12, True)
              
            excelrow += 1
             
            # Set a variable to pull out the full path of the shapefile or FC and set a variable to the source name         
            if row[5] == 'BCGW':
                dataSourcePath = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "BCGW.sde", row[6])
            elif row[5] == 'MTOPROD':
                dataSourcePath = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "MTOPROD.sde", row[6])
            else:
                dataSourcePath = os.path.join(row[5], row[6])       
             
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
     
    sheetCells(sheet, excelrow, excelcol, "Location", 16, True, False, True)
     
    locationList = []
      
    for row in arcpy.da.SearchCursor(xls, xlsFields):
        if row[0] == 'Location':
            locationList.append(row[1])
      
    excelrow += 1
      
    staticexcelrow = excelrow
      
    for row in arcpy.da.SearchCursor(xls, xlsFields):
        if row[1] in sorted(locationList):
             
            excelrow = staticexcelrow
             
            sheetCells(sheet, excelrow, excelcol, row[1], 12, True)
       
            excelrow += 1
               
            # Set a variable to pull out the full path of the shapefile or FC and set a variable to the source name
            if row[5] == 'BCGW':
                dataSourcePath = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "BCGW.sde", row[6])
            elif row[5] == 'MTOPROD':
                dataSourcePath = os.path.join(r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Interim_Files", getuser(), "MTOPROD.sde", row[6])
            else:
                dataSourcePath = os.path.join(row[5], row[6])            
 
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
    
    sheet = book.Worksheets("Input Information")
    
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

# Log into BCGW and MTOPROD Oracle databases    
login(username, mtoprodpassword, bcgwpassword)

# Set scratch geodatabase
output_folder, scratchGDB = createScratchGDB(output_GDB)

# Process AOI to determine feature count and area in hectares
processedAOI, processedAOI_Hectares = processAOI(AOI)

# Set the database option for mineral titles datasets (BCGW or MTOPROD)
xls, xlsFields = getXLSData()

# Process layers against AOI
layerListDict, featsCountDict = processData(processedAOI, processedAOI_Hectares, xls, xlsFields)

# Initialize an excel worksheet for the report
book, excel = initializeSpreadsheet()

# Create the detailed Interest Report Sheet
createInterestReportSheet(book, layerListDict)

# Create a summary sheet for the IOR
createSummarySheet(book, excel, processedAOI, processedAOI_Hectares, featsCountDict)

# Create a sheet that contains information about districts the AOI lies within
createDistrictSheet(book, xls, xlsFields, processedAOI)

# Create a metadata sheet to record user input information
createMetadataSheet(book, output_excel, scratchGDB)

# Save and close the workbook
book.SaveAs(output_excel + "\\" + "Interest_report_" + output_name + "_" + strftime('%Y%b%d') + ".xlsx")

# Quit the instance of excel from the process list in Task Manager
excel.Quit()

# Logout and remove connection Files
logout()
