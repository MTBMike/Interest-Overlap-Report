# Interest Overlap Report Tool

- Documented by: Mike MacRae – April 10, 2014
- Updated: Robert Ehlert - May 16, 2017

**Intended use**
This tool will run an overlap report to determine if different land use types (layers) overlap an area of interest as determined by the user.

**System Requirements**

- Access to government Desktop Terminal Services (DTS GIS access)
-	ArcInfo license
-	Read access to:
    -	Database Connections\MTOPROD.sde
    - Database Connections\BCGW.sde (Access to sensitive data will require signing access agreements)
-	Read/Write access to: \\spatialfiles.bcgov\Work\em\vic\mtb

**Location of Files**

-	Tool location: \\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report
-	Master Layer File Excel Spreadsheet: \\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Excel_Spreadsheets\InterestReport_Layer_List_MASTER.xlsx


Known Limitations

-	Input: 
    -	Area of Interest Input must be a polygon file (it does not support line or point features at this time). If using a line or point file, create buffer of the feature and use the buffer polygon as the input.
    -	AOI must be a feature class or shapefile. Feature classes are preferred to maintain full column headings.
    -	Report will on run on one feature in the Area of Interest input polygon file. If it contains multiple features, you must query to one individual feature using the SQL Query dialogue
    -	For any files used as input, review the projection. KMZ files will need to be re-projected after being converted to feature class in order to get area.

-	Printing (Not recommended):
    -	Printing will need to be on 8.5X11 landscape orientation. You may have to adjust some columns widths to fit to the printable area. Make use of text wrapping in cells
    -	Printing in black and white makes the legend colors to appear the same shade of grey.

-	Errors: 
    -	Report may error out if your ArcCatalogue/Map session has timed out. restart your Arc product and rerun.
    - If any error is encountered, first try:
        1.	Closing all ArcGIS products and ensure all ArcGIS processes have been killed via Task manager. Restart ArcGIS and rerun. If error still exists;
        2.	Close DTA session and restart. Try running tool again.
    - Report any error to the MS Teams [IOR Updates](https://teams.microsoft.com/l/channel/19%3A3e3434fb3490467a86e172ffec8e5abc%40thread.tacv2/IOR%20Updates?groupId=7960a9bb-3ef0-487c-801b-ea3e0c71dd4a&tenantId=6fdb5200-3d0d-4a8a-b036-d3685e359adc) Channel


## Procedures

**Master Layer File Excel Spreadsheet:** The Layer Spreadsheet is a configuration file used to add/remove/update layers used in the tool.

**Updating the spreadsheet**
1.	Navigate to the spreadsheet location
2.	Before you open spreadsheet, right click on the .xls/properties/General tab and uncheck ‘Read Only’.
    - Note: remember to set each spreadsheet to ‘Read Only’ after you have made your edits.
4.	Open the spreadsheet
5.	Activate sheet name 'IOR Data'
6.	Update the list by adding a layer, removing a layer or edit existing layers.
  	1. Remove layers by removing the whole row in the spreadsheet
    2. Adding/editing layers in the spreadsheet:
  
| Column Name              | Description |
|--------------------------|-------------|
| GUID                     | A unique identifer for each layer            |
| Category                 | Categories to organize common layer themes            |
| Featureclass_Name        | The name of the layer as you want to see in the IOR Report. Tip: Use the DataBC Data Record name where available            |
| Restricted               | Is the data restricted in any way (Y or N)            |
| workspace_path           | The database or base path of the dataset. For data found in BCGW or MTOPROD, set the database name. For local files, set base path            |
| dataSource               | The name of the database table or local file            |
| shapeType                | The type of shape (i.e. point, line, poloygon)            |
| Definition_Query         | The definition query used for the layer. Tip: Validate the definiton query in ArcMpa/Pro before adding to the spreadsheet           |
| Query_Layer              | If the layer is a query layer, create the sql file, add it to ~Interest_Overlap_Report\sqls and enter the .sql name here           |
| Join_Table               | The name of the table wher ea join is required            |
| dataSource_Join_Field    | The data source field name for the join            |
| Join_Table_Field         | The join table field name for the join            |
| Buffer_Distance          | Enter a buffer distance (in metres) to report on layers that are within a buffer distance of the AOI            |
| Sort_Field               | The field used to sort the feature in the Interest Report sheet within the IOR. Enter fieldName,direction with no spaces (i.e. TENURE_NUMBER_ID,ACSENDING)            |
| Fields_to_Summarize      | Enter a field to summarize from the layer table            |
| Fields_to_Summarize2     | Enter a field to summarize from the layer table            |
| Fields_to_Summarize3     | Enter a field to summarize from the layer table            |
| Fields_to_Summarize4     | Enter a field to summarize from the layer table            |
| Fields_to_Summarize5     | Enter a field to summarize from the layer table            |
| Fields_to_Summarize6     | Enter a field to summarize from the layer table            |
| Fields_to_Summarize7     | Enter a field to summarize from the layer table            |
| Fields_to_Summarize8     | Enter a field to summarize from the layer table            |
| Fields_to_Summarize9     | Enter a field to summarize from the layer table            |
| Fields_to_Summarize10    | Enter a field to summarize from the layer table            |
| Fields_to_Summarize11    | Enter a field to summarize from the layer table            |
| Fields_to_Summarize12    | Enter a field to summarize from the layer table            |
| Fields_to_Summarize13    | Enter a field to summarize from the layer table            |
| map_label_field          | Field to label on a map (Note: This is not currently implemented in the IOR            |
| DataBC_Metadata_Record   | The permalink of the DataBC Record or link to a readme file for the layer            |
| layerID                  | Enter a comma separated list of iMapBC layer ID's for the layer. Do not use spaces            |
| pma_Layers               | Permitted Mine Area Layers: Include the layer (add a Y) in this predefinied layer list            |
| nencLayers               | Northeast/North Central Region Layers: Include the layer (add a Y) in this predefinied layer list             |
| nwLayers                 | Northwest Region Layers: Include the layer (add a Y) in this predefinied layer list             |
| scLayers                 | Southcentral Region Layers: Include the layer (add a Y) in this predefinied layer list             |
| seLayers                 | Southeast Region Layers: Include the layer (add a Y) in this predefinied layer list             |
| swLayers                 | Southwest Region Layers: Include the layer (add a Y) in this predefinied layer list             |
| App                      | Enter nothing (Please see Mike). Used to add hyperlink to features in a layer, where a URL parameter to an external applicaiton can be used.        |
| ParameterField           | Enter nothing (Please see Mike). The field from the layers attribute table that will be used in the App hyperlink            |


**IMPORTANT**
- the path or name must be IDENTICAL, accounting for capitalization, spelling and special characters (underscores, periods, etc) as seen in ArcCatalogue. Best practice is to copy and paste path names and files names directly from the ‘Location’ toolbar and into the spreadsheet
- Definition Query must be INDENTICAL as it appears in the layers property dialogue box in ArcMap. Please build Definition Query in ArcMap layers property, validate and copy and paste directly from ArcMap into the spreadsheet
- Field names must be IDENTICAL, accounting for capitalization, spelling and special characters (underscores, periods, etc) as seen in ArcCatalogue. Best practice is to copy and paste field names from the feature class properties Fields tab, by right clicking on the feature class in ArcCatalogue

**Adding the tool to ArcMap/Catalogue**

1.	Open ArcToolbox with ArcMap/Catalogue/Pro
2.	Right click on the ‘ArcToolBox’ heading and ‘Add Toolbox...’
3.	Navigate to the following location and add the ‘MTB_Tools’ toolbox: \\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools.tbx

**Running the tool**

1.	Open MTB_Tools.tbx\Reporting_Tools\Interest_Overlap_Tools\Interest Overlap Report
3.	Enter the parameters for the tool (See below)
4.	Run the tool

The following is a screengrab of the tools dialogue box:
