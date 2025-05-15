#Interest Overlap Report Tool

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
