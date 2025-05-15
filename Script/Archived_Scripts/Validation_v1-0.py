import arcpy

class ToolValidator(object):
  """Class for validating a tool's parameter values and controlling
  the behavior of the tool's dialog."""

  def __init__(self):
    """Setup arcpy and the list of tool parameters."""
    self.params = arcpy.GetParameterInfo()

  def initializeParameters(self):
    """Refine the properties of a tool's parameters.  This method is
    called when the tool is opened."""
    from arcpy import env
    env.workspace = r"Database Connections\MEMPRD.sde"
    env.workspace = r"Database Connections\BCGW.sde"
    

    xls = r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\ArcGIS_Tools\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Excel_Spreadsheets\InterestReport_Layer_List_MASTER.xls\'MineralTitles Dataset selection$'"
    #vList = [str(i.getValue("Featureclass_Name")).rstrip() for i in arcpy.SearchCursor(xls) if str(i.getValue("Featureclass_Name")) <> 'None']
    vList = [str(i.getValue("Featureclass_Name")).rstrip() for i in arcpy.SearchCursor(xls) if str(i.getValue("Featureclass_Name")) <> 'None' and str(i.getValue("Category")) not in ('District', 'Location')]
    self.params[4].filter.list = sorted(vList)
    self.params[4].values = self.params[4].filter.list

    del xls
    
    return

  def updateParameters(self):
    """Modify the values and properties of parameters before internal
    validation is performed.  This method is called whenever a parameter
    has been changed."""
    
    if not self.params[0].altered:
      self.params[0].value = r"Database Connections\MEMPRD.sde\MTA_SPATIAL.MTA_ACQUIRED_TENURE_SVW"
      #self.params[1].value = ""
      #self.params[2].value = ""
      #self.params[2].enabled = False

          
    if self.params[0].value:
      if str(self.params[0].value) == r"Database Connections\MEMPRD.sde\MTA_SPATIAL.MTA_ACQUIRED_TENURE_SVW":
        #self.params[1].value = ""
        self.params[2].enabled = False
      elif str(self.params[0].value) == r"Database Connections\MEMPRD.sde\MTA_SPATIAL.MTA_SITE_SVW":
        #self.params[1].value = ""
        self.params[2].enabled = False
      elif str(self.params[0].value) not in [r"Database Connections\MEMPRD.sde\MTA_SPATIAL.MTA_ACQUIRED_TENURE_SVW", r"Database Connections\MEMPRD.sde\MTA_SPATIAL.MTA_SITE_SVW"] and not self.params[0].hasBeenValidated:
        self.params[2].enabled = True
        descFields = arcpy.Describe(self.params[0].value).fields
        shList = [str(field.name) for field in descFields]
        self.params[2].filter.list = shList
        self.params[2].values = self.params[2].filter.list
    
    return

  def updateMessages(self):
    """Modify the messages created by internal validation for each tool
    parameter.  This method is called after internal validation."""
##    if self.params[0].value <> self.params[0].value:
##      self.params[0].setErrorMessage(self.params[0].value)
    return
