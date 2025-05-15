import arcpy, itertools

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
    layerListDict = {}

    for row in arcpy.SearchCursor(xls):
            if str(row.getValue("Category")) not in ('District', 'Location'):
                if row.getValue("Category") in layerListDict:
                    layerListDict[row.getValue("Category")].append(str(row.getValue("Featureclass_Name")))              
                else:
                    layerListDict[row.getValue("Category")] = [str(row.getValue("Featureclass_Name"))]

    mineral_Coal_first = 'Mineral/Coal'

    mineral_coal_first_list = ['Tenure - Coal Leases', 
    'Tenure - Coal License Applications', 
    'Tenure - Coal Licenses', 
    'Tenure - Mineral Claims', 
    'Tenure - Mining Leases', 
    'Tenure - Placer Claims', 
    'Tenure - Placer Leases',
    'Reserves - Coal Land Reserves', 
    'Reserves - Mineral Conditional Reserves', 
    'Reserves - Mineral No Registration Reserves', 
    'Reserves - Placer Conditional Reserves', 
    'Reserves - Placer No Registration Reserves',
    'Placer Designated Claim and Lease Areas', 
    'Placer Designated Claim Areas',
    'Crown Granted 2 Post Mineral Claims', 
    'E&N Grants', 
    'Freehold Coal - Dominion', 
    'Freehold Coal - E&N', 
    'Freehold Coal - Provincial Statusing Dataset', 
    'Freehold Coal - Songhees_CadboroBay',
    'Notice of Work',
    'Coal Bed',
    'Coal Grid Units',
    'MTO Grid Cells', 
    'Southeast Coal Area Based Management Plan',
    '\n']

    
    if mineral_Coal_first in layerListDict:
        loop = itertools.chain([(mineral_Coal_first, mineral_coal_first_list)],((key,valueList) for (key,valueList) in sorted(layerListDict.iteritems()) if key != mineral_Coal_first))
    else:
        loop = itertools.chain((key,valueList) for (key,valueList) in sorted(layerListDict.iteritems()))

    vList = []

    for k,v in loop:
            vList += v

    self.params[4].filter.list = vList
    self.params[4].values = self.params[4].filter.list            

    del xls
    
    return

  def updateParameters(self):
    """Modify the values and properties of parameters before internal
    validation is performed.  This method is called whenever a parameter
    has been changed."""
    
    if not self.params[0].altered:
      self.params[0].value = r"Database Connections\MEMPRD.sde\MTA_SPATIAL.MTA_ACQUIRED_TENURE_SVW"

          
    if self.params[0].value:
      if str(self.params[0].value) == r"Database Connections\MEMPRD.sde\MTA_SPATIAL.MTA_ACQUIRED_TENURE_SVW":
        self.params[2].enabled = False
      elif str(self.params[0].value) == r"Database Connections\MEMPRD.sde\MTA_SPATIAL.MTA_SITE_SVW":
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

    return
