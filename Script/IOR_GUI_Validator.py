import arcpy, itertools
from getpass import getuser
  

class ToolValidator(object):
  """Class for validating a tool's parameter values and controlling
  the behavior of the tool's dialog."""

  def __init__(self):
    """Setup arcpy and the list of tool parameters."""
    self.params = arcpy.GetParameterInfo() 

  def initializeParameters(self):
    """Refine the properties of a tool's parameters.  This method is
    called when the tool is opened."""

    self.params[1].enabled = False
    self.params[2].enabled = False

    valueList = self.processLayers(self.params[3].value)

    self.params[4].filter.list = valueList
    self.params[4].values = self.params[4].filter.list      

    self.params[8].value = getuser()
    self.params[9].value = ''
    self.params[10].value = ''
    
    return

  def updateParameters(self):
    """Modify the values and properties of parameters before internal
    validation is performed.  This method is called whenever a parameter
    has been changed."""

    if self.params[0].altered:
      result = arcpy.GetCount_management(self.params[0].value)
      if not self.params[0].hasBeenValidated:
        if int(result[0]) == 1:
          self.params[1].value = None
          self.params[1].enabled = False
        else:
          self.params[1].enabled = True
          
    if self.params[0].value:

        descFields = arcpy.Describe(self.params[0].value).fields
        shList = [str(field.name) for field in descFields if field.name.upper() not in ['GEOMETRY', 'SHAPE', 'SHAPE_LENGTH', 'GEOMETRY_LENGTH', 'SHAPE_AREA', 'GEOMETRY_AREA', 'OBJECTID', 'OBJECTID_1', 'FID', 'ID']]
        
        self.params[2].filter.list = shList
        self.params[2].values = self.params[2].filter.list
        self.params[2].enabled = True

    if self.params[3].altered:

      valueList = self.processLayers(self.params[3].value)

      if not self.params[3].hasBeenValidated:
        self.params[4].filter.list = valueList
        self.params[4].values = self.params[4].filter.list
      

    return

  def updateMessages(self):
    """Modify the messages created by internal validation for each tool
    parameter.  This method is called after internal validation."""

    if arcpy.ProductInfo() <> 'ArcInfo':
      self.params[0].setErrorMessage("The ArcGIS Desktop License level has not been set to ArcInfo. Please close ArcGIS Desktop and change the license to ArcInfo")
    else:
      pass

    if self.params[1].enabled:
      if self.params[1].value is None:
        self.params[1].setErrorMessage("Query is required on Areas of Interest that have more than one feature.")


    if self.params[0].value is not None:
      if self.params[1].enabled:
        if self.params[1].value is not None:

          arcpy.MakeFeatureLayer_management(self.params[0].value, "lyr", self.params[1].value)
          result = arcpy.GetCount_management("lyr")        

          if int(result[0]) == 0:
            self.params[1].setErrorMessage("Query applied to Area of Interest returns 0 features. \
                                           Please update query to include 1 feature only.")
          elif int(result[0]) > 1:
            self.params[1].setErrorMessage("Query applied to Area of Interest returns more than 1 feature. \
                                           Please update query to include 1 feature only.")
          else:
            pass
          
          arcpy.Delete_management("lyr")
      

    return

  def processLayers(self, layerGroup):
    """Process the layers that will be populated in the Layer List
    in the fifth parameter of the tool. Also, this will eveluate the set of predfined layers
    found in the Predefined Layer list found in the fourth parameter of the tool."""

    xls = r"\\spatialfiles.bcgov\Work\em\vic\mtb\Local\MTB_Scripts\MTB_Tools\Reporting_Tools\Interest_Overlap_Report\Excel_Spreadsheets\InterestReport_Layer_List_MASTER.xlsx\IOR_Data$"

    allLayers = []

    # Mining only Layers
    tenure = []
    reserve = []
    otherMiningLayers = []

    # Regional Layers
    pmaLayers = []
    scLayers = []
    swLayers = []
    seLayers = []
    nencLayers = []
    nwLayers = []

    with arcpy.da.SearchCursor(xls, '*') as cursor:

      for row in cursor:

        if row[30] == 'Y':
          pmaLayers.append(str(row[2]))        

        if row[31] == 'Y':
          nencLayers.append(str(row[2]))

        if row[32] == 'Y':
          nwLayers.append(str(row[2]))

        if row[33] == 'Y':
          scLayers.append(str(row[2]))

        if row[34] == 'Y':
          seLayers.append(str(row[2]))

        if row[35] == 'Y':
          swLayers.append(str(row[2]))
          

        if row[1] not in ('District', 'Location', 'Mineral/Coal'):
            allLayers.append(row[2])        

        if row[1] == 'Mineral/Coal':
            if 'Tenure -' in row[2]:
                tenure.append(str(row[2]))
            elif 'Reserves -' in row[2]:
                reserve.append(str(row[2]))
            else:
              otherMiningLayers.append(row[2])
              

    vList = []

    if layerGroup == 'All Layers':
      vList = tenure + reserve + otherMiningLayers + sorted(allLayers)        
    elif layerGroup == 'Mining Layers Only':
      vList = tenure + reserve + otherMiningLayers
    elif layerGroup == 'PMA Layers':
      vList = tenure + reserve + otherMiningLayers + sorted(pmaLayers)      
    elif layerGroup == 'North East\North Central Permitting Layers':
      vList = tenure + reserve + otherMiningLayers + sorted(nencLayers)
    elif layerGroup == 'North West Permitting Layers':
      vList = tenure + reserve + otherMiningLayers + sorted(nwLayers)          
    elif layerGroup == 'South Central Permitting Layers':
      vList = tenure + reserve + otherMiningLayers + sorted(scLayers)
    elif layerGroup == 'South East Permitting Layers':
      vList = tenure + reserve + otherMiningLayers + sorted(seLayers)      
    elif layerGroup == 'South West Permitting Layers':
      vList = tenure + reserve + otherMiningLayers + sorted(swLayers)

   

    del xls

    return vList
