from pyrevit import DB, forms, HOST_APP, clr

clr.AddReference("RevitAPI")
import Autodesk.Revit.DB as DBX
doc = HOST_APP.doc

categories = [
    DB.BuiltInCategory.OST_Casework,
    DB.BuiltInCategory.OST_ElectricalEquipment,
    DB.BuiltInCategory.OST_ElectricalFixtures,
    DB.BuiltInCategory.OST_Entourage,
    DB.BuiltInCategory.OST_Furniture,
    DB.BuiltInCategory.OST_FurnitureSystems,
    DB.BuiltInCategory.OST_GenericModel,
    # DB.BuiltInCategory.OST_LightingDevices,
    DB.BuiltInCategory.OST_LightingFixtures,
    DB.BuiltInCategory.OST_MechanicalEquipment,
    DB.BuiltInCategory.OST_Parking,
    DB.BuiltInCategory.OST_Planting,
    DB.BuiltInCategory.OST_PlumbingFixtures,
    DB.BuiltInCategory.OST_Site,
    DB.BuiltInCategory.OST_SpecialityEquipment,

]

families = DB.FilteredElementCollector(doc) \
    .OfClass(DB.FamilySymbol) \
    .WhereElementIsElementType()

print("Count: {}".format(families.GetElementCount()))

for family in families:
    print(family.)
