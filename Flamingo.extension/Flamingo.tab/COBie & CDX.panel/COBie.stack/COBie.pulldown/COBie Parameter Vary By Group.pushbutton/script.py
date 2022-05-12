# from Autodesk.Revit import DB
from Autodesk.Revit import DB
from flamingo.cobie import COBieUncheckAll, COBieParameterVaryByGroup
from flamingo.revit import GetAllElementsInModelGroups
from pyrevit import clr, forms, HOST_APP, script
import System
clr.ImportExtensions(System.Linq)

if __name__ == "__main__":
    doc = HOST_APP.doc
    output = script.get_output()

    # try:
    #     componentView = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)\
    #         .WhereElementIsNotElementType()\
    #         .Where(lambda x: x.Name == "COBie.Component").First()
    #     typeView = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)\
    #         .WhereElementIsNotElementType()\
    #         .Where(lambda x: x.Name == "COBie.Type").First()
    # except Exception:
    #     forms.alert(
    #         "Please setup the project for COBie using the BIM Interoperability "
    #         "Tools before running this command.",
    #         exitscript=True
    #     )
    
    # Attempted to uncheck the parameter myself but still run into group issues
    # unsetElements = COBieUncheckAll(typeView, componentView, doc=doc)

    # Attempting to create a new parameter as a placeholder for COBie checks
    # Copy over data from COBie parameter to placeholder
    # Remove COBie parameter from model
    # Add back in with group instances allowed to vary

    # COBieParameterVaryByGroup(typeView, doc)
    elements = GetAllElementsInModelGroups(doc)
    print(elements)

    print("")
    for element in elements:
        output.linkify(element.Id)

    print("Complete")