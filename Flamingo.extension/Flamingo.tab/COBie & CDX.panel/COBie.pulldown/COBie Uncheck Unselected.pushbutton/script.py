# from Autodesk.Revit import DB
from Autodesk.Revit import DB
from flamingo.cobie import COBieUncheckAll
from pyrevit import clr, forms, HOST_APP
import System
clr.ImportExtensions(System.Linq)

if __name__ == "__main__":
    doc = HOST_APP.doc

    try:
        componentView = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)\
            .WhereElementIsNotElementType()\
            .Where(lambda x: x.Name == "COBie.Component").First()
        typeView = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)\
            .WhereElementIsNotElementType()\
            .Where(lambda x: x.Name == "COBie.Type").First()
    except Exception:
        forms.alert(
            "Please setup the project for COBie using the BIM Interoperability "
            "Tools before running this command.",
            exitscript=True
        )
    
    unsetElements = COBieUncheckAll(typeView, componentView, doc=doc)

    print("")
    print("len(unsetElements) = {}".format(len(unsetElements)))
    print("Complete")