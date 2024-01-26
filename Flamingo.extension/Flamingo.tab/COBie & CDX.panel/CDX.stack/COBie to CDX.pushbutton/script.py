from Autodesk.Revit import DB
from flamingo.cobie import COBieParametersToCDX
from flamingo.cobie import CreateScheduleView
from pyrevit import forms, HOST_APP, revit
import clr
import System
clr.ImportExtensions(System.Linq)

if __name__ == "__main__":
    doc = HOST_APP.doc

    try:
        view = DB.FilteredElementCollector(doc)\
            .OfClass(DB.ViewSchedule)\
            .WhereElementIsNotElementType()\
            .Where(lambda x: x.Name == "COBie.Component").First()
    except Exception as e:
        print(e)
        forms.alert(
            "Please setup the project for COBie using the BIM Interoperability "
            "Tools before running this command.",
            exitscript=True
        )
    
    COBieParametersToCDX(doc)