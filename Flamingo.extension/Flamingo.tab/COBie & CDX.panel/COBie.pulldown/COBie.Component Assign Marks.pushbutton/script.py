from Autodesk.Revit import DB
from flamingo.cobie import COBieComponentAssignMarks
from pyrevit import clr, forms, HOST_APP
import System
clr.ImportExtensions(System.Linq)

if __name__ == "__main__":
    doc = HOST_APP.doc
    try:
        view = DB.FilteredElementCollector(doc)\
            .OfClass(DB.ViewSchedule)\
            .WhereElementIsNotElementType()\
            .Where(lambda x: x.Name == "COBie.Component").First()
    except:
        forms.alert(
            "Please setup the project for COBie using the BIM Interoperability "
            "Tools before running this command.",
            exitscript=True
        )

    COBieComponentAssignMarks(view, doc=doc)