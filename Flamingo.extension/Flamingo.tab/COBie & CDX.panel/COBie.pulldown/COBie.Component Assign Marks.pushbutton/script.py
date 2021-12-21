from Autodesk.Revit import DB
from flamingo.cobie import COBieComponentAssignMarks
from pyrevit import clr, HOST_APP
import System
clr.ImportExtensions(System.Linq)

if __name__ == "__main__":
    doc = HOST_APP.doc
    view = DB.FilteredElementCollector(doc)\
        .OfClass(DB.ViewSchedule)\
        .WhereElementIsNotElementType()\
        .Where(lambda x: x.Name == "COBie.Component").First()
    COBieComponentAssignMarks(view, doc=doc)