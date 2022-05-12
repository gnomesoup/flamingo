from Autodesk.Revit import DB
from flamingo.cobie import COBieComponentAutoSelect
from pyrevit import clr, HOST_APP, forms
import System
clr.ImportExtensions(System.Linq)

if __name__ == "__main__":
    doc = HOST_APP.doc
    try:
        view = DB.FilteredElementCollector(doc)\
            .OfClass(DB.ViewSchedule)\
            .WhereElementIsNotElementType()\
            .Where(lambda x: x.Name == "COBie.Type").First()
    except:
        forms.alert(
            "Please setup the project for COBie using the BIM Interoperability "
            "Tools before running this command.",
            exitscript=True
        )
    if forms.alert(
        "Are you sure you would like to update the COBie.Component status for "
        "all family instances in the current model?",
        yes=True,
        no=True,
    ):
        COBieComponentAutoSelect(view, doc=doc)
