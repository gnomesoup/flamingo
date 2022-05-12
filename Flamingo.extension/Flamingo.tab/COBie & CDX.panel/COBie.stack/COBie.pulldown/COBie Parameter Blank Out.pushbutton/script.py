from flamingo.cobie import COBieParameterBlankOut
from pyrevit import clr, DB, HOST_APP, forms
import System
clr.ImportExtensions(System.Linq)

if __name__ == "__main__":
    
    doc = HOST_APP.doc
    schedules = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)\
        .WhereElementIsNotElementType()\
        .Where(lambda x: x.Name.startswith("COBie"))
    if not schedules:
        forms.alert(
            "Please setup the project for COBie using the BIM Interoperability "
            "Tools before running this command.",
            exitscript=True
        )
    COBieParameterBlankOut(schedules, doc=doc)