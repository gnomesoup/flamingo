from Autodesk.Revit import DB
from flamingo.cobie import SetCOBieComponentSpace
from pyrevit import clr, forms, HOST_APP, script
import System
clr.ImportExtensions(System.Linq)

if __name__ == "__main__":
    doc = HOST_APP.doc
    # output = script.get_output()

    try:
        view = DB.FilteredElementCollector(doc)\
            .OfClass(DB.ViewSchedule)\
            .WhereElementIsNotElementType()\
            .Where(lambda x: x.Name == "COBie.Component").First()
    except:
        forms.alert(
            "Please setup the project for COBie using the BIM Interoperability "
            "Tools before running this command."
        )

    selection = forms.CommandSwitchWindow.show(
        context=["All Values", "Blank Values Only"],
        message="Update COBie.Component.Space for:",
    )

    if not selection:
        script.exit()

    blankOnly = True if selection == "Blank Values Only" else False

    SetCOBieComponentSpace(view, blankOnly, doc=doc)