from pyrevit import clr, DB, HOST_APP, forms, revit, script
from flamingo import COBieSpaceSetDescription

import System
clr.ImportExtensions(System.Linq)


if __name__ == "__main__":

    selection = forms.CommandSwitchWindow.show(
        context=["All Values", "Blank Values Only"],
        message="Would you like to update:",
    )

    if not selection:
        script.exit()

    doc = HOST_APP.doc
    blankOnly = selection == "Blank Values Only"
    try:
        view = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)\
            .WhereElementIsNotElementType()\
            .Where(lambda x: x.Name == "COBie.Space (Rooms)").First()
    except:
        forms.alert(
            "Please setup the project for COBie using the BIM Interoperability "
            "Tools before running this command.",
            exitscript=True
        )
    
    with revit.Transaction("Assign COBie.Space Description"):
        rooms = COBieSpaceSetDescription(view, blankOnly, doc)

    n = len(rooms)
    forms.alert(
        "Copied room names to COBie.Space.Description for {} room{}".format(
            n, "" if n == 1 else "s"
        )
    )