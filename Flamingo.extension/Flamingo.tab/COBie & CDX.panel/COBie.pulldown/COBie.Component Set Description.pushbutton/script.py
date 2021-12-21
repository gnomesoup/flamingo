from pyrevit import clr, DB, HOST_APP, forms, revit, script

from flamingo.cobie import COBieComponentSetDescription
import System
clr.ImportExtensions(System.Linq)

if __name__ == "__main__":

    selection = forms.CommandSwitchWindow.show(
        context=["All Values", "Blank Values Only"],
        message="Update COBie.Component.Space for:",
    )

    if not selection:
        script.exit()

    blankOnly = True if selection == "Blank Values Only" else False

    doc = HOST_APP.doc
    try:
        view = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)\
            .WhereElementIsNotElementType()\
            .Where(lambda x: x.Name == "COBie.Component").First()
    except:
        forms.alert(
            "Please setup the project for COBie using the BIM Interoperability "
            "Tools before running this command."
        )
    elements = COBieComponentSetDescription(view, blankOnly=blankOnly, doc=doc)
    n = len(elements)
    forms.alert(
        "Assigned Family and Type names to COBie.Component.Description"
        " for {} element{}".format(
            n, "" if n == 1 else "s"
        )
    )