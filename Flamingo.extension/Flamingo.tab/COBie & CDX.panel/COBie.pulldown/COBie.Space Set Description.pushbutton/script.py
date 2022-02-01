from pyrevit import clr, DB, HOST_APP, forms, revit, script

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
    view = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)\
        .WhereElementIsNotElementType()\
        .Where(lambda x: x.Name == "COBie.Space (Rooms)").First()

    rooms = DB.FilteredElementCollector(doc, view.Id)\
        .WhereElementIsNotElementType().ToElements()

    n = 0
    with revit.Transaction("Assign COBie.Space Description"):
        for room in rooms:
            roomName = room.get_Parameter(
                    DB.BuiltInParameter.ROOM_NAME
                ).AsString()
            spaceDescription = room.LookupParameter("COBie.Space.Description")
            spaceDescriptionValue = spaceDescription.AsString()
            if (
                spaceDescriptionValue == "" or
                spaceDescriptionValue == None or
                selection == "All Values"
            ):
                spaceDescription.Set(roomName)
                n += 1

    forms.alert(
        "Assigned room names to COBie.Space.Description for {} room{}".format(
            n, "" if n == 1 else "s"
        )
    )