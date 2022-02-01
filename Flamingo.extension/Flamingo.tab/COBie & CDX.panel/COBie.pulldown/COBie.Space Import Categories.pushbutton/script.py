from Autodesk.Revit import DB
from flamingo.revit import OpenDetached, GetScheduledParameterByName
from pyrevit import HOST_APP, clr, forms, revit, script

import System
clr.ImportExtensions(System.Linq)

def GetCategoriesBySpaceDescription(doc):
    try: 
        schedule = DB.FilteredElementCollector(originalDoc)\
            .OfClass(DB.ViewSchedule)\
            .WhereElementIsNotElementType()\
            .Where(lambda x: x.Name == "COBie.Space (Rooms)").First()
    except Exception:
        forms.alert(
            "Please setup the project for COBie using the BIM Interoperability "
            "Tools before running this command.",
            exitscript=True
        )
    cobieSpaceName = GetScheduledParameterByName(
        scheduleView=schedule,
        parameterName="COBie.Space.Description"
    )
    cobieSpaceCategory = GetScheduledParameterByName(
        scheduleView=schedule,
        parameterName="COBie.Space.Category"
    )
    rooms = DB.FilteredElementCollector(doc, schedule.Id)\
        .OfCategory(DB.BuiltInCategory.OST_Rooms)\
        .ToElements()
    
    categoryByDescription = {}
    for room in rooms:
        nameParameter = room.get_Parameter(cobieSpaceName.GuidValue)
        if nameParameter is None:
            continue
        name = nameParameter.AsString()
        if not name or name == "" or name in categoryByDescription:
            continue
        categoryParameter = room.get_Parameter(cobieSpaceCategory.GuidValue)
        if categoryParameter is None:
            continue
        category = categoryParameter.AsString()
        if not category or category == "":
            continue
        categoryByDescription[name] = category

    return categoryByDescription

def UpdateSpaceCategoryFromDictionary(
    dictionary, scheduleView, blankOnly=False, doc=None
):
    rooms = DB.FilteredElementCollector(doc, scheduleView.Id)
    cobieSpaceName = GetScheduledParameterByName(
        scheduleView=schedule,
        parameterName="COBie.Space.Description"
    )
    cobieSpaceCategory = GetScheduledParameterByName(
        scheduleView=schedule,
        parameterName="COBie.Space.Category"
    )
    with revit.Transaction("Import COBie.Space.Category"):
        for room in rooms:
            nameParameter = room.get_Parameter(cobieSpaceName.GuidValue)
            if nameParameter is None:
                continue
            name = nameParameter.AsString()
            if name == "" or name not in dictionary:
                continue
            categoryParameter = room.get_Parameter(cobieSpaceCategory.GuidValue)
            if blankOnly:
                category = categoryParameter.AsString()
                if category is not None or category != "":
                    continue
            categoryParameter.Set(dictionary[name])
    return

if __name__ == "__main__":
    
    doc = HOST_APP.doc
    try: 
        schedule = DB.FilteredElementCollector(doc)\
            .OfClass(DB.ViewSchedule)\
            .WhereElementIsNotElementType()\
            .Where(lambda x: x.Name == "COBie.Space (Rooms)").First()
    except:
        forms.alert(
            "Please setup the project for COBie using the BIM Interoperability "
            "Tools before running this command.",
            exitscript=True
        )

    originalRvtPath = forms.pick_file(
        file_ext="rvt",
        title="Select Revit file to import space categories"
    )

    if not originalRvtPath:
        script.exit()

    originalDoc = OpenDetached(filePath=originalRvtPath)
    # originalDoc = HOST_APP.doc

    categoryByDescription = GetCategoriesBySpaceDescription(originalDoc)
    for key, value in categoryByDescription.items():
        print("{}: {}".format(key, value))
    originalDoc.Close(False)

    selection = forms.CommandSwitchWindow.show(
        context=["All Values", "Blank Values Only"],
        message="Update COBie.Component.Space for:",
    )

    if not selection:
        script.exit()

    blankOnly = True if selection == "Blank Values Only" else False

    UpdateSpaceCategoryFromDictionary(
        dictionary=categoryByDescription,
        scheduleView=schedule,
        blankOnly=blankOnly,
        doc=doc
    )