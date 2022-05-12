from Autodesk.Revit import DB
from flamingo.cobie import COBieParametersToCDX
from flamingo.cobie import CreateScheduleView
from pyrevit import forms, HOST_APP, revit
import clr
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
    except Exception as e:
        print(e)
        forms.alert(
            "Please setup the project for COBie using the BIM Interoperability "
            "Tools before running this command.",
            exitscript=True
        )
    
    # viewSchedules = DB.FilteredElementCollector(doc) \
    #     .OfClass(DB.ViewSchedule) \
    #     .ToElements()
    # viewScheduleNames = [view.Name for view in viewSchedules]
    # cdxScheduleViewData = {
    #     "GSA.Floor": DB.ElementId(DB.BuiltInCategory.OST_Levels),
    #     "GSA.Space (Rooms)": DB.ElementId(DB.BuiltInCategory.OST_Rooms),
    #     "GSA.Asset (Type)": DB.ElementId.InvalidElementId,
    #     "GSA.Asset (Instance)": DB.ElementId.InvalidElementId,
    # }
    # with revit.Transaction("Create CDX Schedules"):
    #     for viewName, builtinCategory in cdxScheduleViewData.items():
    #         if viewName not in viewScheduleNames:
    #             CreateScheduleView(
    #                 viewName,
    #                 builtinCategory,
    #             )
    COBieParametersToCDX(doc)