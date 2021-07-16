# -*- coding: utf-8 -*-
from pyrevit import DB
from pyrevit import revit
from pyrevit import clr
from pyrevit import HOST_APP

clr.AddReference("System")
from System.Collections.Generic import List

doc = HOST_APP.doc

print("Hiding unplaced view tags in current view:\r")
# Get the current view
view = doc.ActiveView

# TODO: If current view is a sheet, hide on all non-drafting and non-3D views
# TODO: Provide list of sheet sets to run the script on

# Filter through all elements on the current view to find view tags.
categoryList = (
    DB.BuiltInCategory.OST_Viewers, 
    DB.BuiltInCategory.OST_Elev
)
categoriesTyped = List[DB.BuiltInCategory](categoryList)
categoryFilter = DB.ElementMulticategoryFilter(categoriesTyped)
viewTags = DB.FilteredElementCollector(doc, view.Id) \
    .WhereElementIsNotElementType() \
    .WherePasses(categoryFilter)

# Go through all filtered elements to figure out if they have a sheet number
# If they don't hide them by element
hideList = List[DB.ElementId]()
hideCount = 0
for element in viewTags:
    parameterList = element.GetParameters("Sheet Number")
    if len(parameterList) == 0 or parameterList[0].AsString() == "---":
        hideList.Add(element.Id)
        hideCount = hideCount + 1
    elif len(parameterList) == 0 or parameterList[0].AsString() == "-":
        hideList.Add(element.Id)

if len(hideList) > 0:
    try:
        with revit.Transaction("Hide unplaced views"):
            view.HideElements(hideList)
            print(str(hideCount) + " Views Hidden")
    except Exception as e:
        print(e)
else:
    print("No views were hidden")
