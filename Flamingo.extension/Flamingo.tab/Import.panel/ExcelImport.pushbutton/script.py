# -*- coding: utf-8 -*-
from pyrevit import clr, DB, HOST_APP, revit, forms, script
from flamingo.revit import SetNoteBlockProperties
from flamingo.excel import OpenWorkbook, GetWorksheetData

import System
clr.ImportExtensions(System.Linq)

doc = HOST_APP.doc

# Find the generic annotation family we will use to build our schedule
noteBlockSymbol = None
symbolId = None
symbolName = "Symbol-Schedule-Generic"
allSymbols = DB.FilteredElementCollector(doc). \
    OfCategory(DB.BuiltInCategory.OST_GenericAnnotation). \
    WhereElementIsElementType().\
    ToElements()
for symbol in allSymbols:
    if symbol.FamilyName == symbolName:
        symbolId = symbol.Id
        noteBlockSymbol = symbol
        break

# Alert users if the correct family is not loaded
if not symbolId:
    forms.alert(symbolId, "The " + symbolName
                + " family is not loaded into the project.",
                exitscript=True)


excelPath = forms.pick_excel_file(save=False,
    title="Select an excel file to import")
if not excelPath:
    script.exit()

# Open workbook
workbook = OpenWorkbook(excelPath)

# Present list of worksheets for selection
worksheets = {}
for worksheet in workbook.Worksheets:
    worksheets[worksheet.name] = worksheet
selection = forms.ask_for_one_item(items=worksheets.keys(), 
    default=worksheets.keys()[0],
    prompt="Select worksheet to import:")

if not selection:
    script.exit()
worksheet = worksheets[selection]

# Allow user to specify column/row skip & if header exists
# Request name for schedule.
# Provide list of already created schedules to choose from
elements = DB.FilteredElementCollector(doc)\
    .OfClass(DB.FamilyInstance)\
    .WherePasses(DB.FamilyInstanceFilter(doc, symbolId))\
    .ToElements()

worksheetParameterList = []
for e in elements:
    value = e.LookupParameter("Worksheet").AsString()
    if (value not in worksheetParameterList and
        value != '' and value != "Code Matrix"):
        worksheetParameterList.append(value)

scheduleName = None
while scheduleName is None:
    selection = forms.CommandSwitchWindow.show(
        worksheetParameterList + [worksheet.name, "ENTER CUSTOM NAME"],
        switches={"Include Header": True, "Group Rows": False},
        message="Select schedule to create/update"
    )
    scheduleName = selection[0]
    includeHeader= selection[1]["Include Header"]
    groupData = selection[1]["Group Rows"]

    if scheduleName == "ENTER CUSTOM NAME":
        scheduleName = forms.ask_for_string(
            default=worksheet.name,
            prompt="Enter name of schedule:")
    if not scheduleName:
        script.exit()

    if scheduleName in worksheetParameterList:
        if not forms.alert("You are about to overwrite the previously imported schedule \""
            + scheduleName + "\". Would you like to continue?",
            cancel=True):
            scheduleName = None

deleteList = []
if scheduleName in worksheetParameterList:
    for element in elements:
        parameterValue = element\
            .LookupParameter("Worksheet").AsString()
        if parameterValue == scheduleName:
            deleteList.append(element)

# Get a view to place families
viewName = "«" + scheduleName + "» Import Placeholder"
try:
    matchedView = DB.FilteredElementCollector(doc)\
        .OfClass(DB.ViewDrafting)\
        .WhereElementIsNotElementType()\
        .ToElements()\
        .Where(lambda x: x.Name == viewName)\
        .First()
except:
    matchedView = None

# Make a list of all the parameters in the family
metaParameterNames = ["Worksheet",
                "Group Name",
                "Group Sort",
                "Sort Order"]
dataParameterNames = ["Column" + str(x).zfill(2) for x in range(1, 11)]
parameterNames = metaParameterNames + dataParameterNames

# Get id of ViewFamilyType
draftingView = DB.FilteredElementCollector(doc)\
    .OfClass(DB.ViewFamilyType)\
    .ToElements()\
    .Where(lambda x: x.FamilyName == "Drafting View")\
    .First()

data = GetWorksheetData(worksheet, group=groupData)

with revit.Transaction("Import excel data"):
    if matchedView is None:
        matchedView = DB.ViewDrafting.Create(doc, draftingView.Id)
        matchedView.Name = viewName
        matchedView.Scale = 1
        parameter = matchedView.LookupParameter("Browser Main Folder")
        if parameter:
            parameter.Set("82 Notes Import")

    if deleteList:
        revit.delete.delete_elements(deleteList)
# Import spreadsheet data to generic annotation family
    headerNames = None
    for key, values in data.items():
        metaValues = values["meta"]
        rowValues = values["data"]
        if key == 1 and includeHeader:
            headerNames = rowValues
            continue
        xyz = DB.XYZ(0, float(key) * -3.0 / 16 / 12 , 0)
        rowFamily = doc.Create.NewFamilyInstance(xyz,
                                                noteBlockSymbol,
                                                matchedView)
        metaValues.insert(0, scheduleName) 
        metaValues.insert(3, key)
        for name, value in zip(metaParameterNames, metaValues):
            if value is not None:
                parameter = rowFamily.LookupParameter(name)
                parameter.Set(str(value))
        for name, value in zip(dataParameterNames, rowValues):
            if value is not None:
                parameter = rowFamily.LookupParameter(name)
                parameter.Set(str(value))

# Create schedule view if it doesn't exits
try:
    scheduleView = DB.FilteredElementCollector(doc)\
        .OfClass(DB.ViewSchedule)\
        .ToElements()\
        .Where(lambda x: x.Name == scheduleName)\
        .First()
except:
    scheduleView = None

if scheduleView is None:
    with revit.Transaction("Create import schedule"):
        scheduleView = DB.ViewSchedule.CreateNoteBlock(doc,
            noteBlockSymbol.Family.Id)
        SetNoteBlockProperties(
            scheduleView,
            scheduleName,
            rowFamily,
            metaParameterNames,
            dataParameterNames,
            headerNames=headerNames
        )
