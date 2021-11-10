# -*- coding: utf-8 -*-
from pyrevit import clr, DB, HOST_APP, revit, forms, script
from flamingo.revit import SetNoteBlockProperties, GetScheduleFields
from flamingo.excel import OpenWorkbook, GetWorksheetData
import os
import sys
from string import ascii_uppercase

import System
clr.ImportExtensions(System.Linq)

if __name__ == "__main__":
    doc = HOST_APP.doc

    # Find the generic annotation family we will use to build our schedule
    noteBlockSymbol = None
    symbolId = None
    symbolName = "Symbol-Flamingo-Excel-Import"
    hostVersion = HOST_APP.version
    familyPath = sys.path[1] + "\\" + hostVersion \
        + "\\" + symbolName + ".rfa"
    try:
        noteBlockSymbols = revit.ensure.ensure_family(
            family_name=symbolName,
            family_file=familyPath,
            doc=doc
        )
    except:
        forms.alert(
            msg="Unable to load required family {} from "
                "the Flamingo library.".format(symbolName),
            exitscript=True
        )
    try:
        noteBlockSymbol = noteBlockSymbols.First()
    except AttributeError:
        noteBlockSymbol = noteBlockSymbols

    symbolId = noteBlockSymbol.Id
    excelPath = forms.pick_excel_file(save=False,
        title="Select an excel file to import")
    if not excelPath:
        script.exit()

    # Open workbook
    workbook = OpenWorkbook(excelPath)

    # Present list of worksheets for selection
    workbookName = workbook.Name
    worksheets = {}
    for worksheet in workbook.Worksheets:
        worksheets[worksheet.name] = worksheet

    selections = forms.SelectFromList.show(
        context=worksheets.keys(),
        multiselect=True,
        button_name="Select Worksheets to Import"
    )

    if not selections:
        script.exit()

    elements = DB.FilteredElementCollector(doc)\
        .OfClass(DB.FamilyInstance)\
        .WherePasses(DB.FamilyInstanceFilter(doc, symbolId))\
        .ToElements()

    elements = list(elements)
    worksheetParameterList = []
    for element in elements:
        workbookValue = element.LookupParameter("Workbook").AsString()
        worksheetValue = element.LookupParameter("Worksheet").AsString()
        if (
            workbookValue == workbookName
            and worksheetValue not in worksheetParameterList
            and worksheetValue != ''
        ):
            worksheetParameterList.append(worksheetValue)

    # Allow user to specify column/row skip & if header exists
    # Request name for schedule.
    # Provide list of already created schedules to choose from
    # TODO: Remember settings for switches. Can we remember by excel filename?
    # TODO: Add a "?" button to take you to help
    commandSelections = forms.CommandSwitchWindow.show(
        ["Import"],
        switches={
            "Include Header": False,
            "Group Rows": False,
            "Overwrite Previous Imports": True
        },
        message="Select import options"
    )
    if not commandSelections[0]:
        script.exit()

    includeHeader = commandSelections[1]["Include Header"]
    groupData = commandSelections[1]["Group Rows"]
    overwrite = commandSelections[1]["Overwrite Previous Imports"]

    with revit.Transaction("Import excel data"):
        for selection in selections:
            worksheet = worksheets[selection]
            worksheetName = worksheet.name

            if (not overwrite and worksheetName in worksheetParameterList):
                i = 1
                while (
                    worksheetName in worksheetParameterList
                ):
                    if i > 100:
                        forms.alert(
                            "Unable to find a unique name for worksheet"
                            " {}. Try renaming the worksheet in Excel and"
                            " Try again.".format(worksheet.name),
                            exitscript=True
                        )
                    worksheetName = "{}{}".format(
                        worksheet.name, i
                    )
                    i += 1

            deleteList = []
            if worksheetName in worksheetParameterList:
                for element in elements:
                    workbookValue = element\
                        .LookupParameter("Workbook").AsString()
                    worksheetValue = element\
                        .LookupParameter("Worksheet").AsString()
                    if worksheetValue == worksheetName:
                        deleteList.append(element)

            # Get a view to place families
            viewName = "«" + worksheetName + "» Import Placeholder"
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
            metaParameterNames = [
                "Workbook",
                "Worksheet",
                "Row Number",
                "Sort Number",
                "Sort Name",
            ]
            dataParameterNames = [
                "Column{}".format(letter) for letter in ascii_uppercase
            ]

            # Get id of ViewFamilyType
            draftingView = DB.FilteredElementCollector(doc)\
                .OfClass(DB.ViewFamilyType)\
                .ToElements()\
                .Where(lambda x: x.FamilyName == "Drafting View")\
                .First()

            data = GetWorksheetData(worksheet, group=groupData)

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
                if type(key) is not int:
                    continue
                metaValues = values["meta"]
                rowValues = values["data"]
                if key == 1 and includeHeader:
                    headerNames = rowValues
                    continue
                xyz = DB.XYZ(0, float(key) * -3.0 / 16 / 12 , 0)
                rowFamily = doc.Create.NewFamilyInstance(
                    xyz,
                    noteBlockSymbol,
                    matchedView
                )
                metaValues["Workbook"] = workbookName
                metaValues["Worksheet"] = worksheetName
                for name, value in metaValues.items():
                    if value is not None:
                        parameter = rowFamily.LookupParameter(name)
                        parameter.Set(str(value))
                for name, value in zip(dataParameterNames, rowValues):
                    if value is not None:
                        parameter = rowFamily.LookupParameter(name)
                        parameter.Set(str(value))

            # Create schedule view if it doesn't exits
            # TODO: Get schedule filter for workbook & worksheet to check if 
            # is a duplicate instead of name
            scheduleName = "Excel-{}".format(worksheetName)
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
                        scheduleView=scheduleView,
                        viewName="Excel-{}".format(worksheetName),
                        familyInstance=rowFamily,
                        metaParameterNameList=metaParameterNames,
                        columnNameList=dataParameterNames[0:data["columnCount"]],
                        headerNames=headerNames,
                    )
                    fields = GetScheduleFields(scheduleView)
                    print("fields = {}".format(fields))
                    scheduleDefinition = scheduleView.Definition
                    scheduleDefinition.ClearFilters()
                    scheduleFilter = DB.ScheduleFilter(
                        fields["Workbook"].FieldId,
                        DB.ScheduleFilterType.Equal, workbookName
                    )
                    scheduleFilter = DB.ScheduleFilter(
                        fields["Worksheet"].FieldId,
                        DB.ScheduleFilterType.Equal, worksheetName
                    )
                    scheduleDefinition.AddFilter(scheduleFilter)
                    scheduleDefinition.ClearSortGroupFields()
                    groupSort = DB.ScheduleSortGroupField(
                        fields["Row Number"].FieldId
                    )
                    scheduleDefinition.AddSortGroupField(groupSort)
                    groupName = DB.ScheduleSortGroupField(
                        fields["Sort Name"].FieldId
                    )
                    groupName.ShowHeader = True
                    scheduleDefinition.AddSortGroupField(groupName)
                    sortOrder = DB.ScheduleSortGroupField(
                        fields["Sort Number"].FieldId
                    )
                    scheduleDefinition.AddSortGroupField(sortOrder)
