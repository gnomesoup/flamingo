from pyrevit import clr, DB, HOST_APP, forms, revit, script

import System
clr.ImportExtensions(System.Linq)

if __name__ == "__main__":

    selection = forms.ask_for_one_item(
        items=['Blank Values Only', 'All Values'],
        prompt="Would you like to update:",
        default="Blank Values Only"
    )
    if not selection:
        script.exit()

    doc = HOST_APP.doc
    view = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)\
        .WhereElementIsNotElementType()\
        .Where(lambda x: x.Name == "COBie.Component").First()

    elements = DB.FilteredElementCollector(doc, view.Id)\
        .WhereElementIsNotElementType().ToElements()
    
    n = 0
    with revit.Transaction("Set COBie.Component.Description"):
        for element in elements:
            isCobie = element.LookupParameter("COBie").AsInteger()
            if isCobie == 1:
                description = element.get_Parameter(
                    DB.BuiltInParameter.ELEM_TYPE_PARAM
                ).AsValueString()
                elementDescription = element.LookupParameter(
                    "COBie.Component.Description"
                    )
                if (
                    elementDescription.AsString() == "" or
                    elementDescription.AsString() == None or
                    selection == "All Values"
                ):
                    elementDescription.Set(description)
                    n += 1
    
    forms.alert(
        "Assigned Family and Type names to COBie.Component.Description"
        " for {} element{}".format(
            n, "" if n == 1 else "s"
        )
    )