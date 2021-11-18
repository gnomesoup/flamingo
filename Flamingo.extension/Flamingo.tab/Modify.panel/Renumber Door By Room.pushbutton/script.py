from Autodesk.Revit import DB
from pyrevit import revit, HOST_APP, forms, script
from flamingo.revit import DoorRenameByRoomNumber, FilterByCategory

if __name__ == "__main__":
    doc = HOST_APP.doc
    activeView = doc.ActiveView
    activePhaseParameter = activeView.get_Parameter(
        DB.BuiltInParameter.VIEW_PHASE
    )
    activePhaseId = activePhaseParameter.AsElementId()
    activePhase = doc.GetElement(activePhaseId)

    selection = revit.get_selection()
    if selection.is_empty:
        activeDesignOptionId = DB.DesignOption.GetActiveDesignOptionId(doc)
        elementDesignOptionFilter = DB.ElementDesignOptionFilter(
            activeDesignOptionId
        )
        doors = DB.FilteredElementCollector(doc, activeView.Id)\
            .OfCategory(DB.BuiltInCategory.OST_Doors)\
            .WhereElementIsNotElementType()\
            .WherePasses(elementDesignOptionFilter)\
            .ToElements()
        response = forms.alert(
            msg="Are you sure you would like to renumber all {} active door{} "
                "in this view?\n\nTo limit the number of doors numbered, first "
                "select the elements and run this command again.".format(
                    len(doors), "" if len(doors) == 1 else "s"
                ),
            yes=True,
            cancel=True,
        )
        if not response:
            script.exit()
    else:
        doors = FilterByCategory(
            elements=selection.elements,
            builtInCategory=DB.BuiltInCategory.OST_Doors
        )
    DoorRenameByRoomNumber(
        doors=doors,
        phase=activePhase,
        doc=doc
    )