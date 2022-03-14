from Autodesk.Revit import DB
from flamingo.ensure import set_element_workset
from pyrevit import HOST_APP, forms, revit, script

if __name__ == "__main__":
    doc = HOST_APP.doc

    if not forms.check_workshared():
        script.exit()
    worksets = DB.FilteredWorksetCollector(doc).ToWorksets()
    worksetDictionary = {
        workset.Name: workset.Id
        for workset in worksets
        if workset.Kind == DB.WorksetKind.UserWorkset
    }
    worksetNames = sorted(worksetDictionary.keys())

    worksetName = forms.ask_for_one_item(
        items=worksetNames,
        default=worksetNames[0],
        prompt="Select workset to set",
    )
    if not worksetName:
        script.exit()
    selection = revit.get_selection()
    if selection.is_empty:
        selection = revit.pick_elements()
    else:
        selection = selection.elements

    with revit.Transaction("Lazy Set Workset"):
        setElements = [
            set_element_workset(
                element,
                worksetDictionary[worksetName],
                doc=doc,
            )
            for element in selection
        ]

    forms.alert(
        'Set the workset to "{}" for {} element{}.'.format(
            worksetName, len(setElements), "" if len(setElements) == 1 else "s"
        )
    )
