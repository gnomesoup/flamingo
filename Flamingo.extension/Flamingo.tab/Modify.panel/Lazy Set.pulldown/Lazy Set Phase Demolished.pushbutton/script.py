from Autodesk.Revit.DB import ElementId, Group
from pyrevit import HOST_APP, forms, revit, script
from flamingo.ensure import set_element_phase_demolished

if __name__ == "__main__":
    doc = HOST_APP.doc

    phases = doc.Phases
    phaseDictionary = {phase.Name: phase.Id for phase in phases}
    phaseNames = ["None"] + [phase.Name for phase in phases]
    
    phaseName = forms.ask_for_one_item(
        items=phaseNames,
        default=phaseNames[-1],
        prompt="Select phase to set as phase demolished",
    )
    if not phaseName:
        script.exit()
    selection = revit.get_selection()
    if selection.is_empty:
        selection = revit.pick_elements()
    else:
        selection = selection.elements
    
    with revit.Transaction("Lazy Set Phase Demolished"):
        if phaseName == "None":
            phaseId = ElementId.InvalidElementId
        else:
            phaseId = phaseDictionary[phaseName]
        groupedElementIds = [
            element.GetMemberIds() for element in selection if type(element) == Group
        ]
        # Flatten the list of lists
        groupedElementIds = [
            doc.GetElement(elementId)
            for group in groupedElementIds
            for elementId in group
        ]
        setElements = [
            set_element_phase_demolished(
                element,
                phaseId,
                doc
            ) for element in selection + groupedElementIds
        ]

    forms.alert(
        "Set the phase demolished to \"{}\" for {} element{}.".format(
            phaseName,
            len(setElements),
            "" if len(setElements) == 1 else "s"
        )
    )