from Autodesk.Revit.DB import Group
from pyrevit import HOST_APP, forms, revit
from flamingo.ensure import set_element_phase_created

if __name__ == "__main__":
    doc = HOST_APP.doc

    phases = doc.Phases
    phaseDictionary = {phase.Name: phase.Id for phase in phases}
    phaseNames = [phase.Name for phase in phases]

    phaseName = forms.ask_for_one_item(
        items=phaseNames,
        default=phaseNames[0],
        prompt="Select phase to set as phase created",
    )
    selection = revit.get_selection()
    if selection.is_empty:
        selection = revit.pick_elements()
    else:
        selection = selection.elements

    with revit.Transaction("Lazy Set Phase Created"):
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
            set_element_phase_created(element, phaseDictionary[phaseName], doc)
            for element in selection + groupedElementIds
        ]

    forms.alert(
        'Set the phase created to "{}" for {} element{}.'.format(
            phaseName, len(setElements), "" if len(setElements) == 1 else "s"
        )
    )
