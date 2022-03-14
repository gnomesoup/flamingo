from pyrevit import HOST_APP, forms, revit, script
from flamingo.ensure import set_element_phase_demolished

if __name__ == "__main__":
    doc = HOST_APP.doc

    phases = doc.Phases
    phaseDictionary = {phase.Name: phase.Id for phase in phases}
    phaseNames = [phase.Name for phase in phases]
    
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
        setElements = [
            set_element_phase_demolished(
                element,
                phaseDictionary[phaseName],
                doc
            ) for element in selection
        ]

    forms.alert(
        "Set the phase demolished to \"{}\" for {} element{}.".format(
            phaseName,
            len(setElements),
            "" if len(setElements) == 1 else "s"
        )
    )