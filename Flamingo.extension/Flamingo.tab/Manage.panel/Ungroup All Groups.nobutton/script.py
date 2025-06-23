from flamingo.revit import GetAllGroups
from pyrevit import forms, HOST_APP, revit, script

LOGGER = script.get_logger()

if __name__ == "__main__":
    doc = HOST_APP.doc
    selection = forms.CommandSwitchWindow.show(
        context=["Model Groups", "Detail Groups", "Both Detail & Model Groups"], title="Select Group Types to Ungroup"
    )
    modelGroupIncluded = False
    detailGroupIncluded = False
    if selection == "Model Groups":
        modelGroupIncluded = True
    elif selection == "Detail Groups":
        detailGroupIncluded = True
    elif selection == "Both Detail & Model Groups":
        modelGroupIncluded = True
        detailGroupIncluded = True
    else:
        script.exit()

    groups = GetAllGroups(
        includeDetailGroups=detailGroupIncluded,
        includeModelGroups=modelGroupIncluded,
        doc=doc,
    )
    groupTypes = set(group.GroupType for group in groups)

    with revit.Transaction("Ungroup All"):
        with forms.ProgressBar(title="Ungrouping all groups...") as pb:
            for i, group in enumerate(groups):
                pb.update_progress(i, len(groups))
                try:
                    group.UngroupMembers()
                except Exception as e:
                    LOGGER.warn("Failed to ungroup group: {}".format(group.Name))
                    LOGGER.debug(e)
        if forms.alert(
            "Would you like to purge all ungrouped groups?", yes=True, no=True
        ):
            try:
                revit.delete.delete_elements(groupTypes)
            except Exception:
                print("Unable to remove all unused Group Types")
