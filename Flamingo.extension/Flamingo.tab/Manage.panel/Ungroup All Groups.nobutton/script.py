from flamingo.revit import UngroupAllGroups
from pyrevit import forms, HOST_APP, revit

if __name__ == "__main__":
    doc = HOST_APP.doc
    with revit.Transaction("Ungroup All"):
        groupTypes = UngroupAllGroups(doc=doc)
        if forms.alert(
            "Would you like to purge all ungrouped, groups?",
            yes=True,
            no=True
        ):
            try:
                revit.delete.delete_elements(groupTypes)
            except Exception:
                print("Unable to remove all unused Group Types")

