from pyrevit import DB, HOST_APP, forms, revit, script
from os import listdir, path, walk
import re

# def GetFamilyWidth(family: DB.Family) -> float:
#     """
#     Returns the width of a family in decimal feet
#     """
#     return family.get_Parameter(
#         DB.BuiltInParameter.FAMILY_WIDTH_PARAMETER
#     ).AsDouble()

def PathIsFamilyPath(filepath):
    if (
        filepath.endswith(".rfa") and
        not re.search(r"\.\d{4}\.rfa", filepath)
    ):
        return True
    else:
        return False

if __name__ == "__main__":
    doc = HOST_APP.doc

    familyDirectory = forms.pick_folder(
        title="Select a folder to load families"
    )
    familiesToPlace = []
    for dirpath, dirnames, filenames in walk(familyDirectory):
        for filename in filenames:
            if PathIsFamilyPath(filename):
                familiesToPlace.append(path.join(familyDirectory, filename))

    if familiesToPlace:
        familyString = "families" if len(familiesToPlace) > 1 else "family"
        if not forms.alert(
            "Would you like to place {} {}?".format(
                len(familiesToPlace),
                familyString
            ),
            yes=True,
            no=True,
        ):
            script.exit()
    else:
        forms.alert(
            "There are no valid families in the specified folder to place.",
            exitscript=True
        )
        
    with revit.Transaction("Load families"):
        families = [
            revit.create.load_family(familyPath, doc)
            for familyPath in familiesToPlace
        ]
        
    # for family in families:
    #     print(family.)
    # with revit.Transaction("Place families"):

    print(len(families))
