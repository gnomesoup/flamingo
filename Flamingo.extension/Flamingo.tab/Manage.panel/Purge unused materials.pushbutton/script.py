import clr
from flamingo.revit import GetUnusedAssets, GetUnusedMaterials


from pyrevit import DB, HOST_APP, forms, revit, EXEC_PARAMS, script
import re
import System
clr.ImportExtensions(System.Linq)

if __name__ == "__main__":
    doc = HOST_APP.doc

    if EXEC_PARAMS.config_mode:
        protectedMaterialNames = None
    else:
        protectedMaterialNames = (
            r"^Analytical",
            r"^Dynamo",
            r"^Default",
            r"^Phase -",
            r"^Poche",
        )
    print(protectedMaterialNames)

    unusedMaterials = GetUnusedMaterials(doc=doc, nameFilter=protectedMaterialNames)
    if unusedMaterials:
        materialFormData = forms.SelectFromList.show(
            unusedMaterials,
            multiselect=True,
            button_name="Purge Selected",
            name_attr="Name",
            title="Select materials to purge"
        )

        if materialFormData:
            with revit.Transaction("Purge unused materials"):
                revit.delete.delete_elements(
                    element_list=materialFormData,
                    doc=doc
                )

    unusedAssets = GetUnusedAssets(doc)
    if unusedAssets:
        assetFormData = forms.SelectFromList.show(
            unusedAssets,
            multiselect=True,
            button_name="Purge Selected",
            name_attr="Name",
            title="Select assets to purge"
        )

        if assetFormData:
            with revit.Transaction("Purge unused assets"):
                revit.delete.delete_elements(
                    assetFormData,
                    doc=doc
                )

    if not unusedMaterials and not unusedAssets:
        forms.alert("There are no unused materials or assets to purge")