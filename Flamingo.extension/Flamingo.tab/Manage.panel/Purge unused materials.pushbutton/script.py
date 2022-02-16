import clr
from flamingo.revit import GetAllProjectMaterialIds
from pyrevit import DB, HOST_APP, forms, revit, EXEC_PARAMS, script
import re
import System
clr.ImportExtensions(System.Linq)

if __name__ == "__main__":
    doc = HOST_APP.doc

    if not EXEC_PARAMS.config_mode:
        protectedMaterialNames = (
            r"^Analytical",
            r"^Dynamo",
            r"^Default",
            r"^Phase -",
            r"^Poche",
        )
    protectedRegex = re.compile("|".join(protectedMaterialNames))

    allMaterialIds = GetAllProjectMaterialIds(
        inUseOnly=False,
        nameFilter=protectedMaterialNames,
        doc=doc
    )
    usedMaterialIds = GetAllProjectMaterialIds(
        inUseOnly=True,
        nameFilter=protectedMaterialNames,
        doc=doc
    )

    unusedMaterials = [
        doc.GetElement(materialId) for materialId in allMaterialIds
        if materialId not in usedMaterialIds
    ]

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

    postPurgeMaterials = DB.FilteredElementCollector(doc) \
        .OfClass(DB.Material) \
        .ToElements()

    currentAssetIds = DB.FilteredElementCollector(doc) \
        .OfClass(DB.AppearanceAssetElement) \
        .ToElementIds()

    usedAssetIds = [
        material.AppearanceAssetId for material in postPurgeMaterials
    ]
    unusedAssets = [
        doc.GetElement(elementId) for elementId in currentAssetIds
        if elementId not in usedAssetIds if elementId is not None
    ]

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