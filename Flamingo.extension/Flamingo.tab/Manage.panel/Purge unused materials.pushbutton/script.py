from pyrevit import DB, HOST_APP, forms, revit, EXEC_PARAMS
import re
import clr
import System
clr.ImportExtensions(System.Linq)

def GetMaterialIds(element):
    try:
        elementMaterials = element.GetMaterialIds(False)
        paintMaterials = element.GetMaterialIds(True)
        for materialId in paintMaterials:
            elementMaterials.append(materialId)
        return elementMaterials
    except Exception as e:
        print(e)
        return []

doc = HOST_APP.doc

protectedMaterialNames = (
    r"^Analytical",
    r"^Dynamo",
    r"^Default",
    r"^Phase -",
    r"^Poche",
)

if not EXEC_PARAMS.config_mode:
    protectedMaterialNames = protectedMaterialNames + (r"^\d\d \d\d \d\d",)

protectedRegex = re.compile("|".join(protectedMaterialNames))

elements = DB.FilteredElementCollector(doc) \
    .WhereElementIsNotElementType() \
    .ToElements()

currentMaterials = DB.FilteredElementCollector(doc) \
    .OfClass(DB.Material) \
    .Where(
        lambda x: (
            not protectedRegex.match(x.Name)
        )
    )
currentMaterialIds = [
    element.Id for element in currentMaterials
]
categoryMaterialIds = [
    category.Material.Id for category in doc.Settings.Categories
    if category.Material is not None
]

subcategoryMaterialIds = [
    subcategory.Material.Id 
    for category in doc.Settings.Categories
    if category.CanAddSubcategory
    for subcategory in category.SubCategories
    if subcategory.Material is not None
]

usedMaterialIdsGrouped = [
    GetMaterialIds(element) for element in elements
    ]
usedMaterialIdsGrouped.append(categoryMaterialIds)
usedMaterialIdsGrouped.append(subcategoryMaterialIds)

materialIds = set(
    [
        materialId for idGroup in usedMaterialIdsGrouped
        for materialId in idGroup
    ]
)

unusedMaterials = [
    doc.GetElement(materialId) for materialId in currentMaterialIds
    if materialId not in materialIds
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
            revit.delete.delete_elements(doc, materialFormData)

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
            revit.delete.delete_elements(assetFormData)

if not unusedMaterials and not unusedAssets:
    forms.alert("There are no unused materials or assets to purge")