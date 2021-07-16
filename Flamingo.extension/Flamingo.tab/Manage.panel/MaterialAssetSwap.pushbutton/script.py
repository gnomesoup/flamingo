import clr
import System
import csv
from os import path
from datetime import datetime
from pyrevit import HOST_APP, revit, DB, script, forms
from flamingo.revit import GetModelDirectory
from flamingo.revit import GetParameterFromProjectInfo

clr.ImportExtensions(System.Linq)

doc = HOST_APP.doc
app = doc.Application

materials = DB.FilteredElementCollector(doc) \
    .OfClass(DB.Material) \
    .ToElements()

swapMaterialName = "00 00 00 Asset Swap Base"
swapAssetId = None
for material in materials:
    if material.Name == swapMaterialName:
        swapMaterial = material
        swapAssetId = swapMaterial.AppearanceAssetId
    elif material.Name == "Default":
        defaultMaterial = material

if swapAssetId is None:
    try:
        appearanceAsset = None
        appearanceAssets = DB.FilteredElementCollector(doc) \
            .OfClass(DB.AppearanceAssetElement) \
            .ToElements()
        for asset in appearanceAssets:
            if asset.Name == "Paper (White)":
                appearanceAsset = asset
        with revit.Transaction("Add default swap material"):
            if appearanceAsset is None:
                assets = app.GetAssets(DB.Visual.AssetType.Appearance)
                for asset in assets:
                    if asset.Name == "Prism-319":
                        renderAsset = asset
                        break
                appearanceAsset = DB.AppearanceAssetElement.Create(
                    doc, "Paper (White)", renderAsset
                )
            materialId = DB.Material.Create(doc, swapMaterialName)
            swapMaterial = doc.GetElement(materialId)
            swapAssetId = appearanceAsset.Id
            swapMaterial.AppearanceAssetId = swapAssetId
    except Exception as e:
        print(e)
        forms.alert(
            "Could create default material. Load \"00 00 00 Asset Swap Base\"" +
            " material from UNIFI and try again",
            exitscript=True
        ) 

assetMapPath = None
modelDirectory = GetModelDirectory(doc)
while assetMapPath is None:
    if forms.alert(
        "Would you like to save the current state of the material assets?",
        cancel=False, yes=True, no=True
    ):
        projectNumber = GetParameterFromProjectInfo(doc, "Project Number")
        assetMapFileName = projectNumber + " Asset Map " + \
            datetime.now().strftime("%Y-%m-%d") + ".csv"
        assetMapPath = forms.save_file(
            file_ext="csv",
            default_name=assetMapFileName,
            init_dir=modelDirectory
        )
        if not assetMapPath:
            assetMapPath = None
    else:
        assetMapPath = False

if assetMapPath:
    currentMaterialMap = []
    for material in materials:
        currentMaterialMap.append(
            [
                material.Id.IntegerValue,
                material.AppearanceAssetId.IntegerValue,
                material.UseRenderAppearanceForShading
            ]
        )

    with open(assetMapPath, "wb") as f:
        writer = csv.writer(f)
        writer.writerows(currentMaterialMap)

formOut = forms.CommandSwitchWindow.show(
    context=[
        "Swap all materials to base asset",
        "Swap all but selected to base asset",
        "Load assets from a file",
    ],
    title="Material Asset Swap",
    switches={
        "Do not change glass materials": True,
        "Do not change test fit floors": True
    }
)

if formOut[0] is None:
    script.exit()

excludeGlazing = formOut[1]["Do not change glass materials"]
# TODO Ignore test fit floor materials
excludeTestFitFloors = formOut[1]["Do not change test fit floors"]
# purgeUnusedMaterials = True

if formOut[0] == "Load assets from a file":
    assetMapPath = forms.pick_file(
        file_ext="csv",
        init_dir=modelDirectory,
        title="Select asset map csv file"
    )
    if not assetMapPath:
        script.exit()
    with open(assetMapPath, "rb") as f:
        reader = csv.reader(f)
        assetMap = [row for row in reader]
else:
    assetMap = None

if assetMap:
    with revit.Transaction("Assign assets from file"):
        for row in assetMap:
            try:
                material = doc.GetElement(DB.ElementId(int(row[0])))
                material.AppearanceAssetId = DB.ElementId(int(row[1]))
                if len(row) > 2:
                    material.UseRenderAppearanceForShading = row[2]
            except Exception as e:
                print("Asset assignment warning: " + str(e))

else:
    ignoreMaterials = []
    if formOut[0] == "Swap all but selected to base asset":
        elements = revit.pick_elements(
            "Select elements to remain un-swapped. Click \"Finish\" in the " +
            "options bar when done."
        )
        for element in elements:
            elementMaterials = element.GetMaterialIds(False)
            for materialId in elementMaterials:
                if materialId not in ignoreMaterials:
                    ignoreMaterials.append(materialId)
            paintMaterials = element.GetMaterialIds(True)
            for materialId in paintMaterials:
                if materialId not in ignoreMaterials:
                    ignoreMaterials.append(materialId)

    with revit.Transaction("Swap material assets"):
        osCategories = doc.Settings.Categories
        for category in osCategories:
                if (
                    category.Material is None
                    and category.CategoryType == DB.CategoryType.Model
                ):
                    try:
                        category.Material = defaultMaterial
                    except Exception as e:
                        print("Assign category material: " + str(e))
        for material in materials:
            if (
                excludeGlazing and 
                material.MaterialClass == "Glass"
            ):
                # print("Exclude:", material.Name, material.MaterialClass)
                continue
            if (
                excludeTestFitFloors and
                "00 00 00 Test Fit" in material.Name
            ):
                continue
            if material.Id in ignoreMaterials:
                continue
            material.UseRenderAppearanceForShading = False
            material.AppearanceAssetId = swapAssetId
            # print("Include:", material.Name, material.MaterialClass)