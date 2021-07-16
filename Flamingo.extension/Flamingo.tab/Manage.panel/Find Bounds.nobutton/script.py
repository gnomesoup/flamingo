from pyrevit import DB, HOST_APP, revit, clr
import System
clr.ImportExtensions(System.Linq)

def findBounds(doc):
    pickedPoint = revit.pick_point("Click center of space")
    referencePoint = DB.XYZ(0, 0, 2.0/12.0) + pickedPoint
    print("Picked", pickedPoint)
    print("reference", referencePoint)
    viewFamilyType = DB.FilteredElementCollector(doc) \
        .OfClass(DB.ViewFamilyType) \
        .FirstOrDefault(
            lambda x: x.ViewFamily == DB.ViewFamily.ThreeDimensional
        )
    with revit.DryTransaction("Find bounds") as t:
        view3D = DB.View3D.CreateIsometric(doc, viewFamilyType.Id)
        referenceIntersector = DB.ReferenceIntersector(
            view3D
        )
        result = referenceIntersector.FindNearest(
            referencePoint, DB.XYZ(1, 0, 0)
        )
        revit.delete.delete_elements([view3D])
    print(result)

if __name__ == "__main__":
    doc = HOST_APP.doc
    findBounds(doc)