from Autodesk.Revit import DB
from pyrevit import HOST_APP, revit, forms, script
from math import ceil
from string import ascii_uppercase


def GetSheetWorkArea(titleBlock, marginleft, marginright, margintop, marginbottom):
    width = titleBlock.get_Parameter(DB.BuiltInParameter.SHEET_WIDTH).AsDouble()
    height = titleBlock.get_Parameter(DB.BuiltInParameter.SHEET_HEIGHT).AsDouble()
    origin = titleBlock.Location.Point
    workMinX = origin.X + marginleft
    workMinY = origin.Y + marginbottom
    workMaxX = origin.X + width - marginright
    workMaxY = origin.Y + height - margintop
    return {"MinimumPoint": (workMinX, workMinY), "MaximumPoint": (workMaxX, workMaxY)}


def inBoundingBox(boundingBox, point):
    return (
        boundingBox[0][0] < point.X < boundingBox[1][0]
        and boundingBox[0][1] < point.Y < boundingBox[1][1]
    )


if __name__ == "__main__":
    doc = HOST_APP.doc

    if doc.IsFamilyDocument:
        if not doc.OwnerFamily.FamilyCategory == DB.Category.GetCategory(
            doc, DB.BuiltInCategory.OST_TitleBlocks
        ):
            forms.alert(
                "This utility is for renaming views on a sheet, please try "
                "again on a project sheet. You can also run this in a title"
                "block family to set the default parameters for margin and cell"
                " size."
            )
            script.exit()
        forms.inform_wip()

    currentView = doc.ActiveView
    titleBlocks = (
        DB.FilteredElementCollector(doc, currentView.Id)
        .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
        .ToElements()
    )

    if len(titleBlocks) > 1:
        forms.alert(
            "Found more than one title block on the sheet.\n\n"
            + "Remove extra titleblocks and try again",
            title=scriptTitle,
            exitscript=True,
        )

    marginLeft = 1.5
    marginRight = 4.5
    marginTop = 0.5
    marginBottom = 0.5

    titleBlock = titleBlocks[0]
    workArea = GetSheetWorkArea(
        titleBlock,
        marginLeft / 12.0,
        marginRight / 12.0,
        marginTop / 12.0,
        marginBottom / 12.0,
    )
    width = workArea["MaximumPoint"][0] - workArea["MinimumPoint"][0]
    height = workArea["MaximumPoint"][1] - workArea["MinimumPoint"][1]
    divisionsWidth = ceil(width / 0.5)
    spacingWidth = width / float(divisionsWidth)
    divisionsHeight = ceil(height / 0.5)
    spacingHeight = height / float(divisionsHeight)
    boundingBoxes = {}
    for i in range(int(divisionsWidth)):
        xIndex = divisionsWidth - i
        xMax = spacingWidth * float(xIndex) + workArea["MinimumPoint"][0]
        xMin = xMax - spacingWidth
        for j in range(int(divisionsHeight)):
            yMin = spacingHeight * float(j) + workArea["MinimumPoint"][1]
            yMax = yMin + spacingHeight
            yLetter = ascii_uppercase[j]
            detailNumber = yLetter + str(i + 1)
            boundingBoxes[detailNumber] = ((xMin, yMin), (xMax, yMax))

    viewportIds = currentView.GetAllViewports()
    viewports = {}
    for viewportId in viewportIds:
        viewport = doc.GetElement(viewportId)
        if viewport.Name != "0 No Title":
            parameter = viewport.get_Parameter(
                DB.BuiltInParameter.VIEWPORT_DETAIL_NUMBER
            )
            labelOutline = viewport.GetLabelOutline()
            labelMinPoint = labelOutline.MinimumPoint + DB.XYZ(
                0.25 / 12.0, 0.25 / 12.0, 0
            )
            viewports[viewportId] = {
                "point": labelMinPoint,
                "parameter": parameter,
                "oldLabel": parameter.AsString(),
            }

    if not viewports:
        if not viewportIds:
            message = "There are no views on this sheet to renumber"
        else:
            message = (
                "There are no views on this sheet with a title shown.\n\n"
                + "Swap a view's viewport type to something other than "
                + '"0 No Title" and try again.'
            )
        forms.alert(message, title=scriptTitle, exitscript=True)

    labels = {}
    if len(viewports) == 1:
        labels["1"] = [viewports.keys()[0]]
    else:
        for viewportId, data in viewports.items():
            for label, boundingBox in boundingBoxes.items():
                if inBoundingBox(boundingBox, data["point"]):
                    if labels.has_key(label):
                        labels[label].append(viewportId)
                    else:
                        labels[label] = [viewportId]

    for label, viewportIds in labels.items():
        for i in range(len(viewportIds)):
            if i == 0 and len(viewportIds) == 1:
                newLabel = label
            else:
                newLabel = label + ascii_uppercase[i]
            viewports[viewportIds[i]]["newLabel"] = newLabel

    with revit.Transaction("Renumber Views"):
        for viewportId, data in viewports.items():
            data["parameter"].Set(data["oldLabel"] + "^")
        outOfBounds = 0
        for viewportId, data in viewports.items():
            if data.has_key("newLabel"):
                data["parameter"].Set(data["newLabel"])
            else:
                outOfBounds = outOfBounds + 1
                data["parameter"].Set("-" + str(outOfBounds))

    if outOfBounds > 0:
        plural = " was"
        if outOfBounds > 1:
            plural = "s where"
        forms.alert(
            str(outOfBounds)
            + " view"
            + plural
            + " out of the working space of a standard titleblock."
            + " Adjust the view labels and re-run."
        )
