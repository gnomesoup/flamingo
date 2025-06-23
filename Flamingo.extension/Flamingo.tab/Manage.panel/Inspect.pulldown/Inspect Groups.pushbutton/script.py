from Autodesk.Revit import DB
from pyrevit import HOST_APP, script
from pyrevit.revit import query
from datetime import datetime

LOGGER = script.get_logger()
OUTPUT = script.get_output()

if __name__ == "__main__":
    doc = HOST_APP.doc

    OUTPUT.print_md("# Group Report")
    OUTPUT.print_md(doc.Title)
    OUTPUT.print_md(datetime.now().isoformat())
    groupTypes = (
        DB.FilteredElementCollector(doc)
        .OfClass(DB.GroupType)
        .WhereElementIsElementType()
        .ToElements()
    )
    levels = DB.FilteredElementCollector(doc).OfClass(DB.Level)
    levelNamesByElevation = [
        query.get_name(level) for level in sorted(levels, key=lambda x: x.Elevation)
    ]
    levels = {query.get_name(level): level.Id for level in levels}
    OUTPUT.print_md("Total group types: {}".format(len(groupTypes)))

    modelGroupTypes = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_IOSModelGroups)
        .WhereElementIsElementType()
        .ToElements()
    )
    modelGroupTypes = sorted(modelGroupTypes, key=lambda x: query.get_name(x))
    groupReport = {"Model Groups": [], "Detail Groups": []}
    for modelGroupType in modelGroupTypes:
        typeName = query.get_name(modelGroupType)
        groups = list(modelGroupType.Groups)
        levelData = {}
        for level in levelNamesByElevation:
            groupsOnLevel = [
                group for group in groups if group.LevelId == levels[level]
            ]
            levelData[level] = len(groupsOnLevel)
        groupReport["Model Groups"].append(
            {
                "name": typeName,
                "placed": len(groups),
                "levels": levelData,
            }
        )

    detailGroupTypes = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_IOSDetailGroups)
        .WhereElementIsElementType()
        .ToElements()
    )
    detailGroupTypes = sorted(detailGroupTypes, key=lambda x: query.get_name(x))
    for detailGroupType in detailGroupTypes:
        typeName = query.get_name(detailGroupType)
        groups = list(detailGroupType.Groups)
        viewData = {}
        for group in groups:
            viewName = doc.GetElement(group.OwnerViewId).Title
            if not viewName:
                continue
            if viewName in viewData:
                viewData[viewName] += 1
            else:
                viewData[viewName] = 1
        groupReport["Detail Groups"].append(
            {"name": typeName, "placed": len(groups), "views": viewData}
        )

    groupCategoryList = ["Model Groups", "Detail Groups"]
    for groupCategory in groupCategoryList:
        groupTypes = groupReport[groupCategory]
        OUTPUT.print_md("## {}".format(groupCategory))
        n = 1
        OUTPUT.print_md("Total model groups: {}".format(len(groupTypes)))
        OUTPUT.print_md("Unplaced Groups (:cross_mark:): {}".format(len(
            [groupType for groupType in groupTypes if groupType["placed"] < 1]
        )))
        OUTPUT.print_md("Groups only placed once (:warning:): {}".format(len(
            [groupType for groupType in groupTypes if groupType["placed"] == 1]
        )))

        for groupType in groupTypes:
            if groupType["placed"] < 1:
                prefix = ":cross_mark:"
            elif groupType["placed"] == 1:
                prefix = ":warning:"
            else:
                prefix = ":white_heavy_check_mark:"
            OUTPUT.print_md("{}\. **{}** {}".format(n, groupType["name"], prefix))
            n = n + 1
            OUTPUT.print_md("- Placed: {}".format(groupType["placed"]))
            if "levels" in groupType:
                levelData = groupType["levels"]
                for level in levelNamesByElevation:
                    levelCount = groupType["levels"][level]
                    if levelCount > 0:
                        OUTPUT.print_md("- {}: {}".format(level, levelCount))
            elif "views" in groupType:
                viewData = groupType["views"]
                keys = sorted(viewData.keys())
                for key in keys:
                    OUTPUT.print_md("- {}: {}".format(key, viewData[key]))
