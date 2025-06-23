from Autodesk.Revit import DB
from pyrevit import forms
from pyrevit import HOST_APP, revit, script

LOGGER = script.get_logger()
OUTPUT = script.get_output()

if __name__ == "__main__":
    doc = HOST_APP.doc

    activeView = doc.ActiveView

    rooms = (
        DB.FilteredElementCollector(doc, activeView.Id)
        .OfCategory(DB.BuiltInCategory.OST_Rooms)
        .ToElements()
    )
    roomsByLink = {}
    if rooms:
        roomsByLink["host"] = {"rooms": rooms, "transform": None}

    links = (
        DB.FilteredElementCollector(doc, activeView.Id)
        .OfCategory(DB.BuiltInCategory.OST_RvtLinks)
        .ToElements()
    )

    for link in links:
        linkBoundingBox = activeView.CropBox
        linkBoundingBox.Transform = link.GetTotalTransform().Inverse
        outline = DB.Outline(linkBoundingBox.Min, linkBoundingBox.Max)
        linkDoc = link.GetLinkDocument()
        linkRooms = (
            DB.FilteredElementCollector(linkDoc).OfCategory(
                DB.BuiltInCategory.OST_Rooms
            )
            # .WherePasses(DB.BoundingBoxIntersectsFilter(outline))
            .ToElements()
        )
        if linkRooms:
            roomsByLink[link.Id] = {
                "rooms": linkRooms,
                "transform": link.GetTotalTransform(),
            }

    if not roomsByLink:
        forms.alert(
            "No rooms found in the active view. Switch to a view with visible rooms "
            "and try again",
            exitscript=True,
        )
    tags = (
        DB.FilteredElementCollector(doc, activeView.Id)
        .OfCategory(DB.BuiltInCategory.OST_RoomTags)
        .ToElements()
    )
    taggedRoomIds = [tag.TaggedRoomId for tag in tags if tag.TaggedRoomId is not None]
    viewDirection = activeView.ViewDirection
    indexes = []
    n = (0, 1, 2)
    for i in n:
        if round(viewDirection[i], 3) == 0.000:
            indexes.append(i)
    if len(indexes) != 2:
        forms.alert(
            "Tagging rooms is not supported in this view type.",
            exitscript=True,
        )
    with revit.Transaction("Tag All Rooms"):
        for linkId, linkData in roomsByLink.items():
            LOGGER.debug("Processing link: {}".format(linkId))
            for room in linkData["rooms"]:
                try:
                    if linkId == "host":
                        linkElementId = DB.LinkElementId(room.Id)
                    else:
                        linkElementId = DB.LinkElementId(linkId, room.Id)
                    if linkElementId in taggedRoomIds:
                        LOGGER.debug(
                            "Room already tagged: {}".format(OUTPUT.linkify(room.Id))
                        )
                        continue
                    point = room.Location.Point
                    if linkData["transform"]:
                        LOGGER.debug("Transforming point")
                        point = linkData["transform"].OfPoint(point)
                    LOGGER.debug("Point: {}".format(point))
                    LOGGER.debug("Tagging room: {}".format(OUTPUT.linkify(room.Id)))
                    tag = doc.Create.NewRoomTag(
                        linkElementId,
                        DB.UV(point[indexes[0]], point[indexes[1]]),
                        activeView.Id,
                    )
                except Exception as e:
                    LOGGER.debug(
                        "Failed to tag room: {} ({})".format(room.Id, e),
                    )
