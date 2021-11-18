from Autodesk.Revit import DB
from pyrevit import revit, HOST_APP, forms, script
from string import ascii_uppercase
from flamingo.geometry import GetMinMaxPoints, GetMidPoint

if __name__ == "__main__":
    doc = HOST_APP.doc
    activeView = doc.ActiveView
    activePhaseParameter = activeView.get_Parameter(DB.BuiltInParameter.VIEW_PHASE)
    activePhaseId = activePhaseParameter.AsElementId()
    activePhase = doc.GetElement(activePhaseId)

    selection = revit.get_selection()
    if selection.is_empty:
        activeDesignOptionId = DB.DesignOption.GetActiveDesignOptionId(doc)
        elementDesignOptionFilter = DB.ElementDesignOptionFilter(
            activeDesignOptionId
        )
        doors = DB.FilteredElementCollector(doc, activeView.Id)\
            .OfCategory(DB.BuiltInCategory.OST_Doors)\
            .WhereElementIsNotElementType()\
            .WherePasses(elementDesignOptionFilter)\
            .ToElements()
        response = forms.alert(
            msg="Are you sure you would like to renumber all {} active door{} "
                "in this view?\n\nTo limit the number of doors numbered, first "
                "select the elements and run this command again.".format(
                    len(doors), "" if len(doors) == 1 else "s"
                ),
            yes=True,
            cancel=True,
        )
        if not response:
            script.exit()
    else:
        doorCategory = DB.Category.GetCategory(
            doc,
            DB.BuiltInCategory.OST_Doors
        )
        doors = [
            element for element in selection.elements
            if element.Category.Id.IntegerValue == int(
                DB.BuiltInCategory.OST_Doors
            )
        ]
    doorCount = len(doors)

    # Make a dictionary of rooms with door properties
    roomDoors = {}
    for door in doors:
        if not door:
            continue

        toRoom = (door.ToRoom)[activePhase]
        if toRoom:
            toRoomId = toRoom
            if toRoom.Id in roomDoors:
                roomDoors[toRoom.Id]["doors"] = \
                    roomDoors[toRoom.Id]["doors"] + [door.Id]
            else:
                roomDoors[toRoom.Id] = {"doors": [door.Id]}
                roomDoors[toRoom.Id]["roomNumber"] = toRoom.Number
                point1, point2 = GetMinMaxPoints(toRoom, activeView)
                roomCenter = GetMidPoint(point1, point2)
                roomDoors[toRoom.Id]["roomLocation"] = roomCenter

        fromRoom = (door.FromRoom)[activePhase]
        if fromRoom:
            if fromRoom.Id in roomDoors:
                roomDoors[fromRoom.Id]["doors"] =\
                    roomDoors[fromRoom.Id]["doors"] + [door.Id]
            else:
                roomDoors[fromRoom.Id] = {"doors": [door.Id]}
                roomDoors[fromRoom.Id]["roomNumber"] = fromRoom.Number
                point1, point2 = GetMinMaxPoints(fromRoom, activeView)
                roomCenter = GetMidPoint(point1, point2)
                roomDoors[fromRoom.Id]["roomLocation"] = roomCenter

    maxLength = max([len(values["doors"]) for values in roomDoors.values()])
    for key, values in roomDoors.items():
        roomDoors[key]["doorCount"] = len(values["doors"])
    
    # # print dictionary
    # for roomId, values in roomDoors.items():
    #     print("Room: {}".format(roomId.IntegerValue))
    #     for key, value in values.items():
    #         if type(value) is list:
    #             print("\t{}:".format(key))
    #             for subValue in value:
    #                 print("\t\t{}".format(subValue.IntegerValue))
    #         else:
    #             print("\t{}: {}".format(key, value))

    # Number doors
    numberedDoors = set()
    with revit.Transaction("Renumber Doors"):
        i = 1
        n = 0
        nMax = doorCount * 2

        # iterate through counts of doors connected to each room
        while i < maxLength:
            # avoid a loop by stoping at double the count of doors
            if n > nMax:
                break
            noRooms = True
            # Go through each room and look for door counts that match
            # current level
            for key, values in roomDoors.items():
                if values["doorCount"] == i:
                    noRooms = False
                    doorIds = values["doors"]
                    doorsToNumber = [
                        doorId for doorId in doorIds
                        if doorId not in numberedDoors
                    ]
                    # Go through all the doors connected to the room
                    # and give them a number
                    for j, doorId in enumerate(doorsToNumber):
                        numberedDoors.add(doorId)
                        door = doc.GetElement(doorId)
                        if len(doorsToNumber) > 1:
                            # point1, point2 = GetMinMaxPoints(door, activeView)
                            # roomCenter = values["roomLocation"]
                            # doorCenter = GetMidPoint(point1, point2)
                            # doorVector = doorCenter.Subtract(roomCenter)
                            # angle = (DB.XYZ.BasisY).AngleTo(
                            #     doorVector.Normalize()
                            # )
                            # dotProduct = (DB.XYZ.BasisY).DotProduct(
                            #     doorVector.Normalize()
                            # )
                            mark = "{}{}".format(
                                values["roomNumber"], ascii_uppercase[j]
                            )
                            # print("{}: {} ({})".format(mark, angle, dotProduct))
                        else:
                            mark = values["roomNumber"]
                        markParameter = door.get_Parameter(
                            DB.BuiltInParameter.ALL_MODEL_MARK
                        )
                        markParameter.Set(mark)
            for roomId, values in roomDoors.items():
                unNumberedDoors = filter(None, [
                    doorId for doorId in values["doors"]
                    if doorId not in numberedDoors
                ])
                roomDoors[roomId]["doors"] = unNumberedDoors
                roomDoors[roomId]["doorCount"] = len(unNumberedDoors)
            if noRooms:
                i += 1
            n += 1

    # print("Doors Renumbered: {}".format(doorCount))