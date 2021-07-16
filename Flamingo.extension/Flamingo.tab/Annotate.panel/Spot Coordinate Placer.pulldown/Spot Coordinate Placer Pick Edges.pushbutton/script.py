from pyrevit import DB, revit, UI, forms, script, clr, HOST_APP
from flamingo.geometry import GetMinMaxPoints
from flamingo.geometry import GetMidPointIntersections
from flamingo.forms import AskForLength
from math import pi
import uuid

# clr.AddReference("System")
# from System.Collections.Generic import List
import System
clr.ImportExtensions(System.Linq)

doc = HOST_APP.doc
app = HOST_APP.doc
selection = HOST_APP.uidoc.Selection

geometryOptions = DB.Options()
geometryOptions.ComputeReferences = True
geometryOptions.IncludeNonVisibleObjects = True
geometryOptions.View = doc.ActiveView

activeView = doc.ActiveView
edges = selection.PickObjects(UI.Selection.ObjectType.Edge)

# Create a list of options for a user to pick from
commandOptions = [
    "Place at endpoints",
    "Place at coordinate interval",
    "Place at endpoints and interval",
    # "Place on equal spacing",
    "More info",
]

# Check to see if the user has previously selected a leader choice and set
# a default if not
config = script.get_config()
if not hasattr(config, 'leader'):
    config.leader = True
commandSwitches = {'Add leader':config.leader}

coordinateType = ""
while not coordinateType:
    # Ask user for how spot coordinates should be placed
    options = forms.CommandSwitchWindow.show(
        commandOptions,
    switches=commandSwitches,
        message="Spot Coordinate Placer Options:"
    )
    leaderState = options[1]['Add leader']
    if options[0] is None:
        script.exit()
    elif options[0] == "More info":
        forms.alert(
            "Choose how you would like spot coordinates to be placed on the " +
            "selected edges.\n\n" +
            "Place at endpoints:\nA spot coordinate is placed at the " +
            "endpoint of each edge in the selection\n\n" +
            "Place at coordinate interval:\nA spot coordinate is placed at " +
            "each point where the edge intersects with planes along the " +
            "x, y coordinate system of the model. You will be able to choose "
            "the interval spacing in the next menu. The center of the " + 
            "coordinate system is the project survey point.\n\n"
            "Add leader:\nIf turned on (slider to the right), a leader will " +
            "be included when the spot coordinate family is placed."
        )
    else:
        coordinateType = options[0]
    if leaderState != config.leader:
        config.leader = leaderState
        script.save_config()

# if we are placing on an interval ask for interval
if coordinateType in [
    "Place at coordinate interval",
    "Place at endpoints and interval"
]:
    config = script.get_config()
    if not hasattr(config, 'spacing'):
        config.spacing = 2.0
    spotSpacing = AskForLength(config.spacing)
    if spotSpacing is None:
        script.exit()
    elif spotSpacing != config.spacing:
        config.spacing = spotSpacing
        script.save_config()

# collect info on current view
viewDirection = activeView.ViewDirection
viewScale = activeView.Scale
leaderScaledLength = viewScale * (.25 / 12)
viewRightDirection = activeView.RightDirection
viewUpDirection = activeView.UpDirection
rotation = DB.Transform.CreateRotation(DB.XYZ.BasisZ, 45 * pi / 180)
leaderAngleVector = rotation.OfVector(viewRightDirection * leaderScaledLength)
leaderStraightVector = viewRightDirection * (leaderScaledLength * 2)

# find out if there are already spot coordinates placed so we don't double up.
currentSpots = DB.FilteredElementCollector(doc, activeView.Id)\
    .OfCategory(DB.BuiltInCategory.OST_SpotCoordinates)\
    .WhereElementIsNotElementType()
# Add their origin points to a dictionary of points
allPoints = {}
collectedPoints = []
for spot in currentSpots:
    key = str(uuid.uuid1())
    allPoints[key] = {"reference": None, "point": spot.Origin, "primary": True}
    collectedPoints.append(key)


if coordinateType in [
    "Place at endpoints",
    "Place at endpoints and interval"
]:
    for edge in edges:
        curve = doc.GetElement(edge)\
            .GetGeometryObjectFromReference(edge).AsCurve()
        for i in [0,1]:
            endPoint = curve.GetEndPoint(i)
            allPoints[str(uuid.uuid1())] = {
                "reference": edge,
                "point": endPoint,
                "primary": True
            }

if coordinateType in [
    "Place at coordinate interval",
    "Place at endpoints and interval"
]:
    surveyPoint = DB.FilteredElementCollector(doc)\
        .OfCategory(DB.BuiltInCategory.OST_SharedBasePoint)\
        .ToElements().First().Position
    outline = DB.Outline(surveyPoint, surveyPoint)
    itemsForIntersect = []
    for edge in edges:
        # get min & max points of all elements
        hostElement = doc.GetElement(edge)
        bbPoints = GetMinMaxPoints(hostElement, activeView)
        outline.AddPoint(bbPoints[0])
        outline.AddPoint(bbPoints[1])
        try:
            itemsForIntersect.append(edge)
        except Exception as e:
            print("Get edge error: " + str(e))
            continue
    allPoints.update(GetMidPointIntersections(
        doc,
        surveyPoint,
        DB.XYZ.BasisY,
        outline,
        spotSpacing,
        itemsForIntersect
    ))
    allPoints.update(GetMidPointIntersections(
        doc,
        surveyPoint,
        DB.XYZ.BasisX,
        outline,
        spotSpacing,
        itemsForIntersect
    ))

# Check that points are too close to eachother
for key, value in allPoints.items():
    if value['primary'] == True:
        if collectedPoints is None or value['reference'] is None:
            collectedPoints.append(key)
        else:
            pointMatch = False
            for pointId in collectedPoints:
                if value['point'].DistanceTo(allPoints[pointId]['point']) < 0.5:
                    pointMatch = True
                    continue
            if not pointMatch:
                collectedPoints.append(key)

for key, value in allPoints.items():
    if value['primary'] == False:
        pointMatch = False
        for pointId in collectedPoints:
            if (
                abs(value['point'].X - allPoints[pointId]['point'].X) < 0.5
                and abs(value['point'].Y - allPoints[pointId]['point'].Y) < 0.5
            ):
                pointMatch = True
                continue
        if not pointMatch:
            collectedPoints.append(key)


# Warn about excessive points
if len(collectedPoints) > 1000:
    if not forms.alert("You are about to place " + str(len(collectedPoints)) 
                       + " spot coordinates. Care to continue? "
                       + "This could take a bit."):
        script.exit()

# Actually place spot coordinates
with revit.Transaction('Place spot coordinates'):
    for pointId in collectedPoints:
        point = allPoints[pointId]['point']
        referenceElement = allPoints[pointId]['reference']
        if referenceElement is None:
            continue
        try:
            if type(referenceElement) is DB.Reference:
                reference = referenceElement
            else:
                reference = DB.Reference(referenceElement)
        # If the object doesn't have a valid reference add a
        # generic annotation family to use as a reference
        except Exception as e:
            placeholder = doc.Create.NewFamilyInstance(
                point,
                spotPlaceholder,
                activeView
            )
            references = placeholder.GetReferences(
                DB.FamilyInstanceReferenceType.WeakReference
            )
            reference = references.First()
            point = placeholder.Location.Point
        # Place the spot coordinate
        try:
            bendPoint = point.Add(leaderAngleVector)
            endPoint = bendPoint.Add(leaderStraightVector)
            spot = doc.Create.NewSpotCoordinate(activeView,
                        reference,
                        point,
                        bendPoint,
                        endPoint,
                        point,
                        config.leader)
        except Exception as e:
            print("Make spot coordinate: " + str(e))
