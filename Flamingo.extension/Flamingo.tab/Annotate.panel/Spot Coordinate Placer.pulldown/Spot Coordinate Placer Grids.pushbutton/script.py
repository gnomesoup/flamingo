from pyrevit import forms, revit, DB, script, clr, HOST_APP
from flamingo.geometry import GetMinMaxPoints
from flamingo.geometry import GetMidPointIntersections 
from flamingo.forms import AskForLength
from math import pi
import uuid

doc = HOST_APP.doc

# Set the elements to work with
currentView = doc.ActiveView
if revit.get_selection().is_empty:
    selection = revit.query.get_all_elements_in_view(
        currentView
    )
else:
    selection = revit.get_selection().elements

gridCategory = DB.Category.GetCategory(
    doc, DB.BuiltInCategory.OST_Grids
)

grids = [
    element for element in selection
    if element.Category is not None and 
    element.Category.Id == gridCategory.Id
]

commandOptions = [
    "Place at grid intersections",
    "Place at coordinate interval",
    # "Place on equal spacing",
    "More info",
]

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
            "grids in the current view.\n\n" +
            "Place at grid intersection:\nA spot coordinate is placed at " +
            "the intersections of each grid in the view or selection.\n\n" +
            "Place at coordinate interval:\nA spot coordinate is placed at " +
            "each point where the wall intersects with planes along the " +
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
    
if coordinateType in [
    "Place at coordinate interval",
    "Place on equal spacing",
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
viewDirection = currentView.ViewDirection
viewScale = currentView.Scale
leaderScaledLength = viewScale * (.25 / 12)
viewRightDirection = currentView.RightDirection
viewUpDirection = currentView.UpDirection
rotation = DB.Transform.CreateRotation(DB.XYZ.BasisZ, 45 * pi / 180)
leaderAngleVector = rotation.OfVector(viewRightDirection * leaderScaledLength)
leaderStraightVector = viewRightDirection * (leaderScaledLength * 2)

# find out if there are already spot coordinates placed so we don't double up.
currentSpots = DB.FilteredElementCollector(doc, currentView.Id)\
    .OfCategory(DB.BuiltInCategory.OST_SpotCoordinates)\
    .WhereElementIsNotElementType()

# Add their origin points to a dictionary of points
allPoints = {}
collectedPoints = []
for spot in currentSpots:
    key = str(uuid.uuid1())
    allPoints[key] = {"reference": None, "point": spot.Origin, "primary": True}
    collectedPoints.append(key)

# get intersections of grids with other grids
placementThresholdDistance = 0.5
if coordinateType == "Place at grid intersections":
    placementThresholdDistance = 0.125 / 12.0
    curves = []
    for grid in grids:
        curves.append(grid.Curve)
    for grid in grids:
        for curve in curves:
            intersectInfo = clr.StrongBox[DB.IntersectionResultArray]()
            if (str(grid.Curve.Intersect(curve, intersectInfo))) == "Overlap":
                for i in range(intersectInfo.Size):
                    intersectionPoint = intersectInfo.get_Item(i).XYZPoint
                    allPoints[str(uuid.uuid1())] = {
                        "reference": grid,
                        "point": intersectionPoint,
                        "primary": False
                    }

elif coordinateType in [
    "Place at coordinate interval",
]:
    outline = DB.Outline(DB.XYZ.Zero, DB.XYZ.Zero)
    itemsForIntersect = []
    for grid in grids:
        bbPoints = GetMinMaxPoints(grid, currentView)
        outline.AddPoint(bbPoints[0])
        outline.AddPoint(bbPoints[1])
        try:
            itemsForIntersect.append(grid)
        except Exception as e:
            print("Get grid curve: " + str(e))
            continue

# add points where grids intersect intervals
if coordinateType in [
    "Place at coordinate interval",
]:
    surveyPoints = DB.FilteredElementCollector(doc)\
        .OfCategory(DB.BuiltInCategory.OST_SharedBasePoint)\
        .ToElements()
    surveyPoint = surveyPoints[0].Position
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
                if value['point'].DistanceTo(allPoints[pointId]['point']) < \
                    placementThresholdDistance:
                    pointMatch = True
                    continue
            if not pointMatch:
                collectedPoints.append(key)
    
for key, value in allPoints.items():
    if value['primary'] == False:
        pointMatch = False
        for pointId in collectedPoints:
            if (
                abs(value['point'].X - allPoints[pointId]['point'].X) < \
                    placementThresholdDistance
                and abs(value['point'].Y - allPoints[pointId]['point'].Y) < \
                    placementThresholdDistance
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
            reference = DB.Reference(referenceElement)
        # If the object doesn't have a valid reference add a
        # generic annotation family to use as a reference
        except Exception as e:
            placeholder = doc.Create.NewFamilyInstance(
                point,
                spotPlaceholder,
                currentView
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
            spot = doc.Create.NewSpotCoordinate(currentView,
                        reference,
                        point,
                        bendPoint,
                        endPoint,
                        point,
                        config.leader)
        except Exception as e:
            print("Make spot coordinate: " + str(e))
