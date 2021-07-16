import clr
from pyrevit import HOST_APP, forms, script, DB, revit
import re
import xml.etree.ElementTree as ET

clr.AddReference("System")
from System.Collections.Generic import List

doc = HOST_APP.doc

landXMLPath = forms.pick_file(file_ext="xml", title="Select a LandXML file")
if landXMLPath is None:
    script.exit()
landXML = ET.parse(landXMLPath).getroot()
m = re.match(r"\{.*\}", landXML.tag)
xmlns = m.group(0) if m else ''
surfaces = landXML.find(xmlns + 'Surfaces')
names = [
    surface.attrib['name'] for surface in surfaces.findall(xmlns + 'Surface')
]

surfaceName = forms.ask_for_one_item(items=names,
    default=names[0],
    title="Select a surface to make a topography")

surface = surfaces.find("./" + xmlns + "Surface[@name=\"" + surfaceName + "\"]")
points = []
try:
    print("Collectiong Contour Points")
    contours = surface.find(xmlns + "SourceData")\
        .find(xmlns + "Contours")\
        .findall(xmlns + "Contour")
    for contour in contours:
        z = contour.attrib['elev']
        for pointList in contour.findall(xmlns + "PntList2D"):
            firstPoint = True
            for i in pointList.text.split(" "):
                if firstPoint == True:
                    y = i
                    firstPoint = False
                else:
                    points.append((float(i), float(y), float(z)))
                    firstPoint = True
except:
    print("No contours found.")
try:
    print("Collecting tin points")
    pnts = surface.find(xmlns + "Definition")\
        .find(xmlns + "Pnts")\
        .findall(xmlns + "P")
    for pnt in pnts:
        pntList = pnt.text.split(" ")
        points.append((float(pntList[1]), float(pntList[0]), float(pntList[2])))

except Exception as e:
    print("No tin points found.")

projectLocation = doc.ActiveProjectLocation
print("Current shared coordatine: " + projectLocation.Name)
projectPosition = projectLocation.GetProjectPosition(DB.XYZ(0, 0, 0))
sharedPoint = DB.XYZ(projectPosition.EastWest,
                     projectPosition.NorthSouth,
                     projectPosition.Elevation)
print("Shared Coordinate location: " +  str(sharedPoint))
topoOrigin = DB.XYZ(0, 0, 0).Subtract(sharedPoint)
cleanedPoints = List[DB.XYZ]()

print("Cleaning points")
pointDictionary = {}
for point in points:
    pIndex = (round(point[0], 1), round(point[1], 1))
    tPoint = DB.XYZ(point[0], point[1], point[2]).Add(topoOrigin) 
    if pIndex in pointDictionary:
        dPoint = pointDictionary[pIndex]
        if dPoint.Z < tPoint.Z:
            pointDictionary[pIndex] = tPoint
    else:
        pointDictionary[pIndex] = tPoint

for key, point in pointDictionary.items():
    cleanedPoints.Add(point)

print("Creating Topography with " + str(len(cleanedPoints)) + " points.")

with revit.Transaction("Create Topography"):
    topo = DB.Architecture.TopographySurface.Create(doc, cleanedPoints)

print("Topo Created")