from re import I
from Autodesk.Revit import DB
from pyrevit import HOST_APP, forms
from pyrevit.revit import ensure, Transaction

if __name__ == "__main__":
    doc = HOST_APP.doc
    forms.check_familydoc(doc, exitscript=True)

    #Get all reference planes
    planes = DB.FilteredElementCollector(doc)\
        .OfClass(DB.ReferencePlane).ToElements()
    
    if planes:
        planes = list(planes)
        #Get origin reference planes
        parameter = planes[0].get_Parameter(
            DB.BuiltInParameter.DATUM_PLANE_DEFINES_ORIGIN
        )
        originPlanes = DB.FilteredElementCollector(doc)\
            .OfClass(DB.ReferencePlane)\
            .WherePasses(
                DB.ElementParameterFilter(
                    DB.FilterIntegerRule(
                        DB.ParameterValueProvider(
                            parameter.Id
                        ),
                        DB.FilterNumericEquals(),
                        1
                    )
                )
            )
        x = 0
        y = 0
        z = 0
        for plane in originPlanes:
            normal = plane.Normal
            if(
                normal.IsAlmostEqualTo(DB.XYZ.BasisX) or
                normal.IsAlmostEqualTo(DB.XYZ.BasisX.Negate())
            ):
                x = plane.BubbleEnd.X
            elif(
                normal.IsAlmostEqualTo(DB.XYZ.BasisY) or
                normal.IsAlmostEqualTo(DB.XYZ.BasisY.Negate())
            ):
                y = plane.BubbleEnd.Y
            elif(
                normal.IsAlmostEqualTo(DB.XYZ.BasisZ) or
                normal.IsAlmostEqualTo(DB.XYZ.BasisZ.Negate())
            ):
                z = plane.BubbleEnd.Z
        origin = DB.XYZ(x, y, z)
    else:
        origin = DB.XYZ.Zero
    
    print("origin = {}, {}, {}".format(origin.X, origin.Y, origin.Z))

    #Get list of reference planes that need to be created 
    LENGTH = 5
    OFFSET = 1.5
    refView = doc.ActiveView
    frontView = refView
    referenceTypes = {
        int(DB.FamilyInstanceReferenceType.Front): {
            "name": "Front",
            "bubbleEnd": DB.XYZ(-LENGTH, -OFFSET, OFFSET-LENGTH),
            "freeEnd": DB.XYZ(LENGTH, -OFFSET, OFFSET-LENGTH),
            "cutVector": DB.XYZ.BasisZ,
            "view": refView
        },
        int(DB.FamilyInstanceReferenceType.Back): {
            "name": "Back",
            "bubbleEnd": DB.XYZ(-LENGTH, OFFSET, OFFSET-LENGTH),
            "freeEnd": DB.XYZ(LENGTH, OFFSET, OFFSET-LENGTH),
            "cutVector": DB.XYZ.BasisZ,
            "view": refView
        },
        int(DB.FamilyInstanceReferenceType.Left): {
            "name": "Left",
            "bubbleEnd": DB.XYZ(-OFFSET, LENGTH, OFFSET-LENGTH),
            "freeEnd": DB.XYZ(-OFFSET, -LENGTH, OFFSET-LENGTH),
            "cutVector": DB.XYZ(0, 0, OFFSET * 2),
            "view": refView
        },
        int(DB.FamilyInstanceReferenceType.Right): {
            "name": "Right",
            "bubbleEnd": DB.XYZ(OFFSET, LENGTH, OFFSET-LENGTH),
            "freeEnd": DB.XYZ(OFFSET, -LENGTH, OFFSET-LENGTH),
            "cutVector": DB.XYZ.BasisZ,
            "view": refView
        },
        int(DB.FamilyInstanceReferenceType.Top): {
            "name": "Top",
            "bubbleEnd": DB.XYZ(-LENGTH, -LENGTH, OFFSET * 2),
            "freeEnd": DB.XYZ(LENGTH, -LENGTH, OFFSET * 2),
            "cutVector": DB.XYZ.BasisY,
            "view": frontView
        },
        int(DB.FamilyInstanceReferenceType.Bottom): {
            "name": "Bottom",
            "bubbleEnd": DB.XYZ(-LENGTH, -LENGTH, 0),
            "freeEnd": DB.XYZ(LENGTH, -LENGTH, 0),
            "cutVector": DB.XYZ.BasisY,
            "view": frontView
        },
        int(DB.FamilyInstanceReferenceType.CenterFrontBack): {
            "name": "Center (Front/Back)",
            "bubbleEnd": DB.XYZ(-LENGTH, 0, OFFSET-LENGTH),
            "freeEnd": DB.XYZ(LENGTH, 0, OFFSET-LENGTH),
            "cutVector": DB.XYZ.BasisZ,
            "view": refView
        },
        int(DB.FamilyInstanceReferenceType.CenterLeftRight): {
            "name": "Center (Left/Right)",
            "bubbleEnd": DB.XYZ(0, -LENGTH, OFFSET-LENGTH),
            "freeEnd": DB.XYZ(0, LENGTH, OFFSET-LENGTH),
            "cutVector": DB.XYZ.BasisZ,
            "view": refView
        },
        int(DB.FamilyInstanceReferenceType.CenterElevation): {
            "name": "Center (Elevation)",
            "bubbleEnd": DB.XYZ(-LENGTH, -LENGTH, OFFSET),
            "freeEnd": DB.XYZ(LENGTH, -LENGTH, OFFSET),
            "cutVector": DB.XYZ.BasisY,
            "view": frontView
        }
    }

    parameterValues = [
        plane.get_Parameter(DB.BuiltInParameter.ELEM_REFERENCE_NAME).AsInteger()
        for plane in planes
    ]

    with Transaction("Create reference frame"):
        for key, data in referenceTypes.items():
            skipReference = False
            if key in parameterValues:
                print("Skipping: {}".format(data['name']))
                skipReference = True
            if not skipReference:
                print("Creating: {}".format(data['name']))
                bubbleEnd = origin.Add(data['bubbleEnd'])
                freeEnd = origin.Add(data['freeEnd'])
                cutVector = data['cutVector']
                newReference = doc.FamilyCreate.NewReferencePlane(
                    bubbleEnd,
                    freeEnd,
                    cutVector,
                    data['view']
                )
                newReference.Name = data['name']
                newReference.get_Parameter(
                    DB.BuiltInParameter.ELEM_REFERENCE_NAME
                ).Set(key)


    #TODO Create missing reference planes
    #TODO Assign correct reference types
    #TODO Name reference planes accordingly