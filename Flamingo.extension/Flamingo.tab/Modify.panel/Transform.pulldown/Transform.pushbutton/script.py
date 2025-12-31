from Autodesk.Revit import DB
from flamingo.extensible_storage import GetFlamingoSchema, GetSchemaData
from flamingo.revit import GetParameterValueByName, SetParameter
import json
from math import radians
from pyrevit import HOST_APP, forms, revit, script
import traceback

LOGGER = script.get_logger()
OUTPUT = script.get_output()
CONFIG = script.get_config()


class FailuresPreprocessor(DB.IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        for msg in failuresAccessor.GetFailureMessages():
            LOGGER.error("Failure: {}".format(msg))
        return DB.FailureProcessingResult.Continue


def TranslatePoint(point, transformData):
    origin = DB.XYZ(*transformData["origin"])
    target = DB.XYZ(*transformData["target"])
    angle = transformData["angle"]
    scale = transformData["scale"]
    transformMove = DB.Transform.CreateTranslation(target - origin)
    transformRotate = DB.Transform.CreateRotationAtPoint(
        DB.XYZ.BasisZ, radians(angle), target
    )
    # TODO: Create option for 3D scale
    translatedPoint = origin + DB.XYZ(
        (point.X - origin.X) * scale,
        (point.Y - origin.Y) * scale,
        0,
    )
    translatedPoint = transformMove.OfPoint(translatedPoint)
    translatedPoint = transformRotate.OfPoint(translatedPoint)
    return translatedPoint


def TranslateCurve(curve, transformData):
    LOGGER.debug("Translating Curve")
    # LOGGER.set_debug_mode()
    origin = DB.XYZ(*transformData["origin"])
    target = DB.XYZ(*transformData["target"])
    angle = transformData["angle"]
    scale = transformData["scale"]
    transformMove = DB.Transform.CreateTranslation(target - origin)
    transformRotate = DB.Transform.CreateRotationAtPoint(
        DB.XYZ.BasisZ, radians(angle), target
    )
    transformToOrigin = DB.Transform.CreateTranslation(-origin)
    transformScale = DB.Transform.Identity.ScaleBasis(scale)
    endPointZ = DB.XYZ(0, 0, curve.GetEndPoint(0).Z)
    newCurve = curve.CreateTransformed(transformToOrigin)
    newCurve = newCurve.CreateTransformed(transformScale)
    newCurve = newCurve.CreateTransformed(transformToOrigin.Inverse)
    transformFlat = DB.Transform.CreateTranslation(DB.XYZ(0,0,newCurve.GetEndPoint(0).Z) - endPointZ)
    newCurve = newCurve.CreateTransformed(transformFlat)
    newCurve = newCurve.CreateTransformed(transformMove)
    newCurve = newCurve.CreateTransformed(transformRotate)
    return newCurve


def transformElement(element, transformData):
    """Transforms an element based on element type using the provided


    Args:
        element (DB.Element): Revit Element
        transformData (dict): Transform dictionary
    """

    assert element.IsModifiable, "Element is not modifiable"
    assert hasattr(element, "Location"), "Element has no location"

    origin = DB.XYZ(*transformData["origin"])
    target = DB.XYZ(*transformData["target"])
    angle = transformData["angle"]
    scale = transformData["scale"]
    transformMove = DB.Transform.CreateTranslation(target - origin)
    transformRotate = DB.Transform.CreateRotationAtPoint(
        DB.XYZ.BasisZ, radians(angle), target
    )
    location = element.Location

    if isinstance(location, DB.LocationPoint):
        LOGGER.debug("Transforming Element with LocationPoint {}".format(element.Id))
        # TODO: Create option for 3D scale
        point = origin + DB.XYZ(
            (location.Point.X - origin.X) * scale,
            (location.Point.Y - origin.Y) * scale,
            0,
        )
        element.Location.Point = point
        element.Location.Rotate(
            DB.Line.CreateBound(origin, origin + DB.XYZ.BasisZ), radians(angle)
        )
        element.Location.Move(target - origin)

    if isinstance(location, DB.LocationCurve):
        LOGGER.debug("Transforming Element with LocationCurve {}".format(element.Id))
        element.Location.Curve = TranslateCurve(element.Location.Curve, transformData)
    elif hasattr(element, "SketchId"):
        LOGGER.debug("Transforming Element with SketchId {}".format(element.Id))
        sketch = element.Document.GetElement(element.SketchId)

        assert 
    elif isinstance(element, DB.TextNote):
        LOGGER.debug("Transforming TextNote {}".format(element.Id))
        # TODO: Handle text note
    elif hasattr(element, "BoundingBox"):
        LOGGER.debug("Transforming Element by BoundingBox {}".format(element.Id))
    else:
        raise Exception("Unsupported element type for transform")


if __name__ == "__main__":
    doc = HOST_APP.doc
    selection = revit.get_selection()
    if not selection:
        forms.alert(
            "Please select elements to transform first.",
            title="Flaming: Transform",
            exitscript=True,
        )
    # TODO: Save transform settings to project
    # schema = GetFlamingoSchema(doc=doc)
    # if HOST_APP.version < 2021:
    #     forms.alert(
    #         "Flamingo Transform requires Revit 2021 or later.",
    #         title="Flamingo: Transform",
    #         exitscript=True,
    #     )
    # schemaSettings = doc.ProjectInformation.Get(schema.GetField("Settings"))
    # if not "transforms" in schemaSettings:
    #     forms.ask_for_string(
    #         "No saved transforms found in project.\nPlease add transforms via the Manage Transforms tool first.",
    #         title="Flamingo: Transform",
    #         exitscript=True,
    #     )
    transforms = {
        "Scale Down Rotate": {
            "origin": (2, 1, 0),
            "target": (-1, -2, 0),
            "angle": 15,
            "scale": 0.5,
        },
        "Scale Up Rotate Negative": {
            "origin": (-12, -12, 0),
            "target": (-1, -2, 0),
            "angle": -60,
            "scale": 2.0,
        },
        "No scale": {
            "origin": (-16, -16, 0),
            "target": (16, 16, 0),
            "angle": 90,
            "scale": 1.0,
        },
    }
    keys = transforms.keys()
    try:
        lastTransform = CONFIG.get_option("last_transform", None)
        if lastTransform:
            transforms["Last Transform"] = json.loads(lastTransform)
            keys.insert(0, "Last Transform")
    except:
        pass

    LOGGER.warning("Available Transforms: {}".format(list(keys)))
    option, switches = forms.CommandSwitchWindow.show(
        keys,
        switches=["inverted"],
        message="Select Transform to Apply",
    )
    # TODO: Save switch states to config
    # TODO: Add option to move grouped
    if not option:
        script.exit()

    transformData = transforms[option]
    if switches.get("inverted"):
        transformData["angle"] = -transformData["angle"]
        transformData["scale"] = 1.0 / transformData["scale"]
        transformData["origin"], transformData["target"] = (
            transformData["target"],
            transformData["origin"],
        )

    for key, value in transformData.items():
        LOGGER.debug("{}: {}".format(key, value))

    with revit.Transaction("Transform {}".format(option)):
        unableToTranslate = []
        for element in selection:
            try:
                transformElement(element, transformData)
            except Exception as e:
                LOGGER.warning(
                    "Failed to transform element {}: {}".format(
                        OUTPUT.linkify(element.Id), e
                    )
                )
                LOGGER.set_debug_mode()
                LOGGER.debug(traceback.format_exc())
                LOGGER.reset_level()

    CONFIG.set_option("last_transform", json.dumps(transformData))
    script.save_config()
