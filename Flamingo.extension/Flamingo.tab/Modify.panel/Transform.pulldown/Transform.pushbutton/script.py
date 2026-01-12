from Autodesk.Revit import DB
from flamingo.extensible_storage import (
    SetFlamingoSetting,
    GetFlamingoSetting,
    GetFlamingoSchema,
    GetSchemaMapData,
)
from flamingo.revit import GetParameterValueByName, SetParameter
import json
from math import radians
from pyrevit import HOST_APP, forms, revit, script
import traceback
import System

LOGGER = script.get_logger()
OUTPUT = script.get_output()
CONFIG = script.get_config()


class FailuresPreprocessor(DB.IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        for msg in failuresAccessor.GetFailureMessages():
            LOGGER.error("Failure: {}".format(msg))
        return DB.FailureProcessingResult.Continue


class TransformManagerWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, doc):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.doc = doc
        self.schema = GetFlamingoSchema(doc)

        # Load existing transforms
        transformSettings = GetFlamingoSetting(
            "Transforms", "{}", doc=doc, schema=self.schema
        )
        self.transforms = json.loads(transformSettings) if transformSettings else {}
        self.selected_preset = None

        self._populate_list()

    def _populate_list(self):
        self.preset_listbox.Items.Clear()
        for name in sorted(self.transforms.keys()):
            self.preset_listbox.Items.Add(name)

    def on_preset_selected(self, sender, args):
        if self.preset_listbox.SelectedItem:
            preset_name = str(self.preset_listbox.SelectedItem)
            self.selected_preset = preset_name
            transform = self.transforms[preset_name]

            self.txt_name.Text = preset_name
            self.txt_origin.Text = ", ".join(str(x) for x in transform["origin"])
            self.txt_target.Text = ", ".join(str(x) for x in transform["target"])
            self.txt_angle.Text = str(transform["angle"])
            self.txt_scale.Text = str(transform["scale"])

    def on_add_preset(self, sender, args):
        name = forms.ask_for_string(
            prompt="Enter new preset name:", title="New Transform Preset"
        )
        if name and name not in self.transforms:
            self.transforms[name] = {
                "origin": [0, 0, 0],
                "target": [0, 0, 0],
                "angle": 0,
                "scale": 1.0,
            }
            self._populate_list()
            # Select the new preset
            for i in range(self.preset_listbox.Items.Count):
                if str(self.preset_listbox.Items[i]) == name:
                    self.preset_listbox.SelectedIndex = i
                    break
        elif name in self.transforms:
            forms.alert("Preset '{}' already exists.".format(name))

    def on_remove_preset(self, sender, args):
        if self.preset_listbox.SelectedItem:
            preset_name = str(self.preset_listbox.SelectedItem)
            if forms.alert(
                "Remove preset '{}'?".format(preset_name), yes=True, no=True
            ):
                del self.transforms[preset_name]
                self._populate_list()
                self.txt_name.Text = ""
                self.txt_origin.Text = ""
                self.txt_target.Text = ""
                self.txt_angle.Text = ""
                self.txt_scale.Text = ""

    def on_save_preset(self, sender, args):
        try:
            new_name = str(self.txt_name.Text).strip()
            if not new_name:
                forms.alert("Please enter a preset name.")
                return

            origin = [float(x.strip()) for x in self.txt_origin.Text.split(",")]
            target = [float(x.strip()) for x in self.txt_target.Text.split(",")]
            angle = float(self.txt_angle.Text)
            scale = float(self.txt_scale.Text)

            if len(origin) != 3 or len(target) != 3:
                forms.alert("Origin and Target must have 3 values (X, Y, Z).")
                return

            # Handle rename
            if self.selected_preset and self.selected_preset != new_name:
                del self.transforms[self.selected_preset]

            self.transforms[new_name] = {
                "origin": origin,
                "target": target,
                "angle": angle,
                "scale": scale,
            }

            self.selected_preset = new_name
            self._populate_list()

            # Reselect the saved preset
            for i in range(self.preset_listbox.Items.Count):
                if str(self.preset_listbox.Items[i]) == new_name:
                    self.preset_listbox.SelectedIndex = i
                    break

            forms.alert("Preset '{}' saved.".format(new_name), title="Success")
        except Exception as e:
            forms.alert("Error saving preset: {}".format(e))

    def on_save_all(self, sender, args):
        try:
            transforms_json = json.dumps(self.transforms)
            SetFlamingoSetting(
                "Transforms", transforms_json, doc=self.doc, schema=self.schema
            )
            forms.alert("All transforms saved to project.", title="Success")
            self.Close()
        except Exception as e:
            forms.alert("Error saving transforms: {}".format(e))


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
    transformFlat = DB.Transform.CreateTranslation(
        DB.XYZ(0, 0, newCurve.GetEndPoint(0).Z) - endPointZ
    )
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

        # assert
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
    if HOST_APP.version < 2021:
        forms.alert(
            "Flamingo Transform requires Revit 2021 or later.",
            title="Flamingo: Transform",
            exitscript=True,
        )
    with revit.Transaction("Create Flamingo Settings"):
        LOGGER.set_debug_mode()
        try:
            schema = GetFlamingoSchema(doc)
            settings = GetSchemaMapData(schema, "Settings", doc.ProjectInformation)
            LOGGER.debug("settings by getmap: {}".format(settings))

            transformSettings = GetFlamingoSetting(
                "Transforms", {}, doc=doc, schema=schema
            )
        except Exception as e:
            LOGGER.warning("No existing transform settings found.\n{}".format(e))
        if not transformSettings:
            formOutput = forms.ask_for_string(
                '"Scale Down Rotate": { "origin": [2, 1, 0], "target": [-1, -2, 0], "angle": 15, "scale": 0.5 }',
                title="Flamingo: Transform",
                exitscript=True,
            )
            LOGGER.debug("User entered transform settings: {}".format(formOutput))
            if formOutput:
                transformSettings = "{{{}}}".format(formOutput)
                LOGGER.set_debug_mode()
                LOGGER.debug(
                    "Storing new transform settings: {}".format(transformSettings)
                )
                SetFlamingoSetting(
                    "Transforms", transformSettings, doc=doc, schema=schema
                )
                transforms = json.loads(transformSettings)
            else:
                transforms = None
            LOGGER.reset_level()
        else:
            transforms = json.loads(transformSettings)
        LOGGER.debug("Loaded transforms: {}".format(transforms))
        LOGGER.reset_level()
    if not transforms:
        script.exit()

    # transforms = {
    #     "Scale Down Rotate": { "origin": (2, 1, 0), "target": (-1, -2, 0), "angle": 15, "scale": 0.5 }
    #     "Scale Up Rotate Negative": {
    #         "origin": (-12, -12, 0),
    #         "target": (-1, -2, 0),
    #         "angle": -60,
    #         "scale": 2.0,
    #     },
    #     "No scale": {
    #         "origin": (-16, -16, 0),
    #         "target": (16, 16, 0),
    #         "angle": 90,
    #         "scale": 1.0,
    #     },
    # }
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
