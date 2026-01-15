from Autodesk.Revit import DB
from flamingo.extensible_storage import (
    SetFlamingoSetting,
    GetFlamingoSetting,
    GetFlamingoSchema,
    GetSchemaMapData,
)
from flamingo import (
    FLAMINGO_PINK_DARK,
    FLAMINGO_PINK_LIGHT,
    FLAMINGO_GREY_LIGHT,
    FLAMINGO_GREY_DARK,
    FLAMINGO_GOLD,
)
from flamingo.revit import GetParameterValueByName, SetParameter
import json
from math import radians, degrees
from pyrevit import HOST_APP, forms, revit, script
import traceback
import System
from System.Windows.Media import SolidColorBrush, Color

LOGGER = script.get_logger()
OUTPUT = script.get_output()
CONFIG = script.get_config()


class FailuresPreprocessor(DB.IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        for msg in failuresAccessor.GetFailureMessages():
            LOGGER.error("Failure: {}".format(msg))
        return DB.FailureProcessingResult.Continue


class TransformManagerWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, transforms):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.doc = doc
        self.schema = GetFlamingoSchema(doc)

        self.transforms = transforms
        self.original_transforms = transforms
        self.updated = False
        self.selected_preset = None

        self._populate_list()

        flamingoPinkLight = Color.FromRgb(
            int(FLAMINGO_PINK_LIGHT[1:3], 16),  # R
            int(FLAMINGO_PINK_LIGHT[3:5], 16),  # G
            int(FLAMINGO_PINK_LIGHT[5:7], 16),  # B
        )
        flamingoPinkDark = Color.FromRgb(
            int(FLAMINGO_PINK_DARK[1:3], 16),  # R
            int(FLAMINGO_PINK_DARK[3:5], 16),  # G
            int(FLAMINGO_PINK_DARK[5:7], 16),  # B
        )
        flamingoGreyLight = Color.FromRgb(
            int(FLAMINGO_GREY_LIGHT[1:3], 16),  # R
            int(FLAMINGO_GREY_LIGHT[3:5], 16),  # G
            int(FLAMINGO_GREY_LIGHT[5:7], 16),  # B
        )
        flamingoGreyDark = Color.FromRgb(
            int(FLAMINGO_GREY_DARK[1:3], 16),   # R
            int(FLAMINGO_GREY_DARK[3:5], 16),   # G
            int(FLAMINGO_GREY_DARK[5:7], 16),   # B
        )
        flamingoGold = Color.FromRgb(
            int(FLAMINGO_GOLD[1:3], 16),  # R
            int(FLAMINGO_GOLD[3:5], 16),  # G
            int(FLAMINGO_GOLD[5:7], 16),  # B
        )

        self.Background = SolidColorBrush(flamingoPinkLight)
        self.Foreground = SolidColorBrush(flamingoGreyDark)
        self.BorderBrush = SolidColorBrush(flamingoGreyDark)

        textBoxes = [
            self.txt_name,
            self.txt_origin,
            self.txt_target,
            self.txt_angle,
            self.txt_scale,
        ]

        for txt in textBoxes:
            # txt.Background = SolidColorBrush(flamingoPinkLight)
            txt.Foreground = SolidColorBrush(flamingoGreyDark)
            txt.BorderBrush = SolidColorBrush(flamingoGreyDark)

        # Apply colors to all Buttons
        buttons = [
            self.btn_add,
            self.btn_remove,
            self.btn_pick_origin,
            self.btn_pick_target,
            self.btn_map_points,
            self.btn_save,
            self.btn_save_all,
        ]

        for btn in buttons:
            btn.Background = SolidColorBrush(flamingoPinkDark)
            btn.Foreground = SolidColorBrush(flamingoGreyDark)

    def _populate_list(self):
        self.preset_listbox.Items.Clear()
        for name in sorted(self.transforms.keys()):
            self.preset_listbox.Items.Add(name)

    def on_preset_selected(self, sender, args):
        if self.preset_listbox.SelectedItem:
            # Enable inputs
            self.txt_name.IsEnabled = True
            self.txt_origin.IsEnabled = True
            self.btn_pick_origin.IsEnabled = True
            self.txt_target.IsEnabled = True
            self.btn_pick_target.IsEnabled = True
            self.txt_angle.IsEnabled = True
            self.txt_scale.IsEnabled = True
            self.btn_map_points.IsEnabled = True
            self.btn_save.IsEnabled = True

            preset_name = str(self.preset_listbox.SelectedItem)
            self.selected_preset = preset_name
            transform = self.transforms[preset_name]

            self.txt_name.Text = preset_name
            self.txt_origin.Text = ", ".join(str(x) for x in transform["origin"])
            self.txt_target.Text = ", ".join(str(x) for x in transform["target"])
            self.txt_angle.Text = str(transform["angle"])
            self.txt_scale.Text = str(transform["scale"])
        else:
            # Disable inputs
            self.txt_name.IsEnabled = False
            self.txt_origin.IsEnabled = False
            self.btn_pick_origin.IsEnabled = False
            self.txt_target.IsEnabled = False
            self.btn_pick_target.IsEnabled = False
            self.txt_angle.IsEnabled = False
            self.txt_scale.IsEnabled = False
            self.btn_map_points.IsEnabled = False
            self.btn_save.IsEnabled = False

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
                # Disable inputs
                self.txt_name.IsEnabled = False
                self.txt_origin.IsEnabled = False
                self.btn_pick_origin.IsEnabled = False
                self.txt_target.IsEnabled = False
                self.btn_pick_target.IsEnabled = False
                self.txt_angle.IsEnabled = False
                self.txt_scale.IsEnabled = False
                self.btn_map_points.IsEnabled = False
                self.btn_save.IsEnabled = False

    def on_pick_origin(self, sender, args):
        self.Hide()
        try:
            point = revit.pick_point("Pick origin point")
            if point:
                self.txt_origin.Text = "{}, {}, {}".format(
                    round(point.X, 4), round(point.Y, 4), round(point.Z, 4)
                )
        except Exception as e:
            LOGGER.debug("Failed to pick origin point: {}".format(e))
        finally:
            self.ShowDialog()

    def on_pick_target(self, sender, args):
        self.Hide()
        try:
            point = revit.pick_point("Pick target point")
            if point:
                self.txt_target.Text = "{}, {}, {}".format(
                    round(point.X, 4), round(point.Y, 4), round(point.Z, 4)
                )
        except Exception as e:
            LOGGER.debug("Failed to pick target point: {}".format(e))
        finally:
            self.ShowDialog()

    def on_map_points(self, sender, args):
        self.Hide()
        try:
            point1a = revit.pick_point("Pick first endpoint of first line")
            if not point1a:
                LOGGER.debug("No point picked for first endpoint of first line")
                self.ShowDialog()
                return
            point1b = revit.pick_point("Pick second endpoint of first line")
            if not point1b:
                LOGGER.debug("No point picked for second endpoint of first line")
                self.ShowDialog()
                return
            point2a = revit.pick_point("Pick first endpoint of second line")
            if not point2a:
                LOGGER.debug("No point picked for first endpoint of second line")
                self.ShowDialog()
                return
            point2b = revit.pick_point("Pick second endpoint of second line")
            if not point2b:
                LOGGER.debug("No point picked for second endpoint of second line")
                self.ShowDialog()
                return
            self.txt_origin.Text = "{}, {}, {}".format(
                round(point1a.X, 4), round(point1a.Y, 4), round(point1a.Z, 4)
            )
            self.txt_target.Text = "{}, {}, {}".format(
                round(point2a.X, 4), round(point2a.Y, 4), round(point2a.Z, 4)
            )
            # LOGGER.set_debug_mode()
            vector1 = (point1b - point1a).Normalize()
            vector2 = (point2b - point2a).Normalize()
            LOGGER.debug("Vector 1: {}".format(vector1))
            LOGGER.debug("Vector 2: {}".format(vector2))
            angle = degrees(vector1.AngleOnPlaneTo(vector2, DB.XYZ.BasisZ))
            LOGGER.debug(
                "vector1.AngleOnPlaneTo(vector2, DB.XYZ.BasisZ): {}".format(
                    degrees(vector1.AngleOnPlaneTo(vector2, DB.XYZ.BasisZ))
                )
            )
            self.txt_angle.Text = str(round(angle, 4))
            scale = point2a.DistanceTo(point2b) / point1a.DistanceTo(point1b)
            self.txt_scale.Text = str(round(scale, 4))
        except Exception as e:
            LOGGER.debug("Failed to map points: {}".format(e))
        finally:
            LOGGER.reset_level()
            self.ShowDialog()

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

        except Exception as e:
            forms.alert("Error saving preset: {}".format(e))

    def on_save_all(self, sender, args):
        try:
            self.on_save_preset(sender, args)
            transforms_json = json.dumps(self.transforms)
            SetFlamingoSetting(
                "Transforms", transforms_json, doc=self.doc, schema=self.schema
            )
            self.updated = True
            self.Close()
        except Exception as e:
            forms.alert("Error saving transforms: {}".format(e))

    def close_window(self, sender, args):
        forms.alert(
            "Closing without saving changes to transforms.",
        )
        self.updated = False
        self.transforms = self.original_transforms
        self.Close()


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
    if script.EXEC_PARAMS.config_mode:
        selection = True
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
    with revit.Transaction("Transform Settings"):
        try:
            schema = GetFlamingoSchema(doc)
            transformSettings = GetFlamingoSetting(
                "Transforms", "{}", doc=doc, schema=schema
            )
            transforms = json.loads(transformSettings)
        except Exception as e:
            LOGGER.warning("No existing transform settings found.\n{}".format(e))
            transforms = {}
        if not transformSettings or script.EXEC_PARAMS.config_mode:
            formOutput = TransformManagerWindow(
                "ManageTransforms.xaml", transforms=transforms
            )
            formOutput.ShowDialog()
            if not formOutput.updated:
                script.exit()
            transforms = formOutput.transforms
            SetFlamingoSetting(
                "Transforms",
                json.dumps(transforms),
                doc=doc,
                schema=schema,
            )
        else:
            transforms = json.loads(transformSettings)
        LOGGER.debug("Loaded transforms: {}".format(transforms))
    if transforms and not script.EXEC_PARAMS.config_mode:
        try:
            lastTransform = json.loads(CONFIG.get_option("last_transform", None))
        except:
            lastTransform = None
        manage = True
        count = 0
        while manage:
            count += 1
            if count > 10:
                LOGGER.error("Too many transform management attempts, exiting.")
                script.exit()
            keys = transforms.keys()
            if lastTransform:
                transforms["Last Transform"] = lastTransform
                keys.insert(0, "Last Transform")

            keys.append("Manage Transforms")
            LOGGER.debug("Available Transforms: {}".format(list(keys)))
            option, switches = forms.CommandSwitchWindow.show(
                keys,
                switches=["inverted"],
                message="Select Transform to Apply",
            )
            if option == "Manage Transforms":
                LOGGER.debug("Opening Transform Manager")
                with revit.Transaction("Manage Transforms"):
                    transforms.pop("Last Transform", None)
                    formOutput = TransformManagerWindow(
                        "ManageTransforms.xaml", transforms=transforms
                    )
                    formOutput.ShowDialog()
                    if formOutput.updated and formOutput.transforms:
                        transforms = formOutput.transforms
                        SetFlamingoSetting(
                            "Transforms",
                            json.dumps(transforms),
                            doc=doc,
                            schema=schema,
                        )
                LOGGER.debug("User entered transform settings: {}".format(formOutput))
            else:
                manage = False
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
                    LOGGER.debug(traceback.format_exc())

        CONFIG.set_option("last_transform", json.dumps(transformData))
        script.save_config()
