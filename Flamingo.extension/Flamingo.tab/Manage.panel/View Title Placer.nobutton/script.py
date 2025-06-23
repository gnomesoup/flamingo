from Autodesk.Revit import DB
from pyrevit import HOST_APP, revit, script


LOGGER = script.get_logger()
OUTPUT = script.get_output()

if __name__ == "__main__":
    doc = HOST_APP.doc

    LOGGER.set_debug_mode()
    selection = revit.get_selection()
    if selection and len(selection) >= 1:
        selection = selection[0] if type(selection[0]) is DB.Viewport else None
    if not selection:
        selection = revit.pick_element_by_category("Viewports", "Select a viewport")
    startPoint = revit.pick_point("Pick a start point to place the title")
    endPoint = revit.pick_point("Pick an end point")

    labelOffset = selection.LabelOffset
    viewOutline = selection.GetBoxOutline()
    min = viewOutline.MinimumPoint
    min = DB.XYZ(min.X, min.Y, 0)
    max = viewOutline.MaximumPoint
    max = DB.XYZ(max.X, max.Y, 0)
    LOGGER.reset_level()
    offsetVector = DB.XYZ((41.0 / 32.0) / 12.0, (21.0 / 64.0) / 12.0, 0)
    labelLineLength = endPoint.X - startPoint.X - ((41.0 / 32.0) / 12.0)
    with revit.Transaction("Move View Title"):
        selection.LabelOffset = startPoint - min + offsetVector
        selection.LabelLineLength = (
            labelLineLength if labelLineLength > 0 else 1.0 / 12.0
        )
