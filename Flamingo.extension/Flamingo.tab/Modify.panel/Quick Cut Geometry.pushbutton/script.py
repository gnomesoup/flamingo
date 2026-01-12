from Autodesk.Revit import DB
from pyrevit import HOST_APP, forms, revit, script
from clr import StrongBox

OUTPUT = script.get_output()
LOGGER = script.get_logger()

if __name__ == "__main__":
    doc = HOST_APP.doc

    toCut = revit.pick_elements("Select elements to be cut")
    if not toCut:
        script.exit()
    cuttingElement = revit.pick_element("Select cutting element")
    if not cuttingElement:
        script.exit()
    if not DB.SolidSolidCutUtils.IsAllowedForSolidCut(cuttingElement):
        forms.alert(
            "Cannot cut with the selected cutting element.",
            exitscript=True,
        )
    # LOGGER.set_debug_mode()
    with revit.Transaction("Quick Cut Geometry"):
        for element in toCut:
            if not DB.SolidSolidCutUtils.IsAllowedForSolidCut(element):
                LOGGER.warning(
                    "Cannot cut with element: {}".format(OUTPUT.linkify(element.Id))
                )
            reason = StrongBox[DB.CutFailureReason]()
            DB.SolidSolidCutUtils.CanElementCutElement(element, cuttingElement, reason)
            LOGGER.info(
                "Cutting {} with {}: {}".format(element.Id, cuttingElement.Id, reason)
            )
            if reason.Equals(DB.CutFailureReason.CutAllowed):
                try:
                    DB.SolidSolidCutUtils.AddCutBetweenSolids(
                        doc, element, cuttingElement
                    )
                except Exception as e:
                    LOGGER.error("Cutting failed: {}".format(e))
            else:
                LOGGER.warning("Unable to cut: {}".format(reason))
    # LOGGER.reset_level()
