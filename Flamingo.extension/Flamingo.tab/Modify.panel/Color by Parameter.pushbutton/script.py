from Autodesk.Revit import DB
from flamingo.color import COLORBREWER_QUALITATIVE_1, COLORBREWER_QUALITATIVE_2
from flamingo.revit import GetParameterValueByName, GetSolidFillId
from pyrevit import clr, revit, forms, HOST_APP, script
from pyrevit.framework import List, Type

import System

clr.ImportExtensions(System.Linq)

LOGGER = script.get_logger()
OUTPUT = script.get_output()

if __name__ == "__main__":
    colors = COLORBREWER_QUALITATIVE_1 + COLORBREWER_QUALITATIVE_2
    doc = HOST_APP.doc
    activeView = doc.ActiveView

    # get all parameters

    multiClassTypes = List[Type]()
    for t in [DB.SharedParameterElement, DB.ParameterElement]:
        multiClassTypes.Add(t)
    parameters = (
        DB.FilteredElementCollector(doc)
        .WherePasses(DB.ElementMulticlassFilter(multiClassTypes))
        .ToElements()
    )
    parametersByName = {
        "{} ({})".format(
            p.Name, "Shared" if type(p) == DB.SharedParameterElement else "Project"
        ): p
        for p in parameters
    }
    selection = forms.SelectFromList.show(sorted(parametersByName.keys()))
    if not selection:
        script.exit()
    parameter = parametersByName[selection]

    # get all elements visible in view that have the selected parameter
    elements = (
        DB.FilteredElementCollector(doc, activeView.Id)
        .WhereElementIsNotElementType()
        .WhereElementIsViewIndependent()
    )
    elements = [e for e in elements if e.LookupParameter(parameter.Name)]
    elementValuesById = {
        e.Id: GetParameterValueByName(e, parameter.Name) for e in elements
    }
    colorsByValue = {
        value: colors[i % len(colors)]
        for i, value in enumerate(
            set(value for value in elementValuesById.values() if value is not None)
        )
    }

    # get <Solid Fill> pattern id
    solidFillPatternId = GetSolidFillId(doc)
    LOGGER.debug("Solid Fill pattern id: {}".format(solidFillPatternId))

    # override elements in view
    with revit.Transaction("Color by Parameter"):
        LOGGER.set_debug_mode()
        for e in elements:
            LOGGER.debug("Overriding element {}".format(OUTPUT.linkify(e.Id)))
            overrideGraphingSettings = DB.OverrideGraphicSettings()
            if elementValuesById[e.Id] in colorsByValue:
                LOGGER.debug("element value: {}".format(elementValuesById[e.Id]))
                overrideGraphingSettings.SetSurfaceForegroundPatternId(
                    solidFillPatternId
                )
                overrideGraphingSettings.SetSurfaceForegroundPatternColor(
                    colorsByValue[elementValuesById[e.Id]]
                )
            activeView.SetElementOverrides(
                e.Id,
                overrideGraphingSettings,
            )
        LOGGER.reset_level()
