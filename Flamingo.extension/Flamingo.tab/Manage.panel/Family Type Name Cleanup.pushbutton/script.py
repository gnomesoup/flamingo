from Autodesk.Revit import DB
from pyrevit import revit, HOST_APP, script
import re

LOGGER = script.get_logger()
OUTPUT = script.get_output()

if __name__ == "__main__":
    LOGGER.set_debug_mode()
    doc = HOST_APP.doc
    familySymbols = (
        DB.FilteredElementCollector(doc)
        .OfClass(DB.FamilySymbol)
        .ToElements()
    )
    familiesToClean = {}
    for i in familySymbols:
        name = revit.query.get_name(i)
        if not re.match(r"Family\d+", name):
            continue
        family = i.Family
        familyName = family.Name
        if name != familyName:
            familyId = family.Id
            if not familyId in familiesToClean:
                familiesToClean[familyId] = [i.Id]
            else:
                familiesToClean[familyId].append(i.Id)  
    with revit.Transaction("Family Type Name Cleanup"):
        for familyId, symbolIds in familiesToClean.items():
            family = doc.GetElement(familyId)
            familyName = family.Name
            if len(symbolIds) == 1:
                symbol = doc.GetElement(symbolIds[0])
                symbol.Name = familyName
                oldName = revit.query.get_name(symbol)
                print("Renamed {} to {}".format(oldName, familyName))
            else:
                for n, symbolId in enumerate(symbolIds):
                    symbol = doc.GetElement(symbolId)
                    oldName = revit.query.get_name(symbol)
                    symbol.Name = "{}{}".format(familyName, n + 1)
                    print("Renamed {} to {}".format(oldName, familyName))
    LOGGER.reset_level()