from Autodesk.Revit import DB
from pyrevit import clr, HOST_APP, forms, revit
import System
clr.ImportExtensions(System.Linq)

def COBieTypeIsEnabled(element):
    parameter = element.LookupParameter("COBie.Type")
    return True if parameter.AsInteger() > 0 else False

def COBieEnable(element):
    parameter = element.LookupParameter("COBie")
    parameter.Set(True)

def GetElementSymbol(element):
    if type(element) == DB.Wall:
        return element.WallType
    else:
        return element.Symbol

def COBieComponentAutoSelect(view, doc=None):
    if doc is None:
        doc = HOST_APP.doc

    elements = DB.FilteredElementCollector(doc, view.Id)\
        .WhereElementIsNotElementType()\
        .ToElements()
    familySymbolIds = set()
    for element in elements:
        try:
            symbol = GetElementSymbol(element)
            if COBieTypeIsEnabled(symbol):
                familySymbolIds.add(symbol.Id)
        except Exception as e:
            print("{}: {}".format(element.Id.IntegerValue, e)) 
        
    componentView = DB.FilteredElementCollector(doc)\
        .OfClass(DB.ViewSchedule)\
        .WhereElementIsNotElementType()\
        .Where(lambda x: x.Name == "COBie.Component").First()
    componentElements = DB.FilteredElementCollector(doc, componentView.Id)\
        .ToElements()
    for element in componentElements:
        parameter = element.LookupParameter("COBie")
        if parameter.AsInteger() > 0:
            symbol = GetElementSymbol(element)
            if symbol.Id not in familySymbolIds:
                parameter.Set(0)
    with revit.Transaction("Enable COBie.Component Components"):
        
        for symbolId in familySymbolIds:
            symbolInstances = DB.FilteredElementCollector(doc)\
                .WhereElementIsNotElementType()\
                .WherePasses(
                    DB.FamilyInstanceFilter(doc, symbolId)
                )\
                .ToElements()
            print("len(symbolInstances) = {}".format(len(symbolInstances)))
            for instance in symbolInstances:
                parameter = instance.LookupParameter("COBie")
                parameter.Set(1)
    return

if __name__ == "__main__":
    doc = HOST_APP.doc
    try:
        view = DB.FilteredElementCollector(doc)\
            .OfClass(DB.ViewSchedule)\
            .WhereElementIsNotElementType()\
            .Where(lambda x: x.Name == "COBie.Type").First()
    except:
        forms.alert(
            "Please setup the project for COBie using the BIM Interoperability "
            "Tools before running this command.",
            exitscript=True
        )
    if forms.alert(
        "Are you sure you would like to update the COBie.Component status for "
        "all family instances in the current model?",
        yes=True,
        no=True,
    ):
        COBieComponentAutoSelect(view, doc=doc)


