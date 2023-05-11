from Autodesk.Revit import DB
from pyrevit import forms, HOST_APP, revit, script

OUTPUT = script.get_output()

if __name__ == "__main__":
    doc = HOST_APP.doc
    OUTPUT.print_md("# Selected Lines")
    selection = revit.get_selection()
    length = 0
    errorState = False
    for element in selection:
        try:
            curveLength = element.GeometryCurve.Length
            OUTPUT.print_md("{}: {} ft".format(OUTPUT.linkify(element.Id), curveLength))
            length += curveLength
        except Exception as e:
            errorState = True
            print(e)
    OUTPUT.print_md("# Total: {}".format(length))
    
