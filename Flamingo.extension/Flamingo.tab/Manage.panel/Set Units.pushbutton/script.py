from Autodesk.Revit import DB
from pyrevit import HOST_APP, forms, revit, script

if __name__ == "__main__":
    doc = HOST_APP.doc

    command = forms.CommandSwitchWindow.show(
        context=["Feet and Inches", "Fractional Inches", "Millimeters"],
        title="Select Units",
    )

    if not command:
        script.exit()

    units = doc.GetUnits()
    if command == "Fractional Inches":
        formatOptions = DB.FormatOptions(DB.UnitTypeId.FractionalInches)
        formatOptions.SetSymbolTypeId(DB.SymbolTypeId.InchDoubleQuote)
        formatOptions.Accuracy = 0.00390625
    elif command == "Millimeters":
        formatOptions = DB.FormatOptions(DB.UnitTypeId.Millimeters)
        formatOptions.Accuracy = 0.001

    else:
        formatOptions = DB.FormatOptions(DB.UnitTypeId.FeetFractionalInches)
        formatOptions.Accuracy = 0.00390625
    units.SetFormatOptions(DB.SpecTypeId.Length, formatOptions)
    with revit.Transaction("Set Units"):
        doc.SetUnits(units)
