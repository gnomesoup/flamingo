from Autodesk.Revit import DB
from flamingo.cobie import COBieIsEnabled, GetCDXSpaceCUIValues, SetCDXSpaceCUIValue
from pyrevit import forms, HOST_APP, script, revit
import System

class cuiData():
    def __init__(self, data):
        self.roomElement = data['roomElement']
        self.cuiValue = data['cuiValue']
        self.roomNumber = data['roomNumber']
        self.roomName = data['roomName']
        self.description = "{}: {} (Current CUI-SBU: {})".format(
            self.roomNumber,
            self.roomName,
            self.cuiValue,
        )

if __name__ == "__main__":
    doc = HOST_APP.doc
    roomCUIs = GetCDXSpaceCUIValues(doc)
    if not roomCUIs:
        forms.alert(
            "Please setup the project for COBie and CDX using the BIM Interoperability "
            "Tools before running this command.",
            exitscript=True,
        )
    roomData = [cuiData(item) for item in roomCUIs]
    value = None
    n = 0
    while not value:
        n += 1
        if n > 10:
            break
        commandOptions = ["Yes", "No", "NULL", "Help", "Exit"]
        value = forms.CommandSwitchWindow.show(
            commandOptions,
            message="Select CUI-SBU value to assign:",
        )
        if not value or value == "Exit":
            break
        if value == "Help":
            value = None
            forms.alert(
                'First select a value from the list provided, "Yes", "No", or "'
                '"NULL". Then select the rooms you would like to apply the value '
                'chosen in the previous step. Click "Apply Value To Rooms" to '
                "assign the value to the rooms.\n\n"
                "After the values are assigned, you will be asked a again for a value "
                'to assign. If you are done assigning values, select "Exit"'
            )
            continue
        roomsOut = forms.SelectFromList.show(
            title='Select rooms to apply the CUI-SBU value of "{}":'.format(value),
            context=roomData,
            name_attr="description",
            button_name='Apply "{}" to Rooms'.format(value),
            multiselect=True,
        )
        if not roomsOut:
            value = None
            continue
        with revit.Transaction("Set CUI-SBU values"):
            [SetCDXSpaceCUIValue(item.roomElement, value) for item in roomsOut]
        roomCUIs = GetCDXSpaceCUIValues(doc)
        roomData = [cuiData(item) for item in roomCUIs]
        value = None
