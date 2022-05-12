from Autodesk.Revit import DB
from flamingo.cobie import GetCDXSpaceBOMAValues, SetCDXSpaceBOMAValue
from flamingo.cobie import CDX_ANSI_BOMA_CODE_MAP
from pyrevit import forms, HOST_APP, revit


class bomaData:
    def __init__(self, data):
        self.roomElement = data["roomElement"]
        self.bomaName = data["bomaName"]
        self.bomaCode = data["bomaCode"]
        self.roomNumber = data["roomNumber"]
        self.roomName = data["roomName"]
        self.description = "{}: {} (Current ANSI-BOMA: {}: {})".format(
            self.roomNumber, self.roomName, self.bomaCode, self.bomaName
        )


if __name__ == "__main__":
    doc = HOST_APP.doc
    roomBOMAs = GetCDXSpaceBOMAValues(doc)
    if not roomBOMAs:
        forms.alert(
            "Please setup the project for COBie and CDX using the BIM Interoperability "
            "Tools before running this command.",
            exitscript=True,
        )
    roomData = [bomaData(item) for item in roomBOMAs]
    value = None
    n = 0
    while not value:
        n += 1
        if n > 10:
            break
        commandOptions = sorted(
            [
                "{}: {}".format(value, key)
                for key, value in CDX_ANSI_BOMA_CODE_MAP.items()
            ]
        ) + ["Help", "Exit"]
        value = forms.CommandSwitchWindow.show(
            commandOptions,
            message="Select ANSI-BOMA value to assign:",
        )
        if not value or value == "Exit":
            break
        if value == "Help":
            value = None
            forms.alert(
                "First select a value from the list of eligible values provided. "
                "Then select the rooms you would like to apply the value "
                'chosen in the previous step. Click "Apply Value To Rooms" to '
                "assign the value to the rooms.\n\n"
                "After the values are assigned, you will be asked a again for a value "
                'to assign. If you are done assigning values, select "Exit"'
            )
            continue
        roomsOut = forms.SelectFromList.show(
            title='Select rooms to apply the ANSI-BOMA value of "{}":'.format(
                value
            ),
            context=roomData,
            name_attr="description",
            button_name='Apply "{}" to Rooms'.format(value),
            multiselect=True,
        )
        if not roomsOut:
            value = None
            continue
        with revit.Transaction("Set CUI-SBU values"):
            [
                SetCDXSpaceBOMAValue(
                    item.roomElement, (value.split(": "))[1], CDX_ANSI_BOMA_CODE_MAP
                )
                for item in roomsOut
            ]
        roomBOMAs = GetCDXSpaceBOMAValues(doc)
        roomData = [bomaData(item) for item in roomBOMAs]
        value = None
