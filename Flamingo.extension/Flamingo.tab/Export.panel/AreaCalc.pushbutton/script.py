# -*- coding: utf-8 -*-
import clr
from pyrevit import DB, HOST_APP, script, revit, forms
import re
from os import path
from flamingo.revit import GetParameterFromProjectInfo
from flamingo.excel import OpenWorkbook
from string import ascii_uppercase

import System
clr.ImportExtensions(System.Linq)
clr.AddReference("System.Drawing")

# TODO Correct delete issue
# TODO Allow for occupant storage as valid occupant type
# TODO Title with project name/number - Method - Date
# TODO Radio button for method selection
# TODO Save excel file with date on 03 Area Calc

class areaCalcArea():
    def __init__(self,
                 name = "",
                 areaType = "",
                 level = "",
                 number = "",
                 area = 0):
        self.Name = name
        self.Number = number
        self.AreaType = areaType
        self.Level = level
        self.Area = area
        self.Exclusion = 0
        self.Tenant = 0
        self.Ancillary = 0
        self.Amenity = 0
        self.BuildingService = 0
        self.Circulation = 0
        self.Sort = 0
        self.RowNumber = None
        self.SectionStartRow = None
        self.SectionEndRow = None
        self.BaseBuildingCirculationAlert = False
        self.IsAncillary = False
        self.TenantAreaRatio = 0

    def ByRevitArea(self, method, revitArea):
        self.Name = revitArea.GetParameters("Name")[0].AsString()
        self.Number = revitArea.Number.ToString()
        self.Area = round(revitArea.Area, 2)
        self.Level = revitArea.Level.Name
        self.AreaType = revitArea.GetParameters("Area Calc")[0].AsString()
        if not self.AreaType:
            self.AreaType = "\"Area Type\" left blank"

        # Check for major verticals
        if (self.AreaType == "Major Vertical Penetrations" or
                self.AreaType == "Major Vertical Penetration"):
            self.Exclusion = self.Area
            self.Name = "MAJOR VERTICAL PENETRATIONS"
            self.Sort = 3
        # Check for exclusions
        elif (self.AreaType == "Storage" or
              self.AreaType == "Parking Area" or
              self.AreaType == "Parking" or
              self.AreaType == "Unenclosed Building Feature"):
            self.Exclusion = self.Area
            self.Sort = 2

        # Check for amenity space. We need to handle this differently per
        # method
        elif self.AreaType == "Building Amenity Area":
            if method == "A":
                self.Amenity = self.Area

        # Check for building service areas. This is only required for Method A
        elif (method == "A" and self.AreaType == "Building Service Area"):
            self.BuildingService = self.Area

        # Check for tenant space
        elif (self.AreaType == "Tenant Area" or
              self.AreaType == "Occupant Storage Area" or
              self.AreaType == "Occupant Storage" or
              self.AreaType == "Occupant Area"):
            self.Tenant = self.Area
        # TODO is assigning an ancillary area still necessary?
        # TODO how are ancillary handled from method A?
        elif (self.AreaType == "Tenant Ancillary Area"):
            self.Ancillary = self.Area
            if method == "B":
                self.IsAncillary = True
        else:
            self.Ancillary = 0
        if self.AreaType == "Base Building Circulation":
            if method == "A":
                self.BaseBuildingCirculationAlert = True
            self.Circulation = self.Area
            self.Sort = 2
        else:
            self.Circulation = 0

    def AddAreas(self, addArea):
        if addArea.Area > 0:
            self.Area = self.Area + addArea.Area
        if addArea.Ancillary > 0:
            self.Ancillary = self.Ancillary + addArea.Ancillary
        if addArea.Tenant > 0:
            self.Tenant = self.Tenant + addArea.Tenant
        if addArea.Circulation > 0:
            self.Circulation = self.Circulation + addArea.Circulation
        if addArea.Exclusion > 0:
            self.Exclusion = self.Exclusion + addArea.Exclusion
        if addArea.Amenity > 0:
            self.Amenity = self.Amenity + addArea.Amenity
        if addArea.BuildingService > 0:
            self.BuildingService = self.BuildingService + \
                addArea.BuildingService

    def ItemizeCheck(self):
        if self.Area <= 0:
            return False
        elif (self.AreaType == "Total" or
              self.AreaType == "BLANK"):
            return True
        elif (method == "B" and self.IsAncillary):
            return False
        elif (self.Ancillary <= 0 and
              self.Tenant <= 0 and
              self.Exclusion <= 0 and
              self.Amenity <= 0 and
              self.BuildingService <= 0):
            return False
        else:
            return True

    def Itemize(self):
        return [self.RowNumber,
                self.Number,
                self.Name,
                self.AreaType,
                self.Area,
                self.Exclusion,
                self.Tenant,
                self.Ancillary,
                self.Circulation]

    def ZeroToDashes(self):
        if self.Exclusion <= 0:
            self.Exclusion = "--"
        if self.Tenant <= 0:
            self.Tenant = "--"
        if self.Ancillary <= 0:
            self.Ancillary = "--"
        if self.Circulation <= 0:
            self.Circulation = "--"
        if self.Amenity <= 0:
            self.Amenity = "--"
        if self.Tenant <= 0:
            self.Tenant = "--"

    def MakeRow(self, method, lastRow):
        if method == "A":
            if self.AreaType == "Total":
                rNum = self.RowNumber
                sNum = self.SectionStartRow
                return [
                    # A
                    "Floor Totals",
                    # B
                    self.Area,
                    # C
                    "=SUBTOTAL(9,C" + str(sNum) + ":C" + str(rNum - 1) + ")",
                    # D
                    "=B" + str(rNum) + "-C" + str(rNum),
                    # E
                    "",
                    # F
                    "=SUBTOTAL(9,F" + str(sNum) + ":F" + str(rNum - 1) + ")",
                    # G
                    "=SUBTOTAL(9,G" + str(sNum) + ":G" + str(rNum - 1) + ")",
                    # H
                    "=F" + str(rNum) + "+G" + str(rNum),
                    # I
                    "=SUBTOTAL(9,I" + str(sNum) + ":I" + str(rNum - 1) + ")",
                    # J
                    "=H" + str(rNum) + "+I" + str(rNum),
                    # K
                    "=SUBTOTAL(9,K" + str(sNum) + ":K" + str(rNum - 1) + ")",
                    # L
                    "=D" + str(rNum) + "-J" + str(rNum) + "-K" + str(rNum),
                    # M
                    "=IF(J" + str(rNum) + "=0, 0, (J" + str(rNum) + \
                    "+L" + str(rNum) + ")/J" + str(rNum) + ")",
                    # N
                    "=H" + str(rNum) + "*M" + str(rNum),
                    # O
                    "=I" + str(rNum) + "*M" + str(rNum) + "+K" + str(rNum),
                    # P
                    "=D$" + str(lastRow) + "/(D$" + str(lastRow) + \
                    "-O$" + str(lastRow) + ")",
                    # Q
                    "=N" + str(rNum) + "*P" + str(rNum),
                    # R
                    "--"
                    ]

            elif self.AreaType == "BLANK":
                row = []
                for i in range(18):
                    row.append("")
                return row

            else:
                self.ZeroToDashes()
                rNum = self.RowNumber
                eNum = self.SectionEndRow
                return [
                    # A
                    self.Level,
                    # B
                    "",
                    # C
                    self.Exclusion,
                    # D
                    "",
                    # E
                    self.Name,
                    # F
                    self.Tenant,
                    # G
                    self.Ancillary,
                    # H
                    "=SUMIF(F" + str(rNum) + ":G" + str(rNum) + ",\">0\")",
                    # I
                    self.Amenity,
                    # J
                    "=SUMIF(H" + str(rNum) + ":I" + str(rNum) + ",\">0\")",
                    # K
                    self.BuildingService,
                    # L
                    "",
                    # M
                    "=IF(J$" + str(eNum) + "=0,0,(J$" + str(eNum) + \
                    "+L$" + str(eNum) + ")/J$" + str(eNum) + ")",
                    # N
                    "=IF(COUNTIF(H" + str(rNum) + ",\">0\"),H" + \
                    str(rNum) + "*M" + str(rNum) + ",\"--\")",
                    # O
                    "=IF(COUNTIF(I" + str(rNum) + ",\">0\"),I" + \
                    str(rNum) + "*M" + str(rNum) + ",0)+IF(COUNTIF(K" + \
                    str(rNum) + ",\">0\"),K" + str(rNum) + ",0)",
                    # P
                    "=D$" + str(lastRow) + "/(D$" + str(lastRow) + \
                    "-O$" + str(lastRow) + ")",
                    # Q
                    "=IF(COUNTIF(H" + str(rNum) + ",\">0\"),N" + \
                    str(rNum) + "*P" + str(rNum) + ",\"--\")",
                    # R
                    "=IF(COUNTIF(Q" + str(rNum) + ",\">0\"),Q" + \
                    str(rNum) + "/H" + str(rNum) + ",\"--\")"
                    ]

        elif method == "B":
            if self.AreaType == "Total":
                if self.Circulation <= 0:
                    circulation = ""
                else:
                    circulation = self.Circulation
                return [
                    # A
                    "Floor Totals",
                    # B
                    self.Area,
                    # C
                    "=SUBTOTAL(9,C" + str(self.SectionStartRow) + ":C" + \
                    str(self.RowNumber - 1) + ")",
                    # D
                    "=B" + str(self.RowNumber) + "-C" + str(self.RowNumber),
                    # E
                    "--",
                    # F
                    "=SUBTOTAL(9,F" + str(self.SectionStartRow) + ":F" + \
                    str(self.RowNumber - 1) + ")",
                    # G
                    "=SUBTOTAL(9,G" + str(self.SectionStartRow) + ":G" + \
                    str(self.RowNumber - 1) + ")",
                    # H
                    "=SUBTOTAL(9,H" + str(self.SectionStartRow) + ":H" + \
                    str(self.RowNumber - 1) + ")",
                    # I
                    circulation,
                    # J
                    "=D" + str(self.RowNumber) + "-H" + str(self.RowNumber) + \
                    "-I" + str(self.RowNumber),
                    # K
                    "=D$" + str(lastRow) + "/H$" + str(lastRow),
                    # K
                    "=H" + str(self.RowNumber) + "*K" + str(self.RowNumber)]

            elif self.AreaType == "BLANK":
                row = []
                for i in range(12):
                    row.append("")
                return row

            else:
                if self.Exclusion <= 0:
                    exclusion = "--"
                else:
                    exclusion = self.Exclusion
                if self.Tenant <= 0:
                    tenant = ""
                else:
                    tenant = self.Tenant
                if self.Ancillary <= 0:
                    ancillary = ""
                else:
                    ancillary = self.Ancillary

                return [self.Level,
                        "",
                        exclusion,
                        "",
                        self.Name,
                        tenant,
                        ancillary,
                        "=F" + str(self.RowNumber) + "+G" +
                        str(self.RowNumber),
                        "",
                        "",
                        "=D$" + str(lastRow) + "/H$" + str(lastRow),
                        "=H" + str(self.RowNumber) + "*K" + str(self.RowNumber)
                        ]
        else:
            return False

    @staticmethod
    def Header(method):
        if method == "A":
            header = [["" for i in range(18)]]
            header.append(list("ABCDEFGHIJKLMNOPQR"))
            header.append([
                # A
                "Input",
                # B
                "Input",
                # C
                "Input & ID",
                # D
                "'= B - C",
                # E
                "Input",
                # F
                "Input & ID",
                # G
                "Input & ID",
                # H
                "'= F + G",
                # I
                "Input & ID",
                # J
                "'= H + I",
                # K
                "Input & ID",
                # L
                "'= D - J - K",
                # M
                "'= (J + L) / J",
                # N
                "'= H * M",
                # O
                "'= (I * M) + K",
                # P
                "'= ΣD / (ΣH - ΣO)",
                # Q
                "'= N * P",
                # R
                "'= Q / H"
            ])

            header.append([
                # A
                "Floor Level",
                # B
                "Boundary Area (IPMS 2)",
                # C
                "Rentable Exclusions",
                # D
                "Floor Rentable Area",
                # E
                "Space ID",
                # F
                "Tenant Area (IPMS 3)",
                # G
                "Tenant Ancillary Area",
                # H
                "Occupant Area",
                # I
                "Building Amenity Area",
                # J
                "Floor Usable Area",
                # K
                "Building Service Area",
                # L
                "Floor Service Area",
                # M
                "Floor Allocation Ratio",
                # N
                "Floor Allocation",
                # O
                "Building Service & Amenity Area",
                # P
                "Builing Allocation Ratio",
                # Q
                "Rentable Area",
                # R
                "Load Factor A"
                ])
            header.append(["" for i in range(18)])

        elif method == "B":
            header = [["" for i in range(12)]]
            header.append(list("ABCDEFGHIJKL"))
            header.append([
                # A
                "Input",
                # B
                "Input",
                # C
                "Input & ID",
                # D
                "'= B - C",
                # E
                "Input",
                # F
                "Input & ID",
                # G
                "Input & ID",
                # H
                "'= F + G",
                # I
                "Input",
                # J
                "'= D - H - I",
                # K
                "'= ΣD / ΣH",
                # L
                "'= H * K"])
            header.append([
                # A
                "Floor Level",
                # B
                "Boundary Area (IPMS 2)",
                # C
                "Rentable Exclusions",
                # D
                "Floor Rentable Area",
                # E
                "Space ID",
                # F
                "Tenant Area",
                # G
                "Tenant Ancillary Area",
                # H
                "Occupant Area",
                # I
                "Base Building Circulation",
                # J
                "Service & Amenity Area",
                # K
                "Load Factor B",
                # L
                "Rentable Area"])
            header.append(["" for i in range(12)])

        else:
            header = False
        return header

    @staticmethod
    def GrandTotal(method, firstRow, row):
        if method == "A":
            # blankRow = ["" for i in range(18)]
            totalRow = [
                # A
                "Building Totals (Σ)",
                # B
                "=Sum(B" + str(firstRow) + ":B" + str(row - 1) + ")",
                # C
                "=SUBTOTAL(9,C" + str(firstRow) + ":C" + str(row - 1) + ")",
                # D
                "=Sum(D" + str(firstRow) + ":D" + str(row - 1) + ")",
                # E
                "",
                # F
                "=SUBTOTAL(9,F" + str(firstRow) + ":F" + str(row - 1) + ")",
                # G
                "=SUBTOTAL(9,G" + str(firstRow) + ":G" + str(row - 1) + ")",
                # H
                "=F" + str(row) + "+G" + str(row),
                # I
                "=SUBTOTAL(9, I" + str(firstRow) + ":I" + str(row - 1) + ")",
                # J
                "=H" + str(row) + "+I" + str(row),
                # K
                "=SUBTOTAL(9, K" + str(firstRow) + ":K" + str(row - 1) + ")",
                # L
                "=D" + str(row) + "-J" + str(row) + "-K" + str(row),
                # M
                "=(J" + str(row) + "+L" + str(row) + ")/J" + str(row),
                # N
                "=H" + str(row) + "*M" + str(row),
                # O
                "=I" + str(row) + "*M" + str(row) + "+K" + str(row),
                # P
                "=D$" + str(row) + "/(D$" + str(row) + "-O$" + \
                str(row) + ")",
                # Q
                "=N" + str(row) + "*P" + str(row),
                # R
                "=Q" + str(row) + "/H" + str(row)]
        elif method == "B":
            blankRow = ["" for i in range(12)]
            totalRow = [
                # A
                "Building Totals (Σ)",
                # B
                "=Sum(B" + str(firstRow) + ":B" + str(row - 1) + ")",
                # C
                "=SUBTOTAL(9,C" + str(firstRow) + ":C" + str(row - 1) + ")",
                # D
                "=Sum(D" + str(firstRow) + ":D" + str(row - 1) + ")",
                # E
                "",
                # F
                "=SUBTOTAL(9,F" + str(firstRow) + ":F" + str(row - 1) + ")",
                # G
                "=SUBTOTAL(9,G" + str(firstRow) + ":G" + str(row - 1) + ")",
                # H
                "=F" + str(row) + "+G" + str(row),
                # I
                "=Sum(I" + str(firstRow) + ":I" + str(row - 1) + ")",
                # J
                "=D" + str(row) + "-H" + str(row) + "-I" + str(row),
                # K
                "=D$" + str(row) + "/H$" + str(row),
                # L
                "=H" + str(row) + "*K" + str(row)]
        else:
            # totalRow = False
            blankRow = False
        return [totalRow]

doc = HOST_APP.doc

areaSchemes = DB.FilteredElementCollector(doc)\
                .OfClass(DB.AreaScheme)\
                .ToElements()

calcSchemes = {}
for areaScheme in areaSchemes:
    if (areaScheme.Name not in ["Gross Building", "Load Capacity"] and
        not re.search(r"Test Fit.*", areaScheme.Name)):
        calcSchemes[areaScheme.Name] = areaScheme

methods = ["A", "B"]
scheme = None
scheme, methodSwitch = forms.CommandSwitchWindow.show(calcSchemes.keys(), 
    message="Select a scheme to export",
    switches={"Method B": True})

if scheme is None:
    script.exit()

if methodSwitch["Method B"]:
    method = "B"
else:
    method = "A"

# Collect all areas in the model
# areas = DB.FilteredElementCollector(doc).WherePasses(areaFilter).ToElements()
levels = DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
levels = sorted(levels, key = lambda x: x.Elevation)
levels = [level.Name.ToString() for level in levels]

validAreaTypes = set(["major vertical penetrations",
                      "major vertical penetration",
                      "parking area",
                      "parking",
                      "occupant storage area",
                      "occupant storage",
                      "storage",
                      "tenant area",
                      "tenant ancillary area",
                      "floor service area",
                      "unenclosed building feature",
                      "building amenity area",
                      "building service area",
                      "inter-building amenity area",
                      "inter-building service area",
                      "base building circulation"])
invalidAreaTypes = set()
baseBuildingCirculationAlert = False

# !! Create the rows for the excel export !! #
# loop through area schemes
outList = []
areaSchemes = [areaScheme]
methods = [method]
for areaScheme, method in zip(areaSchemes, methods):
    schemeAreas = []
    # get excel header and setup the row numbering
    schemeRows = areaCalcArea.Header(method)
    row = 6
    firstRow = row

    areaFilter = DB.AreaFilter()
    areas = DB.FilteredElementCollector(doc)\
              .WherePasses(areaFilter)\
              .Where(lambda x: x.AreaScheme.Id.Equals(areaScheme.Id))

    areasToDelete = []
    redundantAreas = []
    for area in areas:
        if (area.Area == 0 and area.Location == None):
            # print("Not Placed", area.Number.ToString())
            areasToDelete.append(area)
        elif area.Area == 0:
            # print("Redundant", area.Number.ToString())
            redundantAreas.append(
                area.Level.Name + ": " +
                area.Number.ToString() + " " +
                area.GetParameters("Name")[0].AsString()
                )
    
    forms.alert_ifnot(
        not redundantAreas,
        "Remove redundant areas from the area scheme and run again:\n\n  " +
        "\n  ".join(redundantAreas),
        title="Area Calc Export",
        exitscript=True
    )

    if forms.alert_ifnot(
        not areasToDelete,
        "There are " + str(len(areasToDelete)) +
        " unplaced areas in the Area Scheme. Would you like to delete them " +
        "from the model?",
        ok=False,
        yes=True,
        no=True,
        title="Area Calc Export"
    ):
        with revit.Transaction("Delete unplaced areas"):
            revit.delete.delete_elements(areasToDelete)

    # loop through the levels
    for level in levels:
        # start building a list of area objects that we can later turn
        # into rows
        levelAreas = []
        # build a level total object
        levelTotal = areaCalcArea(level, "Total", level)
        # add a high sort number to place it at the end of the list later
        levelTotal.Sort = 4
        levelBlank = (areaCalcArea("", "BLANK", ""))
        levelBlank.Sort = 5

        # loop through all the areas
        for area in areas:
            # make sure the level matches the area scheme and level name
            if (area.AreaScheme.Name.ToString() == areaScheme.Name
                and area.Level.Name.ToString() == level):
                # print(area, level)
                # try:
                    # create an area object for the matching area
                    parsedArea = areaCalcArea()
                    parsedArea.ByRevitArea(method, area)
                    # add it to the area total for the level
                    levelTotal.AddAreas(parsedArea)
                    levelBlank.AddAreas(parsedArea)
                    # add the matched area to the list of areas
                    levelAreas.append(parsedArea)
                    if parsedArea.AreaType.lower() not in validAreaTypes:
                        invalidAreaTypes.add(parsedArea.AreaType)
                    if parsedArea.BaseBuildingCirculationAlert:
                        baseBuildingCirculationAlert = True
                # except Exception as e:
                #     print("Error: " + str(e))
        # Distribute ancillary areas
        if (levelAreas and method == "B"):
            # calculate tenant floor area percentage
            for area in levelAreas:
                if area.Tenant > 0:
                    area.TenantAreaRatio = area.Area / levelTotal.Tenant
                    # print(area.Name +
                    #       " " + str(area.Area) +
                    #       "/" + str(levelTotal.Tenant) +
                    #       " = " + str(area.TenantAreaRatio))
                    area.Ancillary = levelTotal.Ancillary * area.TenantAreaRatio

        # add all the areas from the level to the list
        levelAreas.append(levelTotal)
        levelAreas.append(levelBlank)
        # create a dictionary of area names so we can merge areas
        areaNames = set([area.Name for area in levelAreas])
        areaNames = list(areaNames)
        joinedAreas = dict.fromkeys(areaNames)
        # loop through the dictionary area and merge any areas that
        # have the same name
        for name in joinedAreas.keys():
            for area in levelAreas:
                if name == area.Name:
                    if joinedAreas[name] is None:
                        joinedAreas[name] = area
                    else:
                        (joinedAreas[name]).AddAreas(area)
        # sort all the joined areas to get them in the order BOMA expects
        joinedAreas = sorted(joinedAreas.values(), key = lambda x: x.Name)
        joinedAreas.sort(key = lambda x: x.Sort)
        sectionStartRow = row
        # remove areas that don't get listed out
        itemizedAreas = []
        for num, area in enumerate(joinedAreas):
            areaIn = area.ItemizeCheck()
            if area.ItemizeCheck():
                itemizedAreas.append(area)
        for num, area in enumerate(itemizedAreas):
            (itemizedAreas[num]).RowNumber = row
            sectionEndRow = row - 1
            row = row + 1
            if area.AreaType == "Total":
                (itemizedAreas[num]).SectionStartRow = sectionStartRow
                sectionStartRow = row
        for num, area in enumerate(itemizedAreas):
            (itemizedAreas[num]).SectionEndRow = sectionEndRow
        schemeAreas.extend(itemizedAreas)
        #% End Level Loop %#

    # Go through all the collected areas and build the rows for the excel doc
    for area in schemeAreas:
        schemeRows.append(area.MakeRow(method, row))

    # add the grand total loop
    schemeRows.extend(areaCalcArea.GrandTotal(method, firstRow, row))
    outList.append(schemeRows)
#% End Scheme Loop %#
if invalidAreaTypes:
    message = "Please use standard BOMA 2017 Space Classifications. "\
        + "The following area types are invalid:\n"
    for areaType in invalidAreaTypes:
        message = message + "\n" + areaType
    message = message + "\n\nRefer to table in "\
        + "S:\\03 AREACALC\\BOMA 2017\\WHA - Area Fill Color Scheme for BOMA "\
        + " 2017.pdf for allowable classifications."
    forms.alert(message, title="Area Calc Export")

if baseBuildingCirculationAlert:
    message = "'Scheme A' does not allow for Space Classification of "\
        + "'Base Building Circulation'"
    forms.alert(message, title="Area Calc Export")

# Generate the excel file path based of of the project folder of the revit
# file
if doc.IsFamilyDocument:
    forms.alert("Area Calcs cannot be calculated for Revit families. " +
                "Run this command in a project that has been setup for " +
                "area calcs.",
                title="Area Calc Export",
                exitscript=True)
if doc.IsWorkshared:
    revitModelPath = doc.GetWorksharingCentralModelPath()
    revitModelPath = DB.ModelPathUtils\
        .ConvertModelPathToUserVisiblePath(revitModelPath)
else:
    revitModelPath = doc.PathName
if not revitModelPath:
    forms.alert("Please save the model and try again.",
        title="Area Calc Export",
        exitscript=True)
uncReg = re.compile(re.escape("\\\\wha-server02\\projects"),
    re.IGNORECASE)
revitModelPath = uncReg.sub("x:", revitModelPath)
projectNumber = GetParameterFromProjectInfo(doc,
    "Project Number").replace(" ", "")

excelPath = path.dirname(revitModelPath) + "\\" \
    + str(projectNumber) + " Area Calc Export.xlsx"

# print(excelPath)
workbook = OpenWorkbook(excelPath)
worksheet = workbook.Worksheets[1]
worksheet.name = "Method " + method

columnWidths = [12.0] * len(outList[0][1])
if method == "B":
    columnWidths[0] = 16.0
    columnWidths[4] = 32.0

if method == "A":
    round2Columns = ["B", "C", "D", "F", "G", "H", "I", "J", "L",
                     "N", "O", "Q"]
    round6Columns = ["M", "P", "R"]
else:
    round2Columns = ["B", "C", "D", "F", "G", "H", "I", "J", "L"]
    round6Columns = ["K"]

for i, width in enumerate(columnWidths):
    letter = ascii_uppercase[i]
    worksheet.Columns(letter).ColumnWidth = width
    if letter in round2Columns:
        worksheet.Columns(letter).NumberFormat = "0.00"
    elif letter in round6Columns:
        worksheet.Columns(letter).NumberFormat = "0.000000"

green = int("0xDAEFE2", base=16)
grey = int("0xC0C0C0", base=16)
for x, row in enumerate(outList[0]):
    cellRange = worksheet.Range(
        worksheet.Cells(x + 1, 1),
        worksheet.Cells(x + 1, len(row))
    )
    if x <= 3:
        cellRange.HorizontalAlignment = -4108
        cellRange.VerticalAlignment = -4108
    if x == 1:
        cellRange.Font.Bold = True
    if x == 3:
        cellRange.Interior.Color = grey
        cellRange.WrapText = True
        cellRange.Font.Bold = True
        cellRange.Borders.LineStyle = 1
        cellRange.Borders.Weight = 4
        interiorVerticals = cellRange.Borders.Item[11]
        interiorVerticals.Weight = 2
    if "Total" in row[0]:
        cellRange.Interior.Color = green
        cellRange.Font.Bold = True
        bottom = cellRange.Borders.Item[9]
        bottom.LineStyle = 1
        bottom.Weight = 4
    for y, cell in enumerate(row):
        worksheet.Cells[x + 1, y + 1].Value = cell

