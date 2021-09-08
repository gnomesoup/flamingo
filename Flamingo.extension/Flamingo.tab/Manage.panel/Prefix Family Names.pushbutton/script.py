from pyrevit import DB, forms, HOST_APP, revit

if __name__ == "__main__":
    doc = HOST_APP.doc

    categoryBuiltIns = {
        DB.BuiltInCategory.OST_Casework: "CW_",
        DB.BuiltInCategory.OST_DuctTerminal: "DT_",
        DB.BuiltInCategory.OST_ElectricalEquipment: "EL_",
        DB.BuiltInCategory.OST_ElectricalFixtures: "EF_",
        DB.BuiltInCategory.OST_Entourage: "EN_",
        DB.BuiltInCategory.OST_Furniture: "FN_",
        DB.BuiltInCategory.OST_FurnitureSystems: "FS_",
        DB.BuiltInCategory.OST_GenericModel: "GM_",
        DB.BuiltInCategory.OST_LightingFixtures: "LF_",
        DB.BuiltInCategory.OST_MechanicalEquipment: "ME_",
        DB.BuiltInCategory.OST_Parking: "PK_",
        DB.BuiltInCategory.OST_Planting: "PL_",
        DB.BuiltInCategory.OST_PlumbingFixtures: "PL_",
        DB.BuiltInCategory.OST_RailingSupport: "RS_",
        DB.BuiltInCategory.OST_RailingTermination: "RT_",
        DB.BuiltInCategory.OST_Site: "SI_",
        DB.BuiltInCategory.OST_SpecialityEquipment: "SE_",
        DB.BuiltInCategory.OST_Sprinklers: "SP_",
        DB.BuiltInCategory.OST_StructConnections: "SC_",
    }

    families = DB.FilteredElementCollector(doc) \
        .OfClass(DB.Family)

    familiesToRename = []
    correctPrefixCount = 0
    for family in families:
        for category in categoryBuiltIns.keys():
            familyCategory = family.FamilyCategory
            if familyCategory.Id.IntegerValue == category.value__:
                prefix = categoryBuiltIns[category]
                if (family.Name).startswith(prefix):
                    correctPrefixCount += 1
                    break
                familiesToRename.append(
                    {
                        "family": family,
                        "newName": prefix + family.Name,
                    }
                )
                break

    if len(familiesToRename) < 1:
        familyString = "families" if correctPrefixCount > 1 else "family"
        message = "{} {} had the correct prefix. No" \
            " families required renaming.".format(
                correctPrefixCount,
                familyString
            )
        forms.alert(message)
    else:
        familyString = "families" if len(familiesToRename) > 1 else "family"
        message = "Would you like to correct the prefix on {} {}?".format(
            len(familiesToRename),
            familyString
        )
        if forms.alert(message, yes=True, no=True):
            with revit.Transaction("Rename families with category prefix"):
                for item in familiesToRename:
                    item['family'].Name = item['newName']
