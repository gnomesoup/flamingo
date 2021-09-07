from pyrevit import DB, forms, HOST_APP, revit

doc = HOST_APP.doc

categoryBuiltIns = {
    DB.BuiltInCategory.OST_Casework: "CW_",
    DB.BuiltInCategory.OST_ElectricalEquipment: "EL_",
    DB.BuiltInCategory.OST_ElectricalFixtures: "EF_",
    DB.BuiltInCategory.OST_Entourage: "EL_",
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
}

families = DB.FilteredElementCollector(doc) \
    .OfClass(DB.Family) \
    # .WhereElementIsElementType()

print("Unfiltered Count: {}\n".format(families.GetElementCount()))

with revit.Transaction("Rename Families By Category"):
    for family in families:
        for category in categoryBuiltIns.keys():
            familyCategory = family.FamilyCategory
            if familyCategory.Id.IntegerValue == category.value__:
                prefix = categoryBuiltIns[category]
                if (family.Name).startswith(prefix):
                    print(
                        "Family: {} :)".format(
                            family.Name,
                        )
                    )
                    continue
                newName =  prefix + family.Name
                print(
                    "Family: {} -> {}".format(
                        family.Name,
                        newName
                    )
                )
                family.Name = newName
                break
