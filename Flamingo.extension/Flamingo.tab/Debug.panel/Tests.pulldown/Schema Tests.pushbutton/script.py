from flamingo import extensible_storage
from pyrevit import forms, HOST_APP, revit, script

OUTPUT = script.get_output()
LOGGER = script.get_logger()

if __name__ == "__main__":
    doc = HOST_APP.doc
    # LOGGER.set_debug_mode()
    try:
        with revit.Transaction("Schema Test") as t:
            try:
                schema = extensible_storage.GetFlamingoSchema(doc)
                projectInfo = doc.ProjectInformation
                entity = doc.ProjectInformation.GetEntity(schema)
                version = entity.Get[str](schema.GetField("Version"))
                OUTPUT.print_md(
                    ":white_heavy_check_mark: GetFlamingoSchema succeeded. "
                    "{} with version {}created/found.".format(
                        schema.SchemaName, version
                    )
                )
            except Exception as e:
                OUTPUT.print_md(":cross_mark: GetFlamingoSchema failed .\n{}".format(e))
                script.exit()
            try:
                data = extensible_storage.SetSchemaMapData(
                    schema,
                    "Settings",
                    projectInfo,
                    {"Key1": "Value1", "Key2": "Value2"},
                )
                assert data["Key1"] == "Value1"
                assert data["Key2"] == "Value2"
                OUTPUT.print_md(":white_heavy_check_mark: SetSchemaMapData succeeded.")
            except Exception as e:
                OUTPUT.print_md(":cross_mark: SetSchemaMapData failed .\n{}".format(e))
            try:
                extensible_storage.SetFlamingoSetting("TestSetting", "TestValue", doc)
                OUTPUT.print_md(
                    ":white_heavy_check_mark: SetFlamingoSetting succeeded."
                )
            except Exception as e:
                OUTPUT.print_md(
                    ":cross_mark: SetFlamingoSetting failed .\n{}".format(e)
                )
            try:
                value = extensible_storage.GetFlamingoSetting("Key1", doc=doc)
                assert value == "Value1"
                value = extensible_storage.GetFlamingoSetting("TestSetting", doc=doc)
                assert value == "TestValue"
                defaultValue = extensible_storage.GetFlamingoSetting(
                    "NonExistingSetting", default="DefaultValue", doc=doc
                )
                assert defaultValue == "DefaultValue"
                OUTPUT.print_md(
                    ":white_heavy_check_mark: GetFlamingoSetting succeeded."
                )
            except Exception as e:
                OUTPUT.print_md(
                    ":cross_mark: GetFlamingoSetting failed .\n{}".format(e)
                )
    except Exception as e:
        LOGGER.exception("Error during schema test:\n{}".format(e))
    # finally:
    #     LOGGER.reset_level()
