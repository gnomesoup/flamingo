from flamingo.revit import GetElementIdValue
from pyrevit import script, HOST_APP
from System import Int64

OUTPUT = script.get_output()
LOGGER = script.get_logger()

if __name__ == "__main__":
    doc = HOST_APP.doc
    LOGGER.set_debug_mode()
    try:
        element = doc.ProjectInformation
        elementId = element.Id
        id_value = GetElementIdValue(elementId)
        LOGGER.info("Element ID Value: {}".format(id_value))
        LOGGER.info("type: {}".format(type(id_value)))
        assert (
            type(id_value) is int
            if HOST_APP.version < "2024"
            else type(id_value) is Int64
        )
        OUTPUT.print_md(
            ":white_heavy_check_mark: GetElementIdValue succeeded."
            "Element ID: {} has value: {}".format(elementId, id_value)
        )
    except Exception as e:
        OUTPUT.print_md(":cross_mark: GetElementIdValue failed.")
        LOGGER.exception("Error during GetElementIdValue: {}".format(e))
    try:
        LOGGER.info("Host App Version: {}".format(HOST_APP.version))
        id_value = GetElementIdValue(elementId, version=HOST_APP.version)
        assert (
            type(id_value) is int
            if HOST_APP.version < "2024"
            else type(id_value) is Int64
        )
        OUTPUT.print_md(
            ":white_heavy_check_mark: GetElementIdValue with passed version succeeded."
            "Element ID: {} has value: {}".format(elementId, id_value)
        )
    except Exception as e:
        OUTPUT.print_md(":cross_mark: GetElementIdValue failed.")
        LOGGER.exception("Error during GetElementIdValue: {}".format(e))
    LOGGER.reset_level()
