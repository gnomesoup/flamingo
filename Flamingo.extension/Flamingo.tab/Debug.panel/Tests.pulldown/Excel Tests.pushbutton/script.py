from flamingo import excel
from pyrevit import script, EXEC_PARAMS
from pyrevit.interop import xl

OUTPUT = script.get_output()
LOGGER = script.get_logger()

if __name__ == "__main__":
    # LOGGER.set_debug_mode()
    LOGGER.debug("Starting Excel open workbook test")

    try:
        # create a temporary workbook for testing
        filePath = script.get_instance_data_file("test.xlsx")
        xl.dump(filePath, {"Sheet1": [["Header1", "Header2"], [1, 2], [3, 4], [5, 6]]})
    except Exception as e:
        OUTPUT.print_md(":cross_mark: Excel test setup failed.")
        LOGGER.exception("Error creating test Excel file: {}".format(e))
        script.exit()

    try:
        workbook = excel.OpenWorkbook(filePath, createNew=True)
        OUTPUT.print_md(":white_heavy_check_mark: Excel open workbook test succeeded.")
    except Exception as e:
        OUTPUT.print_md(":cross_mark: Excel open workbook test failed.")
        LOGGER.exception("Error opening workbook: {}".format(e))
        script.exit()
    try:
        workbook = excel.OpenWorkbook(filePath, createNew=False)
        data = excel.GetWorksheetData(excel.GetWorksheetByName(workbook, "Sheet1"))
        validDict = {
            "rowCount": 4,
            1: {
                "data": ["Header1", "Header2"],
                "meta": {"Sort Name": None, "Row Number": 1, "Sort Number": 0},
            },
            2: {
                "data": ["1", "2"],
                "meta": {"Sort Name": None, "Row Number": 2, "Sort Number": 0},
            },
            3: {
                "data": ["3", "4"],
                "meta": {"Sort Name": None, "Row Number": 3, "Sort Number": 0},
            },
            4: {
                "data": ["5", "6"],
                "meta": {"Sort Name": None, "Row Number": 4, "Sort Number": 0},
            },
            "columnCount": 2,
        }

        print("--- Retrieved Data ---")
        print(data)
        print("--- Valid Data ---")
        print(validDict)
        assert data == validDict, "Worksheet data does not match expected data"
        LOGGER.success(
            ":white_heavy_check_mark: Excel get worksheet data test succeeded."
        )
    except Exception as e:
        OUTPUT.print_md(":cross_mark: Excel get worksheet data test failed.")
        LOGGER.exception("Error getting worksheet data: {}".format(e))

    try:
        excel.CloseWorkbook(workbook)
        LOGGER.success(":white_heavy_check_mark: Excel close workbook test succeeded.")
    except Exception as e:
        OUTPUT.print_md(":cross_mark: Excel close workbook test failed.")
        LOGGER.exception("Error closing workbook: {}".format(e))

    # LOGGER.reset_level()
