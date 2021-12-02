from pyrevit import HOST_APP, forms, script, clr, DB
from pyrevit.coreutils import ribbon
from flamingo.excel import OpenWorkbook
import System

clr.ImportExtensions(System.Linq)

def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    """
    Args:
        script_cmp: script component that contains info on this script
        ui_button_cmp: this is the UI button component
        __rvt__: Revit UIApplication

    Returns:
        bool: Return True if successful, False if not
    """
    
    button_icon = script_cmp.get_bundle_file(
        'export.png' if True else 'import.png'
    )
    ui_button_cmp.set_icon(button_icon, icon_size=ribbon.ICON_LARGE)
    ui_button_cmp.set_title(
        "Edit in \nExcel" if True else "Return from \nExcel"
    )

if __name__ == "__main__":
    doc = HOST_APP.doc

    workbookPath = script.get_document_data_file(
            file_id="COBieSheets",
            file_ext="xlsx"
        )
    print("workbookPath = {}".format(workbookPath))
    workbook = OpenWorkbook(workbookPath)
    cobieSchedules = DB.FilteredElementCollector(doc)\
        .OfClass(DB.ViewSchedule)\
        .Where(lambda x: x.Name.startswith("COBie"))

    for i, schedule in enumerate(cobieSchedules):
        if i == 0:
            worksheet = workbook.ActiveSheet
        else:
            worksheet = workbook.Worksheets.Add(
                After=workbook.ActiveSheet
            )
        worksheet.name = schedule.Name[0:30]

    script.exit()
    button = script.get_button()
    config = script.get_config()
    buttonIcon = None
    if config.has_option("Excel"):
        if config.get_option("Excel") == True:
            buttonIcon = script.get_bundle_file(
                'import.png'
            )
            buttonText = "Return from \nExcel"
    if not buttonIcon:
        buttonIcon = script.get_bundle_file(
            'import.png'
        )
        buttonText = "Edit in \nExcel"

    button.set_icon(buttonIcon, icon_size=ribbon.ICON_LARGE)
    button.set_title(buttonText)
    forms.inform_wip()