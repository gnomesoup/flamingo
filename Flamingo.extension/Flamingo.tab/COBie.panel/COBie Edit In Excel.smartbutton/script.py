from pyrevit import forms, script 
from pyrevit.coreutils import ribbon

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
    button = script.get_button()
    config = script.get_config()
    buttonIcon = None
    if config.has_option("Excel"):
        if config.get_option("Excel") == True:
            buttonIcon = script.get_bundle_file(
                'import.png'
            )
    if not buttonIcon:
        buttonIcon = script.get_bundle_file(
            'import.png'
        )

    button.set_icon(buttonIcon, icon_size=ribbon.ICON_LARGE)
    button.set_title("Test")
    forms.inform_wip()