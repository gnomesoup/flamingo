from pyrevit import HOST_APP, revit, script

LOGGER = script.get_logger()
OUTPUT = script.get_output()

if __name__ == "__main__":
    doc = HOST_APP.doc
    
    selection = revit.get_selection()
    if not selection:
        selection = revit.pick_elements()
        
    if not selection:
        script.exit()
        
    print(selection)