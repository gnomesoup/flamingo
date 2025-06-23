from Autodesk.Revit import DB
from pyrevit import forms, HOST_APP, revit, script
from os import walk

LOGGER = script.get_logger()
OUTPUT = script.get_output()

if __name__ == "__main__":
    doc = HOST_APP.doc
    LOGGER.set_debug_mode()
    linkDirectory = forms.pick_folder()
    LOGGER.debug("Link directory: {}".format(linkDirectory))

    # Get all files in the directory
    linkPaths = []
    for (dirpath, dirnames, filenames) in walk(linkDirectory):
        # Check if the file has the extension .rvt
        for f in filenames:
            if f.endswith(".rvt"):
                linkPaths.append(f)

    linkPaths = forms.SelectFromList.show(context=linkPaths, multiselect=True)
    LOGGER.debug("Link paths: {}".format(linkPaths))

    # Place the links in the current document
    for linkPath in linkPaths:
        linkPath = linkDirectory + "\\" + linkPath
        LOGGER.debug("Link path: {}".format(linkPath))
        linkType = DB.RevitLinkType.Create(doc, linkPath)
        linkInstance = DB.RevitLinkInstance.Create(doc, linkType.ElementId)
        LOGGER.debug("Link instance: {}".format(linkInstance))
        
    LOGGER.reset_level()
