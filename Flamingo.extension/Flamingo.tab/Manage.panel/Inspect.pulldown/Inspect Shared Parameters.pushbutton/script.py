from Autodesk.Revit import DB
import codecs
from pyrevit import forms, HOST_APP, script, revit
import re

OUTPUT = script.get_output()

if __name__ == "__main__":
    doc = HOST_APP.doc

    if script.EXEC_PARAMS.config_mode:
        sharedParameterPath = forms.pick_file(
            file_ext="txt", title="Select shared parameter file"
        )
    else:
        sharedParameterPath = revit.query.get_sharedparam_definition_file().Filename

    parameters = (
        DB.FilteredElementCollector(doc).OfClass(DB.SharedParameterElement).ToElements()
    )

    projectParameters = [parameter.Name for parameter in parameters]

    paramFileParams = []
    with codecs.open(sharedParameterPath, "r", encoding="UTF-16") as f:
        for line in f:
            m = re.search(r"^PARAM\t[^\t]+\t([^\t]+)\t", line)
            if m:
                paramFileParams.append(m.group(1))
    OUTPUT.print_md("# Shared Parameter File Path")
    OUTPUT.print_code(sharedParameterPath)
    OUTPUT.print_md("# Shared Parameters In Project")
    for param in projectParameters:
        OUTPUT.print_md(
            "{} {}".format(
                ":white_heavy_check_mark:"
                if param in paramFileParams
                else ":no_entry:",
                param,
            )
        )
