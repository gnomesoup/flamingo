from flamingo.revit import MarkModelTransmitted
from pyrevit import forms, script
from os import path, walk

LOGGER = script.get_logger()
OUTPUT = script.get_output()

if __name__ == "__main__":
    commandList = ["Selected model(s)", "All models in a directory"]
    command = forms.CommandSwitchWindow.show(commandList)

    if not command:
        script.exit()

    modelPaths = []
    if command == "Selected model(s)":
        modelPaths = forms.pick_file(file_ext="rvt", multi_file=True)
    elif command == "All models in a directory":
        folder = forms.pick_folder()
        if not folder:
            script.exit()
        # Get all Revit files in the selected directory and subdirectories
        for root, dirs, files in walk(folder):
            for file in files:
                if file.endswith(".rvt"):
                    # Append the full path to the list
                    modelPaths.append(path.join(root, file))
        OUTPUT.print_md("# Models to mark as transmitted")
        for modelPath in modelPaths:
            OUTPUT.print_md("- {}".format(modelPath))
    if not modelPaths:
        script.exit()

    if not forms.alert(
        "Are you sure you want to mark the {} model({}) as transmitted?".format(
            len(modelPaths), "" if len(modelPaths) == 1 else "s"
        ),
        yes=True,
        cancel=True,
    ):
        script.exit()
    OUTPUT.print_md("# Marking models as transmitted")
    for modelPath in modelPaths:
        try:
            MarkModelTransmitted(modelPath)
            OUTPUT.print_md("- :white_heavy_check_mark: {}".format(modelPath))
        except Exception as e:
            OUTPUT.print_md("- :warning: {}".format(modelPath))
