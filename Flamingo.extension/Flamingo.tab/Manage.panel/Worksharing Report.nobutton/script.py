from flamingo.worksharing import GetWorksharingReport
from glob import glob
import json
from os import path
from pyrevit import forms, script

if __name__ == "__main__":
    
    output = script.get_output()
    revitModelPath = forms.pick_file(
        file_ext="rvt",
        init_dir="N:\\",
        title="Pick a revit file to report"
    )
    if not revitModelPath:
        script.exit()
    revitModel, x = path.splitext(revitModelPath)
    backupFolderPath = revitModel + "_backup"
    if path.exists(backupFolderPath):
    #     output.print_md(listdir(backupFolderPath, ".slog"))
        slogPaths = glob(backupFolderPath + "/*.slog")
        if len(slogPaths) > 1:
            forms.alert(
                "Multiple slog files found. Creating report off of the fist file"
                " This model should be recentralized to cleanup the backup folder."
            )
        elif not slogPaths:
            forms.alert(
                "No worksharing log file could be found in the specified folder",
                exitscript=True
            )
        slogPath = slogPaths[0]
    else:
        forms.alert(
            "The workshare folder for the specified model does not exist.",
            exitscript=True
        )

    GetWorksharingReport(slogPath, output, printReport=True)