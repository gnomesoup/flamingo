from pyrevit import forms, script
import os
import re

regEx = re.compile(r"\d{4}\.r(vt|fa|te)$")

path = forms.pick_folder(
    title="Select folder to cleanup backups"
)
if not path:
    script.exit()

rvtFiles = [
    f for f in os.listdir(path)
    if re.search(r"\.r(vt|fa|te)$", f.lower())
]

backupFiles = [
    f for f in rvtFiles
    if regEx.search(f.lower())
]

if not backupFiles:
    forms.alert(
        "There are no backup files in this folder to remove.",
        exitscript=True
    )

removedCount = 0
errorCount = 0
for f in backupFiles:
    try:
        os.remove(os.path.join(path, f))
        removedCount += 1
    except Exception as e:
        errorCount += 1
        print(e)

if errorCount:
    msg = "Error removing {} file{}.".format(
        errorCount,
        "" if errorCount == 1 else "s"
    )
else:
    msg = "Removed {} backup file{}".format(
        removedCount,
        "" if removedCount == 1 else "s"
    )

forms.alert(msg)

# TODO If recursive???
# for root, directory, files in os.walk(path):

