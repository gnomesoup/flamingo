from pyrevit import forms
from flamingo.revit import MarkModelTransmitted

centralPath = forms.pick_file("rvt")
print(centralPath)
MarkModelTransmitted(centralPath)