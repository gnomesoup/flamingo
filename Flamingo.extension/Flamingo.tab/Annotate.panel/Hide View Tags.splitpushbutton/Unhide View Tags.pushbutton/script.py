from pyrevit.revit import HOST_APP
from flamingo.revit import UnhideViewTags
import clr

clr.AddReference("System")
from System.Collections.Generic import List

if __name__ == "__main__":
    doc = HOST_APP.doc
    activeView = doc.ActiveView

    UnhideViewTags(view=activeView, doc=doc)

