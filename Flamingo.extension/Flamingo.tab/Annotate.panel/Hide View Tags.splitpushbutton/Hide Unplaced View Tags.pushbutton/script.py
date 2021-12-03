# -*- coding: utf-8 -*-
from pyrevit import HOST_APP
from flamingo.revit import HideUnplacedViewTags


if __name__ == "__main__":
    doc = HOST_APP.doc

    # Get the current view
    view = doc.ActiveView
    HideUnplacedViewTags(viewId=view.Id, doc=doc)


    # Filter through all elements on the current view to find view tags.
