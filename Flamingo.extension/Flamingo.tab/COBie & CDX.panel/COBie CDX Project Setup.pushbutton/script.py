from Autodesk.Revit import DB
from pyrevit import DOCS, forms
from flamingo.ensure import EnsureLibraryDoc


# forms.inform_wip()

COBIE_CDX_TEMPLATE_NAME = "COBie-CDX Templates"

if __name__ == "__main__":
    doc = DOCS.doc
    libraryDoc = EnsureLibraryDoc(COBIE_CDX_TEMPLATE_NAME)

    libraryDoc.Close(False)

