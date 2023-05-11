from pyrevit import DOCS, HOST_APP, script


OUTPUT = script.get_output()
if __name__ == "__main__":
    documents = HOST_APP.docs
    uidoc = HOST_APP.uidoc
    # openViews = uidoc.GetOpenUIViews()
    # for uiView in openViews:
    #     print(uiView.ViewId)
    for doc in documents:
        title = doc.Title
        try:
            OUTPUT.print_md("## {}".format(title))
            OUTPUT.print_md("IsModifiable: {}".format(str(doc.IsModifiable)))
            OUTPUT.print_md("IsModified: {}".format(str(doc.IsModified)))
            OUTPUT.print_md("IsReadOnly: {}".format(str(doc.IsReadOnly)))
            OUTPUT.print_md("IsLinked: {}".format(str(doc.IsLinked)))
            OUTPUT.print_md("ActiveView: {}".format(str(doc.ActiveView)))
            OUTPUT.print_md("PathName: {}".format(doc.PathName))
            if doc.ActiveView:
                OUTPUT.print_md("Active project")
            else:
                try:
                    doc.Close(True)
                    OUTPUT.print_md("Foreground? Save with close")
                except Exception as e:
                    OUTPUT.print_md("Close error: {}".format(e))
                    doc.Close(False)
                    OUTPUT.print_md("Background? Close without save")
                # OUTPUT.print_md("Closed".format(title))
        except Exception as e:
            OUTPUT.print_md(str(e))