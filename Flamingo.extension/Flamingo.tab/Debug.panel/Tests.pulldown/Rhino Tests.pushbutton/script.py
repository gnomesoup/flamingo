from pyrevit import script, forms, clr

try: 
    clr.AddReference("RhinoCommon")
    clr.AddReference("RhinoInside.Revit")
    from flamingo import rhino
    from Rhino import RhinoDoc
except Exception as e:
    forms.alert(
        "Start Rhino from Rhino.Inside.Revit before running Rhino tests.",
        exitscript=True,
    )



OUTPUT = script.get_output()
LOGGER = script.get_logger()

if __name__ == "__main__":
    rhinoDoc = RhinoDoc.ActiveDoc
    try:
        version_exists = rhino.check_rhino_version("7.0")
        assert version_exists == True
        OUTPUT.print_md(":white_heavy_check_mark: Rhino version 7.0 check passed.")
        version_exists = rhino.check_rhino_version("20.0")
        assert version_exists == False
        OUTPUT.print_md(
            ":white_heavy_check_mark: Rhino future version not available check passed."
        )
    except Exception as e:
        OUTPUT.print_md(":cross_mark: Rhino version check failed.")
        LOGGER.exception("Error during Rhino version check: {}".format(e))
    try:
        layer_name = "Default"
        layerIndex = rhino.FindOrAddRhinoLayer(layer_name, rhinoDoc)
        assert layerIndex is not None
        layer = rhinoDoc.Layers[layerIndex]
        assert layer.Name == layer_name
        OUTPUT.print_md(
            ":white_heavy_check_mark: FindOrAddRhinoLayer find layer succeeded. "
            "Layer '{}' found.".format(layer_name)
        )
        layer_name = "Test Layer"
        layerIndex = rhino.FindOrAddRhinoLayer(layer_name, rhinoDoc)
        assert layerIndex is not None
        layer = rhinoDoc.Layers[layerIndex]
        assert layer.Name == layer_name
        OUTPUT.print_md(
            ":white_heavy_check_mark: FindOrAddRhinoLayer new layer succeeded. "
            "Layer '{}' found or added.".format(layer_name)
        )
        layer_name = "Test Layer::Child Layer"
        layerIndex = rhino.FindOrAddRhinoLayer(layer_name, rhinoDoc)
        assert layerIndex is not None
        layer = rhinoDoc.Layers[layerIndex]
        assert layer.FullPath == layer_name
        OUTPUT.print_md(
            ":white_heavy_check_mark: FindOrAddRhinoLayer new sublayer succeeded. "
            "Layer '{}' found or added.".format(layer_name)
        )
    except Exception as e:
        OUTPUT.print_md(":cross_mark: FindOrAddRhinoLayer failed.")
        LOGGER.exception("Error during FindOrAddRhinoLayer: {}".format(e))
