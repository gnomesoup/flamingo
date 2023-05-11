from Autodesk.Revit import DB
from pyrevit import forms
from pyrevit import HOST_APP, revit, script
from System import Type
from System.Collections.Generic import List

LOGGER = script.get_logger()
OUTPUT = script.get_output()

if __name__ == "__main__":
    doc = HOST_APP.doc

    classes = List[Type]()
    classes.Add(DB.ParameterElement)
    classes.Add(DB.SharedParameterElement)
    parameters = (
        DB.FilteredElementCollector(doc)
        .WherePasses(DB.ElementMulticlassFilter(classes))
        .ToElements()
    )
    parametersByName = {
        "{} (Id:{})".format(parameter.Name, parameter.Id): parameter
        for parameter in parameters
    }
    selectedParameterName = forms.SelectFromList.show(
        context=sorted(parametersByName.keys()), title="Select prototype parameter"
    )
    if not selectedParameterName:
        script.exit()
    selectedParameter = parametersByName[selectedParameterName]
    parameterBindings = doc.ParameterBindings
    binding = parameterBindings.get_Item(selectedParameter.GetDefinition())
    potentialParameters = [
        parameter for parameter in parameters
        if type(parameterBindings.get_Item(parameter.GetDefinition())) == type(binding)
    ]
    potentialParametersByName = {
        "{} (Id:{})".format(parameter.Name, parameter.Id): parameter
        for parameter in potentialParameters
    }
    toMatchParameterNames = forms.SelectFromList.show(
        context=sorted(potentialParametersByName.keys()),
        title="Select parameters to be matched",
        multiselect=True
    )
    if not toMatchParameterNames:
        script.exit()
    with revit.Transaction("Match parameter categories"):
        for parameterName in toMatchParameterNames:
            toMatchParameter = potentialParametersByName[parameterName]
            parameterBindings.ReInsert(toMatchParameter.GetDefinition(), binding)

