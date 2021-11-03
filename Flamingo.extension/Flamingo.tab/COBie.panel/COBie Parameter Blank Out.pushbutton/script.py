from pyrevit import clr, DB, HOST_APP, forms, revit, script
import System
clr.ImportExtensions(System.Linq)

def getSymbolId(element):
    try:
        if element.LookupParameter("COBie").AsInteger() == 1:
            symbolId = element.Symbol.Id
        else: 
            symbolId = None
    except:
        symbolId = None
    return symbolId
        

if __name__ == "__main__":
    
    doc = HOST_APP.doc
    schedules = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)\
        .WhereElementIsNotElementType()\
        .Where(lambda x: x.Name.startswith("COBie"))
    
    scheduledParameters = {}
    groupedParameterNames = {}
    for schedule in schedules:
        # print(schedule.Name)

        fieldOrder = schedule.Definition.GetFieldOrder()
        fields = [
            schedule.Definition.GetField(fieldId) for fieldId in fieldOrder
            if fieldId is not None
        ]
        parameterIds = [field.ParameterId for field in fields]
        parameterNames = []
        for parameterId in parameterIds:
            if parameterId is not None and parameterId.IntegerValue > 0:
                parameter = doc.GetElement(parameterId)
                # print("\t{}: {}".format(
                #     parameterId,
                #     parameter.Name if parameter is not None else type(parameter)
                # ))
                scheduledParameters[parameter.Name] = parameterId
                parameterNames.append(parameter.Name)
        groupedParameterNames[schedule.Name] = parameterNames
    parameterNamesOut = forms.SelectFromList.show(
        groupedParameterNames,
        multiselect=True,
        group_selector_title='COBie Schedule',
        button_name='Select Parameters'
    )

    bindings = doc.ParameterBindings
    with revit.Transaction("COBie Parameter Blank Out"):
        for parameterName in parameterNamesOut:
            parameterId = scheduledParameters[parameterName]
            parameter = doc.GetElement(parameterId)
            definition = parameter.GetDefinition()
            parameterBinding = bindings.get_Item(definition)
            parameterValueProvider = DB.ParameterValueProvider(
                parameterId
            )
            filterStringRule = DB.FilterStringRule(
                parameterValueProvider,
                DB.FilterStringGreater(),
                "",
                True
            )
            elementParameterFilter = DB.ElementParameterFilter(
                filterStringRule
            )
            if type(parameterBinding) is DB.TypeBinding:
                collector = DB.FilteredElementCollector(doc)\
                    .WhereElementIsElementType()
            else:
                collector = DB.FilteredElementCollector(doc)\
                    .WhereElementIsNotElementType()
            elements = collector.WherePasses(elementParameterFilter).ToElements()

            for element in elements:
                elementParameter = element.get_Parameter(parameter.GuidValue)
                elementParameter.Set("")