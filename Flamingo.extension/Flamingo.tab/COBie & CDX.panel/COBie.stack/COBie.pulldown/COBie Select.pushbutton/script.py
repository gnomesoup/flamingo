# from Autodesk.Revit import DB
from Autodesk.Revit import DB
from flamingo.cobie import COBieUncheckAll, COBieParameterVaryByGroup
from flamingo.revit import GetAllElementsInModelGroups
from pyrevit import clr, forms, HOST_APP, script
from pyrevit.revit import query
import System

clr.ImportExtensions(System.Linq)


class TypeElementOption(forms.TemplateListItem):
    """Element wrapper for SelectFromList"""

    def __init__(self, orig_item, checked=False, checkable=True, name_attr=None):
        super(TypeElementOption, self).__init__(
            orig_item, checked, checkable, name_attr
        )

    @property
    def name(self):
        """Family and Type name."""
        parameter = self.item.get_Parameter(
            DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM
        )
        value = parameter.AsValueString()
        if not value:
            value = query.get_name(self.item)
            print(output.linkify(self.item.Id))
        return value

    @property
    def category(self):
        """Category"""
        category = query.get_category(self.item)
        print(category.Name)
        return category.Name


if __name__ == "__main__":
    doc = HOST_APP.doc
    output = script.get_output()

    try:
        cobieTypeParameter = (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.SharedParameterElement)
            .Where(lambda x: x.Name == "COBie.Type")
            .First()
        )
    except Exception:
        forms.alert(
            "Please setup the project for COBie using the BIM Interoperability "
            "Tools before running this command.",
            exitscript=True,
        )

    # Attempted to uncheck the parameter myself but still run into group issues
    # unsetElements = COBieUncheckAll(typeView, componentView, doc=doc)

    # Attempting to create a new parameter as a placeholder for COBie checks
    # Copy over data from COBie parameter to placeholder
    # Remove COBie parameter from model
    # Add back in with group instances allowed to vary

    # COBieParameterVaryByGroup(typeView, doc)
    groupedElements = GetAllElementsInModelGroups(doc)
    cobieTypeElements = (
        DB.FilteredElementCollector(doc)
        .WherePasses(
            DB.ElementParameterFilter(
                DB.FilterIntegerRule(
                    DB.ParameterValueProvider(cobieTypeParameter.Id),
                    DB.FilterNumericGreaterorEqual(),
                    0
                )
            )
                # DB.ParameterFilterRuleFactory.CreateSharedParameterApplicableRule(
                #     "COBie.Type"
                # )
        )
        .ToElements()
    )
    print(len(cobieTypeElements))
    wrappedElements = [TypeElementOption(element) for element in cobieTypeElements]
    elementCategories = set(element.category for element in wrappedElements)
    categorizedElements = {
        elementCategory: sorted(
            [
                element
                for element in wrappedElements
                if element.category == elementCategory
            ],
            key=lambda x: x.name,
        )
        for elementCategory in elementCategories
    }

    selection = forms.SelectFromList.show(
        context=categorizedElements,
        title="Select elements to include in COBie & CDX output",
        multiselect=True,
    )

    for element in groupedElements:
        print("{} {}".format(output.linkify(element.Id), query.get_name(element)))

    print("Complete")
