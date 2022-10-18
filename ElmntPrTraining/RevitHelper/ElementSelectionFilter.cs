using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace ElmntPrTraining.RevitHelper
{
    public class ElementSelectionFilter : ISelectionFilter
    {
        bool ISelectionFilter.AllowElement(Element elem)
        {
            // Selected Walls Only
            return (BuiltInCategory)GetCategoryIdAsInteger(elem) == BuiltInCategory.OST_Walls;
        }

        bool ISelectionFilter.AllowReference(Reference reference, XYZ position)
        {
            return false;
        }

        public int GetCategoryIdAsInteger(Element element)
        {
            return element?.Category?.Id?.IntegerValue ?? -1;       
        }
    }
}
