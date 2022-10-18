#region NameSpaces
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using System.Reflection;

#endregion
namespace DataUnwrapping
{
    /// <summary>
    /// 1. after pressing the button the form appears with parameters on check box 
    /// 2. selecting parameters and then press exporting the selected parametrs only
    /// 3. the same form for all elements parameters
    /// 4. bounus : dialog for LookUp Parameter 
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class BeamsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData cmd, ref string message, ElementSet elements)
        {
            UIApplication UiApp = cmd.Application;
            Document doc = UiApp.ActiveUIDocument.Document;
            DWColFrm frm = new DWColFrm(doc);
            frm.ShowDialog();

            return Result.Succeeded;

        }
    }
}
