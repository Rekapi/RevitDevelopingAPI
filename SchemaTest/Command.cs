using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Xml;

namespace SchemaTest
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication UiApp = commandData.Application;
            Document doc = UiApp.ActiveUIDocument.Document;

            // retrieve all sheets 
            FilteredElementCollector Sheet_Filtr = new FilteredElementCollector(doc);
            Sheet_Filtr.OfCategory(BuiltInCategory.OST_Sheets);
            Sheet_Filtr.OfClass(typeof(ViewSheet));

            // Create collection of all relevant data 
            List<SheetData> data = new List<SheetData>();

            // iterate through the data
            foreach (ViewSheet v in Sheet_Filtr)
            {
                // create some data for each sheet and add to some serializable collection called data 
                SheetData item = new SheetData(v);
                data.Add(item);
            }

            // write out data collection to xml 
            XmlTextWriter xw = new XmlTextWriter(@"F:/SheetData.xml", null);
            xw.Formatting = Formatting.Indented;
            xw.WriteStartDocument();
            //xw.WriteComment(string.Format(" SheetData from {0} on {1} by Jeremy ",doc.PathName, DateTime.Now));

            
            xw.WriteStartElement("ViewSheets");

            foreach (SheetData item in data)
            {
                xw.WriteStartElement("ViewSheet");
                xw.WriteElementString("IsPlaceholder", item.IsPlaceHolder.ToString());
                xw.WriteElementString("Name", item.Name);
                xw.WriteElementString("SheetNumber", item.SheetNumber);
                xw.WriteElementString("SheetScale", item.SheetScale.ToString());
               // xw.WriteElementString("Sheet Title", item.Title);

                xw.WriteEndElement();
            }
            xw.WriteEndElement();
            xw.WriteEndDocument();
            xw.Close();
            return Result.Succeeded;
        }
    }
    class SheetData
    {
        public bool IsPlaceHolder { get; set; }
        public string Name { get; set; }
        public string SheetNumber { get; set; }

        // adding Sheet Scale 
        public int SheetScale { get; set; }

        public SheetData(ViewSheet viewSheet)
        {
            IsPlaceHolder = viewSheet.IsPlaceholder;
            Name = viewSheet.Name;
            SheetNumber = viewSheet.SheetNumber;
            SheetScale = viewSheet.Scale;            
        }
    }
}
