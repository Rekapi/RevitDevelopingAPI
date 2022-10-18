using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ElmntPrTraining.RevitHelper;
using System.Collections;
using Autodesk.Revit.UI.Selection;
using System.Windows.Markup;
using System.Diagnostics;

using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;



namespace ElmntPrTraining
{
    /// <summary>
    /// Calcualting wall area 
    /// using the following equation wall.width * (BuiltInParameter_TOPOFFSET - BuiltInParameter_BASEOFFSET)
    /// </summary>

    [Transaction(TransactionMode.Manual)]
    public class WallsArea : IExternalCommand
    {
        public const double converter = 3.283582089552239;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            
            //string prompt = "";
            Result result = Result.Succeeded;
            UIApplication UiApp = commandData.Application;
            Document doc = UiApp.ActiveUIDocument.Document;

            Stopwatch sw = Stopwatch.StartNew();
      
            FilteredElementCollector filter = new FilteredElementCollector(doc);
            ElementClassFilter FamilyInstFilter = new ElementClassFilter(typeof(FamilyInstance));
            ElementCategoryFilter ColCategFilter = new ElementCategoryFilter(BuiltInCategory.OST_StructuralColumns);
            LogicalAndFilter columns = new LogicalAndFilter(FamilyInstFilter, ColCategFilter);
            ICollection<Element> Cols = filter.WherePasses(columns).ToElements();

            // Another Way for family symbol
            //      FilteredElementCollector colFilter = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)) 
            //                                                                   .OfCategory(BuiltInCategory.OST_StructuralColumns);

            System.Windows.Forms.OpenFileDialog dialogRead = new System.Windows.Forms.OpenFileDialog();
            dialogRead.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            dialogRead.DefaultExt = ".xlsx";
            dialogRead.Filter = "xlsx files (*.xlsx) |*.xlsx";
            dialogRead.ShowDialog();
            string ExcelName = dialogRead.FileName;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workSheet = excel.Workbooks.Open(ExcelName);
            Microsoft.Office.Interop.Excel.Worksheet sheet = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

            sheet.Columns.AutoFit();
            Excel.Range AdRange = sheet.UsedRange;
            AdRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            AdRange.Borders.Weight = Excel.XlBorderWeight.xlThin; 
            // Item Description Cell - Type Parameter
            sheet.Range[sheet.Cells[1, 2], sheet.Cells[2, 2]].Merge();
            excel.Cells[1, 2].Interior.Color = Excel.XlRgbColor.rgbDarkGray;
            excel.Cells[1, 2].Font.Name = "Century Gothic";
            excel.Cells[1, 2].Font.Bold = true;
            excel.Cells[1, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;       
            excel.Cells[1, 2] = "Item Description";
            sheet.Range["1:2"].RowHeight = 25;
            

            // Column ID
            sheet.Range[sheet.Cells[1, 3], sheet.Cells[2, 3]].Merge();
            excel.Cells[1, 3].Interior.Color = Excel.XlRgbColor.rgbDarkGray;
            excel.Cells[1, 3].Font.Name = "Century Gothic";
            excel.Cells[1, 3].Font.Bold = true;
            excel.Cells[1, 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            excel.Cells[1, 3] = "Columns ID";

            // Place Cell - Column Location Mark
            sheet.Range[sheet.Cells[1, 4], sheet.Cells[2, 4]].Merge();
            excel.Cells[1, 4].Interior.Color = Excel.XlRgbColor.rgbDarkGray;
            excel.Cells[1, 4].Font.Name = "Century Gothic";
            excel.Cells[1, 4].Font.Bold = true;
            excel.Cells[1, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            excel.Cells[1, 4] = "Axis";

            // Level Cell 
            sheet.Range[sheet.Cells[1, 5], sheet.Cells[2, 5]].Merge();
            excel.Cells[1, 5].Interior.Color = Excel.XlRgbColor.rgbDarkGray;
            excel.Cells[1, 5].Font.Name = "Century Gothic";
            excel.Cells[1, 5].Font.Bold = true;
            excel.Cells[1, 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            excel.Cells[1, 5] = "Base Level";

            // Dimension Cell
            sheet.Range[sheet.Cells[1, 6], sheet.Cells[1, 8]].Merge();
            excel.Cells[1, 6].Interior.Color = Excel.XlRgbColor.rgbDarkGray;
            excel.Cells[1, 6].Font.Name = "Century Gothic";
            excel.Cells[1, 6].Font.Bold = true;
            excel.Cells[1, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            excel.Cells[1, 6] = "Dimension";

            #region DimensionCells
            // Width - Length - Height
            excel.Cells[2, 6].Interior.Color = Excel.XlRgbColor.rgbDarkGray;
            excel.Cells[2, 6].Font.Name = "Century Gothic";
            excel.Cells[2, 6].Font.Bold = true;
            excel.Cells[2, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            excel.Cells[2, 6] = "WIDTH";

            excel.Cells[2, 7].Interior.Color = Excel.XlRgbColor.rgbDarkGray;
            excel.Cells[2, 7].Font.Name = "Century Gothic";
            excel.Cells[2, 7].Font.Bold = true;
            excel.Cells[2, 7].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            excel.Cells[2, 7] = "LENGTH";

            excel.Cells[2, 8].Interior.Color = Excel.XlRgbColor.rgbDarkGray;
            excel.Cells[2, 8].Font.Name = "Century Gothic";
            excel.Cells[2, 8].Font.Bold = true;
            excel.Cells[2, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            excel.Cells[2, 8] = "HEIGHT";
            #endregion

            // Volume 
            sheet.Range[sheet.Cells[1, 9], sheet.Cells[2, 9]].Merge();
            excel.Cells[1, 9].Interior.Color = Excel.XlRgbColor.rgbDarkGray;
            excel.Cells[1, 9].Font.Name = "Century Gothic";
            excel.Cells[1, 9].Font.Bold = true;
            excel.Cells[1, 9].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            excel.Cells[1, 9] = "Volume";


            int Xlrow = 2;
            foreach (Element el in Cols)
            {
                
                ElementType elType = doc.GetElement(el.GetTypeId()) as ElementType;
                // column location mark
                //                string ColAxis = el.get_Parameter(BuiltInParameter.COLUMN_LOCATION_MARK).AsString();
                // Level 
                //string ColLevel = el.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsValueString();
                //// Volume
                string ColVolume = el.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsValueString();

                #region Dimensions
                double length = el.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM).AsDouble();
                double Width = elType.LookupParameter("WIDTH").AsDouble();
                double Height = elType.LookupParameter("LENGTH").AsDouble();
                double ColLength = Math.Round(length / converter, 2);
                double ColWidth = Math.Round(Width / converter, 2);
                double ColHeight = Math.Round(Height / converter, 2);
                ElementId elId = el.GetTypeId() as ElementId;
                #endregion

                // Columns ID
                string colID = elType.LookupParameter("COLUMN-ID").AsString();
                //prompt += $"{Environment.NewLine} Type:  {elType.Name}{Environment.NewLine} Family:  {elType.FamilyName}{Environment.NewLine} Level:  {ColLevel}{Environment.NewLine}Width:  {ColWidth}{Environment.NewLine} Height:  {ColHeight}{Environment.NewLine} Length:  {ColLength} {Environment.NewLine} Axis:  {ColAxis} {Environment.NewLine} ColVolume :  {ColVolume} {Environment.NewLine}";

                //excel.Cells[Xlrow+1, 1] = elId.IntegerValue.ToString();
                excel.Cells[Xlrow, 2] = elType.Name; // Item Description 
                excel.Cells[Xlrow, 3] = colID.ToString(); // Columns ID 
                excel.Cells[Xlrow, 4] = el.get_Parameter(BuiltInParameter.COLUMN_LOCATION_MARK).AsString(); // Axis 
                excel.Cells[Xlrow, 5] = el.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsValueString(); // Base Level
                excel.Cells[Xlrow+1, 6] = ColWidth.ToString(); // Width
                excel.Cells[Xlrow+1, 7] = ColHeight.ToString(); // Length
                excel.Cells[Xlrow+1, 8] = ColLength.ToString(); // Height
                excel.Cells[Xlrow + 1, 9] = ColVolume.ToString(); // Volume

                ++Xlrow;
            }
            workSheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
            sw.Stop();
            string timeElapsed = Math.Round((sw.Elapsed.TotalSeconds), 2).ToString();
            TaskDialog.Show("Status", "The Task is Completed and it took about (" + timeElapsed + ") Seconds");
            return result;
        }

        // ******************** Get Parameter Value Method ********************
        public string GetParamterValue(Parameter parameter)
        {
            switch (parameter.StorageType)
            {
                case StorageType.None:
                    return parameter.AsValueString();               
                case StorageType.Integer:
                    return parameter.AsValueString();
                case StorageType.Double:
                    return parameter.AsValueString();
                case StorageType.String:
                    return parameter.AsString();
                case StorageType.ElementId:
                    return parameter.AsElementId().IntegerValue.ToString();
                default:
                    return "";      
            }         
        }
    }
}
