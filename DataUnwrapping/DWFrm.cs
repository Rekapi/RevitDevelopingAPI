using System;
using System.Collections.Generic;
using System.Windows.Forms;

#region NameSpaces
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
#endregion

namespace DataUnwrapping
{
    public partial class DWColFrm : System.Windows.Forms.Form
    {
        public const double converter = 3.28083989501313;
        public Stopwatch sw;
        Document Doc { get; }
        public DWColFrm(Document doc)
        {
            Doc = doc;
            InitializeComponent();
            // Columns 
            List<string> ColParams = new List<string>()
            {
                "Element ID","Column Location Mark","Column Style","Base Level","Base Offset",
                "Top Level","Top Offset","Length","Width","Height","Volume"

            };
            foreach (string param in ColParams)
            {
                CParList.Items.Add(param);
            }    
        }

        // Method for using Escape key to close the form 
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Escape))
            {
                Close();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void BtnOpen_Click(object sender, EventArgs e)
        {
            #region OpenFileDialog
            System.Windows.Forms.OpenFileDialog dialogRead = new System.Windows.Forms.OpenFileDialog();
            dialogRead.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            dialogRead.DefaultExt = ".xlsx";
            dialogRead.Filter = "xlsx files (*.xlsx) |*.xlsx";
            dialogRead.ShowDialog();
            string ExcelName = dialogRead.FileName;
            Close();

            //PrGrssPar pr = new PrGrssPar();
            
            //pr.Show();
            sw = Stopwatch.StartNew();
            #endregion
            try
            {                
                using (Transaction t = new Transaction(Doc, "Exporting Param"))
                {
                    t.Start();

                    // Excel References
                    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook workSheet = excel.Workbooks.Open(ExcelName);
                    Excel.Worksheet sheet = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

                    FilteredElementCollector filter = new FilteredElementCollector(Doc);
                    ElementClassFilter FamilyInstFilter = new ElementClassFilter(typeof(FamilyInstance));
                    ElementCategoryFilter ColCategFilter = new ElementCategoryFilter(BuiltInCategory.OST_StructuralColumns);
                    LogicalAndFilter columns = new LogicalAndFilter(FamilyInstFilter, ColCategFilter);
                    IList<Element> Cols = filter.WherePasses(columns).ToElements();

                    #region ParameterInExcel
                    // Element ID
                    //sheet.Range[sheet.Cells[1, 2], sheet.Cells[2, 2]].Merge();
                    excel.Cells[1, 1].Interior.Color = Excel.XlRgbColor.rgbAliceBlue;
                    excel.Cells[1, 1].Font.Name = "Century Gothic";
                    excel.Cells[1, 1].Font.Bold = true;
                    excel.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excel.Cells[1, 1] = "Element ID";
                    excel.Cells[1, 1].EntireColumn.AutoFit();
                    sheet.Range["1:1"].RowHeight = 15;

                    // Column Location Mark
                    excel.Cells[1, 2].Interior.Color = Excel.XlRgbColor.rgbAliceBlue;
                    excel.Cells[1, 2].Font.Name = "Century Gothic";
                    excel.Cells[1, 2].Font.Bold = true;
                    excel.Cells[1, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excel.Cells[1, 2] = "Axis";
                    excel.Cells[1, 2].EntireColumn.AutoFit();
                    sheet.Range["1:2"].RowHeight = 15;

                    // Column Style
                    excel.Cells[1, 3].Interior.Color = Excel.XlRgbColor.rgbAliceBlue;
                    excel.Cells[1, 3].Font.Name = "Century Gothic";
                    excel.Cells[1, 3].Font.Bold = true;
                    excel.Cells[1, 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excel.Cells[1, 3] = "Column Style";
                    excel.Cells[1, 3].EntireColumn.AutoFit();
                    sheet.Range["1:3"].RowHeight = 15;

                    // Base Level
                    excel.Cells[1, 4].Interior.Color = Excel.XlRgbColor.rgbAliceBlue;
                    excel.Cells[1, 4].Font.Name = "Century Gothic";
                    excel.Cells[1, 4].Font.Bold = true;
                    excel.Cells[1, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excel.Cells[1, 4] = "Base Level";
                    excel.Cells[1, 4].EntireColumn.AutoFit();
                    sheet.Range["1:4"].RowHeight = 15;

                    // Base Offset
                    excel.Cells[1, 5].Interior.Color = Excel.XlRgbColor.rgbAliceBlue;
                    excel.Cells[1, 5].Font.Name = "Century Gothic";
                    excel.Cells[1, 5].Font.Bold = true;
                    excel.Cells[1, 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excel.Cells[1, 5] = "Base Offset";
                    excel.Cells[1, 5].EntireColumn.AutoFit();
                    sheet.Range["1:5"].RowHeight = 15;

                    // Top Level
                    excel.Cells[1, 6].Interior.Color = Excel.XlRgbColor.rgbAliceBlue;
                    excel.Cells[1, 6].Font.Name = "Century Gothic";
                    excel.Cells[1, 6].Font.Bold = true;
                    excel.Cells[1, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excel.Cells[1, 6] = "Top Level";
                    excel.Cells[1, 6].EntireColumn.AutoFit();
                    sheet.Range["1:6"].RowHeight = 15;

                    // Top Offset
                    excel.Cells[1, 7].Interior.Color = Excel.XlRgbColor.rgbAliceBlue;
                    excel.Cells[1, 7].Font.Name = "Century Gothic";
                    excel.Cells[1, 7].Font.Bold = true;
                    excel.Cells[1, 7].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excel.Cells[1, 7] = "Top Offset";
                    excel.Cells[1, 7].EntireColumn.AutoFit();
                    sheet.Range["1:7"].RowHeight = 15;

                    // Length
                    excel.Cells[1, 8].Interior.Color = Excel.XlRgbColor.rgbAliceBlue;
                    excel.Cells[1, 8].Font.Name = "Century Gothic";
                    excel.Cells[1, 8].Font.Bold = true;
                    excel.Cells[1, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excel.Cells[1, 8] = "Length";
                    excel.Cells[1, 8].EntireColumn.AutoFit();
                    sheet.Range["1:8"].RowHeight = 15;

                    // Width
                    excel.Cells[1, 9].Interior.Color = Excel.XlRgbColor.rgbAliceBlue;
                    excel.Cells[1, 9].Font.Name = "Century Gothic";
                    excel.Cells[1, 9].Font.Bold = true;
                    excel.Cells[1, 9].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excel.Cells[1, 9] = "Width";
                    excel.Cells[1, 9].EntireColumn.AutoFit();
                    sheet.Range["1:9"].RowHeight = 15;

                    // Height
                    excel.Cells[1, 10].Interior.Color = Excel.XlRgbColor.rgbAliceBlue;
                    excel.Cells[1, 10].Font.Name = "Century Gothic";
                    excel.Cells[1, 10].Font.Bold = true;
                    excel.Cells[1, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excel.Cells[1, 10] = "Height";
                    excel.Cells[1, 10].EntireColumn.AutoFit();
                    sheet.Range["1:10"].RowHeight = 15;

                    // Volume
                    excel.Cells[1, 11].Interior.Color = Excel.XlRgbColor.rgbAliceBlue;
                    excel.Cells[1, 11].Font.Name = "Century Gothic";
                    excel.Cells[1, 11].Font.Bold = true;
                    excel.Cells[1, 11].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    excel.Cells[1, 11] = "Volume";
                    excel.Cells[1, 11].EntireColumn.AutoFit();
                    sheet.Range["1:11"].RowHeight = 15;

                    #endregion

                    #region Element Parameters
                    int XlRow = 2;
                    foreach (Element elem in Cols)
                    {
                        ElementType elType = Doc.GetElement(elem.GetTypeId()) as ElementType;
                        foreach (int index in CParList.CheckedIndices)
                        {
                            if (index == 0)
                            {
                                // Parameters - Element ID
                                ElementId elemID = elem.GetTypeId() as ElementId;
                                int eID = elemID.IntegerValue;
                                // Excel Cells 
                                excel.Cells[XlRow, 1] = eID;
                            }
                            else if (index == 1)
                            {
                                // Parameters - Columns Mark Locations
                                string axe = elem.get_Parameter(BuiltInParameter.COLUMN_LOCATION_MARK).AsString();
                                excel.Cells[XlRow, 2] = axe;
                            }
                            else if (index == 2)
                            {
                                // Parameters - Column Style
                                string style = elem.get_Parameter(BuiltInParameter.SLANTED_COLUMN_TYPE_PARAM).AsValueString();
                                excel.Cells[XlRow, 3] = style;
                            }
                            else if (index == 3)
                            {
                                // Parameters - Base Level
                                string baseLevel = elem.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsValueString();
                                excel.Cells[XlRow, 4] = baseLevel;
                            }
                            else if (index == 4)
                            {
                                // Parameters - Base Offset
                                string baseOffset = elem.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).AsDouble().ToString();
                                excel.Cells[XlRow, 5] = baseOffset;
                            }
                            else if (index == 5)
                            {
                                // Parameters - Top Level
                                string topLevel = elem.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsValueString();
                                excel.Cells[XlRow, 6] = topLevel;
                            }
                            else if (index == 6)
                            {
                                // Parameters - Top Offset
                                string topOffset = elem.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).AsDouble().ToString();
                                excel.Cells[XlRow, 7] = topOffset;
                            }
                            else if (index == 7)
                            {
                                // Parameters - Length
                                double h = elType.LookupParameter("h").AsDouble();
                                double ColLength = Math.Round(h / converter, 2);
                                excel.Cells[XlRow, 8] = ColLength.ToString();
                            }
                            else if (index == 8)
                            {
                                // Parameters - Width
                                double w = elType.LookupParameter("b").AsDouble();
                                double ColWidth = Math.Round(w / converter, 2);
                                excel.Cells[XlRow, 9] = ColWidth.ToString();
                            }
                            else if (index == 9)
                            {
                                // Parameters - Height
                                double l = elem.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM).AsDouble();
                                double ColHeight = Math.Round(l / converter, 2);
                                excel.Cells[XlRow, 10] = ColHeight.ToString();
                            }
                            else if (index == 10)
                            {
                                // Parameters - Volume 
                                excel.Cells[XlRow, 11] = elem.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsValueString();
                            }
                            else
                            {
                                TaskDialog.Show("Fatal Error", "Please Choose Excel File or Select item");
                            }
                        }
                        ++XlRow;
                        sheet.Columns.AutoFit();
                        Excel.Range AdRange = sheet.UsedRange;
                        AdRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        AdRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    }
                   
                    #endregion
                    workSheet.Close(true, Type.Missing, Type.Missing);
                    excel.Visible = true;
                    sw.Stop();
                //    string timeElapsed = Math.Round((sw.Elapsed.TotalSeconds), 2).ToString();
                    //excel.Quit();
                    excel.Workbooks.Open(ExcelName);
                    t.Commit();
                    }   
            }
            catch (Exception)
            {
                TaskDialog.Show("Error","Please select a file");
                Close();
            }
            
            
            ///TODO :
            /// 
            /// 
            /// 
            /// 4. progress bar on exporting process
         
        }

        private void SelectAllBtn_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < CParList.Items.Count; i++)
            {
                CParList.SetItemChecked(i, true);
            }
        }

        private void ClearBtn_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < CParList.Items.Count; i++)
            {
                CParList.SetItemChecked(i, false);
            }
        }
    }
}
