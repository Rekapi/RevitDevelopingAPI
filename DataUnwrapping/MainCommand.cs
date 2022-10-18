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
using System.Windows.Media.Imaging;
#endregion

namespace DataUnwrapping
{
    // UIExternal Application class
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MainCommand : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication app)
        {
            #region UI Clarifications
            string t_name = "Data Wrapper";
            string p_name = "UnWapping Elements Parameters";

            // crearing the tab 
            app.CreateRibbonTab(t_name);

            // create the panel 
            var panel = app.CreateRibbonPanel(t_name, p_name);           

            // creating the Columns Button 
            var BtColumn = new PushButtonData("BtColumn", " Columns ", Assembly.GetExecutingAssembly().Location, "DataUnwrapping.ColsCommand");
            BtColumn.ToolTip = "Uwrapping Structural Columns Parameters";
            var btColumn = panel.AddItem(BtColumn) as PushButton;
            // Set the large image shown on button 
            BitmapImage largeImgCol = new BitmapImage(new Uri(@"F:\01-RevitAPiDeveloping\DataUnwrapping\Images\Column.png"));
            btColumn.LargeImage = largeImgCol;

            panel.AddSeparator();

            // creating the Beams Button 
            var BtBeams = new PushButtonData("BtBeams", " Framing ", Assembly.GetExecutingAssembly().Location, "DataUnwrapping.ColsCommand"); //BeamsCommand
            BtBeams.ToolTip = "Uwrapping Structural Framing Parameters";
            var btBeams = panel.AddItem(BtBeams) as PushButton;
            // Set the large image shown on button 
            BitmapImage largeImgBeams = new BitmapImage(new Uri(@"F:\01-RevitAPiDeveloping\DataUnwrapping\Images\FraminIcon.png"));
            btBeams.LargeImage = largeImgBeams;

            panel.AddSeparator();

            // creating the Walls Button 
            var BtWalls = new PushButtonData("BtWalls", " Str Walls ", Assembly.GetExecutingAssembly().Location, "DataUnwrapping.ColsCommand"); //BeamsCommand
            BtWalls.ToolTip = "Uwrapping Structural Walls Parameters";
            var btWalls = panel.AddItem(BtWalls) as PushButton;
            // Set the large image shown on button 
            BitmapImage largeImgWalls = new BitmapImage(new Uri(@"F:\01-RevitAPiDeveloping\DataUnwrapping\Images\walls-icon.png"));
            btWalls.LargeImage = largeImgWalls;


            #endregion

            return Result.Succeeded;
        }
    }
}
