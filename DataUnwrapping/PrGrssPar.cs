using Autodesk.Revit.DB;
using System;
using System.Windows.Forms;

namespace DataUnwrapping
{
    public partial class PrGrssPar : System.Windows.Forms.Form
    {
        Document Doc { get; }
        public PrGrssPar()
        {
            InitializeComponent();
        }

     

        public void Time_Tick(object sender, EventArgs e)
        {
            PrgBar.Minimum = 0;
            DWColFrm fr = new DWColFrm(Doc);
            var stWatch = fr.sw;
            int timeElapsed = (int)Math.Round((stWatch.Elapsed.TotalSeconds), 2);
            PrgBar.Maximum = timeElapsed;

            if (PrgBar.Value < PrgBar.Maximum)
            {
                PrgBar.Value = PrgBar.Value + 5;
              //  Time.Enabled = true;
            }

        }

        private void PrGrssPar_Load(object sender, EventArgs e)
        {
            Time.Enabled = true;
        }
    }
}
