using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Zygotine.WebExpo;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExpostatsExcel2013AddIn
{
    [System.Runtime.InteropServices.ComVisible(true)]
    [System.Runtime.InteropServices.ClassInterface(
    System.Runtime.InteropServices.ClassInterfaceType.None)]
    public partial class Feuil9 : IFeuil9
    {
        const string OEL_CELL_POS = "C7";
        const string MEAS_LIST_START_CELL_POS = "G3";
        const string MCMC_CHAINS_RESULT_CELL_POS = "E9";

        private void Feuil9_Startup(object sender, System.EventArgs e)
        {
        }

        private void Feuil9_Shutdown(object sender, System.EventArgs e)
        {
        }

        public String ReadObservations(ref String obs, String sep)
        {
            String err;
            String[] obsArr = ReadObservations(out err);
            obs = String.Join(sep, obsArr);
            return err;
        }

        public String[] ReadObservations(out String errMsg)
        {
            String[] obsArr = null;
            errMsg = "Invalid observation value(s) or too few observations";
            Range obsCells;
            int elemCount = 0;
            Object[,] obsData = null;

            bool nullCellFound = false;
            int obsCount;
            char measCol = MEAS_LIST_START_CELL_POS[0];
            char measStartRowIdx = MEAS_LIST_START_CELL_POS[1];

            for (obsCount = 0; !nullCellFound; obsCount++)
            {
                Range cell = this.Cells[obsCount + 3, ColIndex(measCol)];
                if (cell.Value == null)
                {
                    nullCellFound = true;
                    obsCount--;
                }
            }
            int measEndRowIdx = measStartRowIdx - '0' + obsCount - 1;
            try
            {
                obsArr = new string[obsCount];
                var rng = String.Format("{0}:{1}{2}", MEAS_LIST_START_CELL_POS, measCol, measEndRowIdx);
                if (obsCount > 1)
                {
                    obsCells = this.Range[rng];
                    obsData = obsCells.Value2;
                }
                else if ( obsCount == 1 )
                {
                    obsData = new Object[1, 1];
                    double singleCellVal = this.Cells[measStartRowIdx - '0', ColIndex(measCol)].Value;
                    obsData[0,0] = singleCellVal;
                }
                
                foreach (object elem in obsData)
                {
                    String obsStr = elem.ToString();
                    obsArr[elemCount] = obsStr;
                    elemCount++;
                }

                // Use dummy val just to validate observations
                double dummyOel = 99.9;
                MeasureList ml = new MeasureList(obsArr, dummyOel);
                if (ml.Error.lst.Count == 0)
                {
                    errMsg = "";
                }
            }
            catch (Exception)
            {
                
            }

            return obsArr;
        }

        private int ColIndex(char col)
        {
            return col - 'A' + 1;
        }
        public string Ohhai()
        {
            return "OH HAI";
        }

        protected override object GetAutomationObject()
        {
            return this;
        }
        #region Code généré par le Concepteur VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Feuil9_Startup);
            this.Shutdown += new System.EventHandler(Feuil9_Shutdown);
        }

        #endregion

    }
}
