using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Zygotine.WebExpo;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExpostatsExcel2013AddIn
{
    [System.Runtime.InteropServices.ComVisible(true)]
    [System.Runtime.InteropServices.ClassInterface(
    System.Runtime.InteropServices.ClassInterfaceType.None)]
    public partial class Feuil4 : IFeuil4
    {
        const int MCMC_CHAINS_RESULT_COL_WIDTH = 3;
        const string MCMC_CHAINS_RESULT_CELL_POS = "E9";
        string[] chainNames = new string[] { "mu", "sd" };
        const string OEL_CELL_POS = "C7";
        const string MEAS_LIST_START_CELL_POS = "G3";
        const int N_ITER = 25000;
        string DELAY_WARNING_MSG = "Performing this operation may temporarily freeze your Excel window (for up to {0} seconds). Continue?";
        private void Feuil4_Startup(object sender, System.EventArgs e)
        {
        }

        private void Feuil4_Shutdown(object sender, System.EventArgs e)
        {
        }

        public string GetDelayWarningMsg(int delaySecs)
        {
            return String.Format(DELAY_WARNING_MSG, delaySecs);
        }

        public string GetWorkCompletedMsg()
        {
            return Utils.WORK_COMPLETED_MSG;
        }
        public void CalcMCMCChains(String obs, String sep, double oel, bool confirmDelay = false)
        {
            int estimatedDelaySecs = 10;
            if (!confirmDelay || Utils.ShowPopUpMsg(GetDelayWarningMsg(estimatedDelaySecs), MessageBoxButton.YesNo, MessageBoxImage.Exclamation) == MessageBoxResult.Yes)
            {
                Dictionary<string, double[]> mcmcChains = GetMcmcChains(obs, sep, oel, N_ITER);

                char startCol = MCMC_CHAINS_RESULT_CELL_POS[0];
                int startRow = int.Parse(MCMC_CHAINS_RESULT_CELL_POS.Substring(1).ToString());

                for (int j = 0; j <= chainNames.Length; j++)
                {
                    char col = (char)(startCol + j);
                    double[] vals = j > 0
                        ? mcmcChains[chainNames[j - 1]].ToArray()
                        : Enumerable.Range(1, N_ITER).Select(i => Convert.ToDouble(i)).ToArray();
                    Utils.WriteRange(this.Range, vals, col, startRow);
                }
            }
        }

        public void EraseMCMCChains()
        {
            /* Clear previous results */
            int lastUsedRow = this.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            String RangeToClear = String.Format("{0}:{1}{2}", MCMC_CHAINS_RESULT_CELL_POS, (char) (((int) MCMC_CHAINS_RESULT_CELL_POS[0])+MCMC_CHAINS_RESULT_COL_WIDTH), lastUsedRow);
            this.Range[RangeToClear].Clear(); ;
        }

        private Dictionary<string, double[]> GetMcmcChains(String obsstr, String sep, double oel, int niter)
        {
            Dictionary<string, double[]> ret = null;

            String errMsg = "";
            ret = new Dictionary<string, double[]>();
            McmcParameters mcmc = new McmcParameters();
            SEGInformedVarModelParameters specParams;
            SEGInformedVarModel model = null;
            MeasureList ml;

            String[] obs = obsstr.Split(sep[0]);
            ml = new MeasureList(obs, oel);

            mcmc.NIter = niter;

            specParams = SEGInformedVarModelParameters.GetDefaults(true);

            model = new SEGInformedVarModel(ml, specParams, mcmc);
            model.Compute();

            ModelResult mr = model.Result;
            string[] prms = new string[] { "mu", "sd" };
            foreach (string prm in prms)
            {
                ret.Add(prm, mr.GetChainByName(String.Format("{0}Sample", prm)));
            }
            if (errMsg.Length > 0)
            {
                Utils.ShowPopUpMsg(String.Format("Operation failed: {0}", errMsg), MessageBoxButton.OK, MessageBoxImage.Error);
                ret = null;
            }
            return ret;
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
            this.Startup += new System.EventHandler(Feuil4_Startup);
            this.Shutdown += new System.EventHandler(Feuil4_Shutdown);
        }

        #endregion

    }
}
