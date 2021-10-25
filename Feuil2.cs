using System;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using MessageBox = System.Windows.MessageBox;
using Office = Microsoft.Office.Core;

namespace ExpostatsExcel2013AddIn
{
    [System.Runtime.InteropServices.ComVisible(true)]
    [System.Runtime.InteropServices.ClassInterface(
    System.Runtime.InteropServices.ClassInterfaceType.None)]
    public partial class Feuil2 : IFeuil2
    {
        const char chainsStartCol = 'E';
        string[] chainNames = new string[] { "mu", "sd" };

        const int MCMC_CHAINS_RESULT_COL_WIDTH = 3;
        const int N_ITER = 25000;
        const int START_ROW_IDX = 10;
        readonly (char, int, char) ROS_OBS_OUTPUT_CELL_ADDR = ('B', START_ROW_IDX, 'H');
        readonly (char, int) ROS_REG_OUTPUT_CELL_ADDR = ('J', START_ROW_IDX);
        readonly (char, int, char) ROS_DET_OUTPUT_CELL_ADDR = ('M', START_ROW_IDX, 'P');

        private void Feuil2_Startup(object sender, System.EventArgs e)
        {
            //Application.Run("Compiler");
        }

        private void Feuil2_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void EraseROSResults()
        {
            /* Clear previous results */
            int lastUsedRow = this.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            String range1ToClear = String.Format("{0}{1}:{2}{3}", ROS_OBS_OUTPUT_CELL_ADDR.Item1, START_ROW_IDX, ROS_OBS_OUTPUT_CELL_ADDR.Item3, lastUsedRow);
            String range2ToClear = String.Format("{0}{1}:{2}{3}", ROS_DET_OUTPUT_CELL_ADDR.Item1, START_ROW_IDX, ROS_DET_OUTPUT_CELL_ADDR.Item3, lastUsedRow);

            this.Range[range1ToClear].Clear();
            this.Range[range2ToClear].Clear();
        }

        public String CalcROSResults(String obs_concat, String sep)
        {
            NDExpo nde = new NDExpo();
            nde.reset();
            String err = "";
            CultureInfo culture1 = CultureInfo.CurrentCulture;
            String[] obs = obs_concat.Split(new char[] { sep[0] });

            Func<Group, Double> ParseFloat = (Group g) => Double.Parse(Regex.Replace(g.ToString(), @"[,.]", culture1.NumberFormat.NumberDecimalSeparator));
            void CopyRow(object[,] contents, object[] row, int rowI)
            {
                for (int colJ = 0; colJ < row.Length; colJ++)
                {
                    contents[rowI, colJ] = row[colJ];
                }
            }

            if (err.Length == 0)
            {
                try
                {
                    int pos = 0;
                    foreach (String o in obs)
                    {
                        var patternDetected = @"^\s*([0-9]+([,.][0-9]+)?)\s*$";
                        var patternLeftOrRightCensored = @"^\s*([<>])\s*([0-9]+([,.][0-9]+)?)\s*$";
                        var patternIntCensored = @"^\s*\[\s*([0-9]+([,.][0-9]+)?)\s*\-\s*([0-9]+([,.][0-9]+)?)\s*\]\s*$";
                        var match0 = Regex.Match(o, patternDetected);
                        var match1 = Regex.Match(o, patternLeftOrRightCensored);
                        var match2 = Regex.Match(o, patternIntCensored);
                        int isNd = 0;
                        double val = -1, val2 = -1;
                        if (match0.Success)
                        {
                            val = ParseFloat(match0.Groups[1]);
                        }
                        else
                        {
                            isNd = 1;
                            if (match1.Success)
                            {
                                val = ParseFloat(match1.Groups[2]);
                                if (match1.Groups[1].ToString() == ">")
                                {
                                    val *= 9 / 4;
                                }
                            }
                            else
                            if (match2.Success)
                            {
                                val = ParseFloat(match2.Groups[1]);
                                val2 = ParseFloat(match2.Groups[3]);
                                val = (val + val2) / 2;
                            }
                            else
                            {
                                throw new Exception(String.Format("Invalid observation: {0}", o));
                            }
                        }
                        nde.addDatum(isNd, val, pos++);
                    }

                    nde.doCalc();
                    if (nde.error == 0) {
                        NDExpo.GraphData gdata = new NDExpo.GraphData(nde);
                        gdata.getDataForChart();

                        EraseROSResults();

                        int rowWidth = (int)ROS_OBS_OUTPUT_CELL_ADDR.Item3 - ROS_OBS_OUTPUT_CELL_ADDR.Item1 + 1;
                        object[,] contents = new object[nde.dataSet.Count, rowWidth];
                        int rowI = 0;
                        bool writeCol = false;
                        int totalND = 0;
                        foreach (NDExpo.Datum d in nde.dataSet)
                        {
                            object[] rw = new object[]
                            {
                                                    d.position+1,
                                                    d.isND == 1 ? d.detectionLimitValue : d.value,
                                                    d.isND,
                                                    d.plottingPosition,
                                                    d.score,
                                                    d.finalValue,
                                                    Math.Log(d.finalValue)
                            };
                            CopyRow(contents, rw, rowI++);
                            totalND += d.isND;
                        }
                        Utils.WriteRange(this.Range, contents, ROS_OBS_OUTPUT_CELL_ADDR.Item1, ROS_OBS_OUTPUT_CELL_ADDR.Item2);
                        Utils.WriteRange(this.Range, new object[] { nde.global.slope, nde.global.intercept }, ROS_REG_OUTPUT_CELL_ADDR.Item1, ROS_REG_OUTPUT_CELL_ADDR.Item2, writeCol);

                        int[] nRowsMini = { nde.dataSet.Count - totalND, totalND };
                        for (int nd = 0; nd <= 1; nd++)
                        {
                            object[,] miniTable = new object[nRowsMini[nd], 2];
                            int i = 0;

                            foreach (NDExpo.Datum d in nde.dataSet)
                            {
                                if (d.isND == nd)
                                {
                                    CopyRow(miniTable, new object[] { d.score, Math.Log(d.finalValue) }, i++);
                                }
                            }
                            char col = (char)(ROS_DET_OUTPUT_CELL_ADDR.Item1 + 2 * nd);
                            Utils.WriteRange(this.Range, miniTable, col, START_ROW_IDX);
                        }

                        this.Columns.AutoFit();
                    } else
                    {
                        switch (nde.error)
                        {
                            case NDExpo.ERR_TOOMANY_ND:
                                err = "The proportion of non-detects cannot be greater than 80%";
                                break;
                            case NDExpo.ERR_GRTSTDL:
                                err = "The highest limit of detection cannot be higher than the highest detected value";
                                break;
                            case NDExpo.ERR_NENGH_DATA:
                                err = "The procedure requires at least 5 observations";
                                break;
                            case NDExpo.ERR_NENGH_DET:
                                err = "The procedure requires at least 3 detected observations";
                                break;
                            default:
                                err = "Unknown error";
                                break;
                        }                           
                    }
                }
                catch (Exception ex)
                {
                    err = ex.Message;
                }
            }

            return err;
        }

        private Microsoft.Office.Tools.Excel.NamedRange namedRange1;

        public void CreateVstoNamedRange(Excel.Range range, string name)
        {
            if (!this.Controls.Contains(name))
            {
                namedRange1 = this.Controls.AddNamedRange(range, name);
                namedRange1.Selected += new Excel.DocEvents_SelectionChangeEventHandler(
                        namedRange1_Selected);
            }
            else
            {
                MessageBox.Show("A named range with this specific name " +
                    "already exists on the worksheet.");
            }
        }

        private void namedRange1_Selected(Microsoft.Office.Interop.Excel.Range Target)
        {
            MessageBox.Show("This named range was created by Visual Studio " +
                "Tools for Office.");
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
            this.Startup += new System.EventHandler(Feuil2_Startup);
            this.Shutdown += new System.EventHandler(Feuil2_Shutdown);
        }

        #endregion

    }
}
