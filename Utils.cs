using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExpostatsExcel2013AddIn
{
    [System.Runtime.InteropServices.ClassInterface(
    System.Runtime.InteropServices.ClassInterfaceType.None)]
    class Utils
    {
        public static String WORK_COMPLETED_MSG = "Operation completed.";
        static public MessageBoxResult ShowPopUpMsg(string msg = null, MessageBoxButton btn = MessageBoxButton.OK, MessageBoxImage icon = MessageBoxImage.Information, string popUpTitle = "Expostats")
        {
            msg = msg ?? WORK_COMPLETED_MSG;
            System.Windows.Window w = new System.Windows.Window();
            return MessageBox.Show(w, msg, popUpTitle, btn, icon);
        }

        static public int ColIndex(char col)
        {
            return col - 'A' + 1;
        }

        static public void WriteRange(Worksheet_RangeType range, double[] contents, char startCol, int startRow, bool writeColumn = true)
        {
            object[] objs = new object[contents.Length];
            contents.CopyTo(objs, 0);
            WriteRange(range, objs, startCol, startRow, writeColumn);
        }

        static public void WriteRange(Worksheet_RangeType range, object[] contents, char startCol, int startRow, bool writeColumn = true)
        {
            int leng = contents.Length;
            object[,] arr = writeColumn ? new object[leng, 1] : new object[1, leng];
            for (int r = 0; r < leng; r++)
            {
                if (writeColumn)
                    arr[r, 0] = contents[r];
                else
                    arr[0, r] = contents[r];
            }

            WriteRange(range, arr, startCol, startRow);
        }

        static public void WriteRange(Worksheet_RangeType range, object[,] contents, char startCol, int startRow)
        {
            String rng = String.Format("{0}{1}:{2}{3}",
                startCol,
                startRow,
                (char)(startCol + contents.GetLength(1) - 1),
                startRow + contents.GetLength(0) - 1);
            Range chainRange = range[rng];
            chainRange.Value = contents;
        }
    }
}
