using Microsoft.Office.Interop.Excel;
using System;

namespace ExpostatsExcel2013AddIn
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public interface IFeuil2
    {
        String CalcROSResults(String obs_concat, String sep);
        void EraseROSResults();
        void CreateVstoNamedRange(Range range, string name);
    }
}