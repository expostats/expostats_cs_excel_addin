using System;

namespace ExpostatsExcel2013AddIn
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public interface IFeuil4
    {
        String CalcMCMCChains(string obs, string sep, double oel, bool confirmDelay = false);

        void EraseMCMCChains();
        string GetDelayWarningMsg(int delaySecs);
        string GetWorkCompletedMsg();
    }
}