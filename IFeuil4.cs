namespace ExpostatsExcel2013AddIn
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public interface IFeuil4
    {
        void CalcMCMCChains(string obs, string sep, double oel, bool confirmDelay = false);
        string GetDelayWarningMsg(int delaySecs);
        string GetWorkCompletedMsg();
    }
}