/**
 * 
 *  Solidworks API wrapper in C#
 * 
 * Author: John Glatts
 *  Date: 8/29/23
 */
using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using SolidWorks.Interop.sldworks;

/*
 * Helper class to get an instance of SolidWorks running
 */
class OpenSldWrksApp {

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

    public static ISldWorks getApp()
    {
        const string SW_PATH = @"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe";

        try
        {
            return StartSwApp(SW_PATH);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to connect to SOLIDWORKS instance: " + ex.Message);
            return null;
        }

    }

    private static ISldWorks StartSwApp(string appPath, int timeoutSec = 50)
    {
        var timeout = TimeSpan.FromSeconds(timeoutSec);
        var startTime = DateTime.Now;

        var prc = Process.Start(appPath);   // start the SW process
        ISldWorks app = null;
        while (app == null)
        {
            if (DateTime.Now - startTime > timeout)
            {
                throw new TimeoutException();
            }
            app = GetSwAppFromProcess(prc.Id);
        }

        return app;
    }

    private static ISldWorks GetSwAppFromProcess(int processId)
    {
        var monikerName = "SolidWorks_PID_" + processId.ToString();
        IBindCtx context = null;
        IRunningObjectTable rot = null;
        IEnumMoniker monikers = null;

        try
        {
            CreateBindCtx(0, out context);

            context.GetRunningObjectTable(out rot);
            rot.EnumRunning(out monikers);

            var moniker = new IMoniker[1];

            while (monikers.Next(1, moniker, IntPtr.Zero) == 0)
            {
                var curMoniker = moniker.First();

                string name = null;

                if (curMoniker != null)
                {
                    try
                    {
                        curMoniker.GetDisplayName(context, null, out name);
                    }
                    catch (UnauthorizedAccessException) { }
                }

                if (string.Equals(monikerName,
                    name, StringComparison.CurrentCultureIgnoreCase))
                {
                    object app;
                    rot.GetObject(curMoniker, out app);
                    return app as ISldWorks;
                }
            }
        }
        finally
        {
            if (monikers != null)
            {
                Marshal.ReleaseComObject(monikers);
            }

            if (rot != null)
            {
                Marshal.ReleaseComObject(rot);
            }

            if (context != null)
            {
                Marshal.ReleaseComObject(context);
            }
        }

        return null;
    }

}
