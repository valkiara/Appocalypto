using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Appocalypto
{
    public class Mob
    {
        public void Run(int Minutes = 1)
        {
            var orgPID = getProcessID();

            var startTimeSpan = TimeSpan.Zero;
            var periodTimeSpan = TimeSpan.FromMinutes(Minutes);

            var timer = new System.Threading.Timer((e) =>
            {
                var res = getProcessID();

                if (res != orgPID)
                {
                    System.Windows.Forms.Application.Exit();
                }

            }, null, startTimeSpan, periodTimeSpan);
        }

        private int getProcessID()
        {
            var gui = new SboGuiApi();
            try
            {
                gui.Connect("0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056");
            }
            catch (Exception)
            {
                return -1;
            }

            int sessionID = Process.GetCurrentProcess().SessionId;

            var application = gui.GetApplication();

            var appId = application.AppId;
            var processes = Process.GetProcessesByName(@"SAP Business One");

            foreach (var process in processes.Where(x => x.SessionId == sessionID))
            {
                try
                {
                    var processAppId = gui.GetAppIdFromProcessId(process.Id);
                    if (appId == processAppId)
                    {
                        return process.Id;
                    }
                }
                catch (COMException)
                {
                    // GetAppIdFromProcessId will throw when the current process is not the one we are connected to
                    continue;
                }
            }

            return -1;
        }
    }
}
