using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace Appocalypto
{
    public class Mob
    {
        public int OriginalPID { get; set; }
        public void Run(int Minutes = 1)
        {
            OriginalPID = getProcessID();

            System.Timers.Timer t = new System.Timers.Timer(60000 * Minutes); // 1 sec = 1000, 60 sec = 60000
            t.AutoReset = true;
            t.Elapsed += new System.Timers.ElapsedEventHandler(t_Elapsed);
            t.Start();


        }

        private void t_Elapsed(object sender, ElapsedEventArgs e)
        {

            var res = getProcessID();

            if (res != OriginalPID)
            {
                System.Windows.Forms.Application.Exit();
            }
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


            var appId = gui.GetApplication().AppId;

            var processes = Process.GetProcessesByName(@"SAP Business One").Where(x => x.SessionId == Process.GetCurrentProcess().SessionId);


            foreach (var process in processes)
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
