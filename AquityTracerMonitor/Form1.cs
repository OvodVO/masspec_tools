using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using test1 = WATERSACQUITYSTATUSLib;
using test2 = ACQUISITIONLib;
using discover = DiscoverInstrumentsLib;
using AcquityStatus = Waters.ACQUITY.AppMonitor;// .SystemStatusControl;

using logComm = AcquityLog;

using v1 = AcquityASMServerLib;


//using serv = Waters.ACQUITY.AcquityServer;

namespace AquityTracerMonitor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            v1.IInstrumentProperties dff = new v1.IInstrumentProperties();
            dff.StatusProperties

            AcquityLog.AcquityLogService dff = new logComm.AcquityLogService();

            Waters.ACQUITY.AcquityServer serv = new Waters.ACQUITY.AcquityServer();
            serv.

            logComm.CollectData ttt = new AcquityLog.CollectData();

            logComm.Log _log = new logComm.Log();

          //  logComm.

          //  Waters AcquityPump pump = new _log.AcquityPump();

          //  _log.AcquityPump.

            //AcquityStatus. .SystemStatusControl

            //discover.Discover disk = new discover.Discover();
            //disk.Discover

           // test1..WatersACQUITYStatus ifaceAcStatus = new test1.WatersACQUITYStatus();

           // ifaceAcStatus.


            test2.Acquisition acquisition = new test2.Acquisition();
          //  acquisition.

            OpenFileDialog methodfiledlg = new OpenFileDialog();

            if (methodfiledlg.ShowDialog() == DialogResult.OK)
            {
               
            }

            
        }
    }
}
