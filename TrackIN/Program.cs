using System;
using System.Collections.Generic;
using System.Deployment.Application;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WashU.BatemanLab.MassSpec.TrackIN
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            char[] argsDelimiter = { '*' };

            if (args.Length > 0)
            {
                string[] cmdArgs = SplitArgs(args[0], argsDelimiter);
                Application.Run(new MainForm(cmdArgs));
            }
            else
            {
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    var activationData = AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData[0];
                    string[] activationArgs = SplitArgs(activationData, argsDelimiter);
                    if (activationArgs.Length > 0)
                    {
                        Application.Run(new MainForm(activationArgs));
                    }
                    else
                    {
                        Application.Run(new MainForm());
                    }
                }
                else
                {
                    Application.Run(new MainForm());
                }
            }
        }

        public static string[] SplitArgs(string combinedArgs, char[] delimiter)
        {
            return combinedArgs.Split(delimiter, StringSplitOptions.None);
        }
    }
}
