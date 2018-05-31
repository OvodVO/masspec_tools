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
            MainForm mainForm = null;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            char[] argsDelimiter = { '*' };
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                if (AppDomain.CurrentDomain.SetupInformation.ActivationArguments != null)
                {
                    if (AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData != null
                        && AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData.Length > 0)
                    {
                        var activationData = AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData[0];
                        string[] activationArgs = SplitArgs(activationData, argsDelimiter);

                        mainForm = new MainForm(activationArgs);
                    }
                    else
                    {
                        mainForm = new MainForm();
                    }
                }
            }
            else
            {
                if (args.Length > 0)
                {
                    string[] cmdArgs = SplitArgs(args[0], argsDelimiter);
                    mainForm = new MainForm(cmdArgs);
                }
                else
                {
                    mainForm = new MainForm();
                }
            }
            Application.Run(mainForm);
        }

        public static string[] SplitArgs(string combinedArgs, char[] delimiter)
        {
            return combinedArgs.Split(delimiter, StringSplitOptions.None);
        }
    }
}
