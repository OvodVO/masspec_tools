using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WashU.BatemanLab.MassSpec.Tools.AnalysisTargets;
using WashU.BatemanLab.MassSpec.Tools.ProcessRawData;

namespace WashU.BatemanLab.MassSpec.Tools.AnalysisResults
{
    public class AnalysisResults
    {
        private List<MsDataFileImplExtAgg> _analysisResults;

        public static string GetMSRunShorten(string msRunName, string format)
        {
            string result = "";
            char separator = '_';

            string[] msRunSplit = msRunName.Split('_');
            string[] formatSplit = format.Split(',');
            int count = 0;
            for (int i = 0; i < formatSplit.Length; i++)
            {
                int index;
                if (int.TryParse(formatSplit[i], out index))
                {
                    if (index < msRunSplit.Length)
                    {
                        if (count > 0 ) { result += separator; }
                        result += msRunSplit[index];
                        count++;
                    }
                };
            }
            return result.Replace(".raw", "");
        }

        public List<MsDataFileImplExtAgg> Results
        {
            get { return _analysisResults; }
        }

        public AnalysisResults()
        {
            _analysisResults = new List<MsDataFileImplExtAgg>();
        }

        public async Task LoadAnalysisResults(string[] msfilesToLoad)
        {           
            foreach (string path in msfilesToLoad)
            {
                /*await Task.Factory.StartNew( () => { _analysisResults.Add(new MsDataFileImplExtAgg(path)); } );*/
                _analysisResults.Add(new MsDataFileImplExtAgg(path));
            }
        }

        public void PerformAnalysis()
        {
            foreach (MsDataFileImplExtAgg msrun in _analysisResults)
            {
                msrun.GetMsDataSpectrums();
                msrun.GetChromatograms(0.1);
            }
        }
        public async Task PerformAnalysis(List<Protein> targets)
        {
            foreach (MsDataFileImplExtAgg msrun in _analysisResults)
            {
                /* await Task.Factory.StartNew(() => { msrun.GetMsDataSpectrums(); })
                     .ContinueWith( (a) => { msrun.GetChromatograms(targets, 0.1); }); */
                 msrun.GetMsDataSpectrums();
                 msrun.GetChromatograms(targets, 0.1);
            }
        }
    }
}
