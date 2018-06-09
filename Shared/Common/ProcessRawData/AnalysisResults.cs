using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WashU.BatemanLab.Common;
using pwiz.ProteowizardWrapper;
using WashU.BatemanLab.MassSpec.Tools.ProcessRawData;


namespace WashU.BatemanLab.MassSpec.Tools.Analysis
{
    public class AnalysisResults
    {
        private List<MsDataFileImplExtAgg> _analysisResults;
        private AnalysisTargets _analysisTargets;
        public AnalysisTargets AnalysisTargets { get { return _analysisTargets; } set { _analysisTargets = value; } }

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
            _analysisTargets = new AnalysisTargets();
        }

        public Task<bool> ReadAndAnalyzeMSFile(string path)
        {
            var IsCompleted = false;
            var c = _analysisResults.Count;
            var progress = new ReadAndAnalyzeProgressInfo();
            var msrun = new MsDataFileImplExtAgg(path);
             msrun.GetMsDataSpectrums();
             msrun.GetChromatograms(_analysisTargets.Proteins, 0.1);
             _analysisResults.Add(msrun);
            if (_analysisResults.Count - c == 1)
                IsCompleted = true;
            return Task.FromResult(IsCompleted); 
        }
    }
}
