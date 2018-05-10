using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WashU.BatemanLab.MassSpec.Tools.ProcessRawData;

namespace WashU.BatemanLab.MassSpec.Tools.AnalysisResults
{
    public class AnalysisResults
    {
        private List<MsDataFileImplExtAgg> _analysisResults;

        public List<MsDataFileImplExtAgg> Results
        {
            get { return _analysisResults; }
        }

        public AnalysisResults()
        {
            _analysisResults = new List<MsDataFileImplExtAgg>();
        }

        public void LoadAnalysisResults(string[] msfilesToLoad)
        {           
            foreach (string path in msfilesToLoad)
            {
                _analysisResults.Add( new MsDataFileImplExtAgg(path) );
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
    }
}
