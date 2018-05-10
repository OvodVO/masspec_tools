using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WashU.BatemanLab.MassSpec.Tools.ProcessRawData;

namespace WashU.BatemanLab.MassSpec.Tools.AnalysisResults
{
    class AnalysisResults
    {
        private List<MsDataFileImplExtAgg> _analysisResults;

        public AnalysisResults()
        {
            _analysisResults = new List<MsDataFileImplExtAgg>();
        }
    }
}
