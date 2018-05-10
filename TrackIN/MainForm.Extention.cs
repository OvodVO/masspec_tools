using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WashU.BatemanLab.MassSpec.Tools.AnalysisResults;

namespace WashU.BatemanLab.MassSpec.TrackIN
{
    partial class MainForm
    {
        private AnalysisResults _analysisResults;

        private void PlotChromatograms()
        {
            foreach (var msrun in _analysisResults.Results)
            foreach (var chromatogram in msrun.Chromatograms)
            {
                var ChromLine = zedGraphControlTest.GraphPane.AddCurve(String.Format("{0} ({1}): [{2}] - {3}", chromatogram.Peptide, chromatogram.IsotopeLabelType, chromatogram.PrecursorMZ, "PosMatch"),
                                                                       chromatogram.RetentionTimes,
                                                                       chromatogram.SumOfPositiveMatch,
                                                                       Color.Green);
                ChromLine.Symbol.IsVisible = false;
            }

            zedGraphControlTest.AxisChange();
            zedGraphControlTest.Refresh();
        }

    }
}
