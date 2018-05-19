using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SkylineTool;
using WashU.BatemanLab.MassSpec.Tools.AnalysisResults;
using WashU.BatemanLab.MassSpec.Tools.AnalysisTargets;
using WashU.BatemanLab.Common;

namespace WashU.BatemanLab.MassSpec.TrackIN
{
    partial class MainForm
    {
        private Graph _graphPeptideRatios;
        private AnalysisResults _analysisResults;
        private SkylineToolClient _toolClient;
        private bool HasPeptideRatiosTabActivated = false;
        public bool IsConnectedToSkylineDoc { get; set; } = false;
        
        private void PlotChromatograms(ZedGraph.ZedGraphControl graph)
        {
            foreach (var msrun in _analysisResults.Results)
            foreach (var chromatogram in msrun.Chromatograms)
            {
                var ChromLine = graph.GraphPane.AddCurve(String.Format("{0} ({1}): [{2}] - {3}", chromatogram.Peptide, chromatogram.IsotopeLabelType, chromatogram.PrecursorMZ, "PosMatch"),
                                                                       chromatogram.RetentionTimes,
                                                                       chromatogram.SumOfPositiveMatch,
                                                                       Color.Green);
                ChromLine.Symbol.IsVisible = false;
            }

            graph.AxisChange();
            graph.Refresh();
        }

        private List<Tuple<string, string, string>> GetPossibleRatios()
        {
            IReport reportPeptideRatios = _toolClient.GetReport("BLR Peptide Ratios");

            var PeptideList = reportPeptideRatios.Cells.Where(p => p[3] != null).Select(p => p[3]).Distinct();

            var PossibleRatios = from peptideN in PeptideList
                                 from peptideD in PeptideList
                                 where peptideN != peptideD
                                 select Tuple.Create (Peptide.GetPeptideShortName(peptideN),
                                                      Peptide.GetPeptideShortName(peptideD),
                                                      String.Format("{0}/{1}", Peptide.GetPeptideShortName(peptideN), Peptide.GetPeptideShortName(peptideD)));
            return PossibleRatios.ToList();
        }

        private void BuildPeptideRatiosMenuStrips()
        {
            var mnuItems = mnuRatioSelection.DropDownItems;
            mnuItems.Clear();
            
            foreach (var variant in GetPossibleRatios())
            {
                var mnuItem = new ToolStripMenuItem(variant.Item3);

                mnuItem.Click += new EventHandler(DynamicMenuItemClicked);

                mnuItems.Add(mnuItem);
                
            }
        }

        private void DynamicMenuItemClicked(object sender, EventArgs e)
        {
            var item = (ToolStripMenuItem)sender;
            item.Checked = !item.Checked;
            BuildPeptideRatiosGraph();
        }

        private void ActivatePeptideRatiosTab()
        {
            if (IsConnectedToSkylineDoc && !HasPeptideRatiosTabActivated)
            {
                BuildPeptideRatiosMenuStrips();
                _graphPeptideRatios = new Graph(graphPeptideRatios, "MS Runs", "Peptide Ratio");
                BuildPeptideRatiosGraph();
                HasPeptideRatiosTabActivated = true;
            }
        }

        private void BuildPeptideRatiosGraph()
        {
            IReport reportPeptideRatios = _toolClient.GetReport("BLR Peptide Ratios");

            var graph = this.graphPeptideRatios;

            var graphPane = this.graphPeptideRatios.GraphPane;
            graphPane.CurveList.Clear();

            var PeptideList = reportPeptideRatios.Cells.Where(p => p[3] != null).Select(p => p[3]).Distinct();

            var PossibleRatios = from peptideN in PeptideList
                                 from peptideD in PeptideList
                                 where peptideN != peptideD
                                 select new
                                 {
                                     Nominator = Peptide.GetPeptideShortName(peptideN),
                                     Denominator = Peptide.GetPeptideShortName(peptideD),
                                     RatioName = String.Format("{0}/{1}", Peptide.GetPeptideShortName(peptideN), Peptide.GetPeptideShortName(peptideD))
                                 };

            var SelectedRatios = from mnuItem in mnuRatioSelection.DropDownItems.Cast<ToolStripMenuItem>()
                                 where mnuItem.Checked
                                 select mnuItem.Text;


            var GrouppedByPeptide = from reportRow in reportPeptideRatios.Cells
                                    group reportRow by Peptide.GetPeptideShortName(reportRow[3]) into GrouppedRows
                                    select new { Peptide = GrouppedRows.Key, Rows = GrouppedRows };


            foreach (var ratioVariant in PossibleRatios)
            {
                var Nominators = GrouppedByPeptide.Where(p => p.Peptide == ratioVariant.Nominator).Single().Rows.Where(r => r[4] != null).Select(r => r);

                var Denominators = GrouppedByPeptide.Where(p => p.Peptide == ratioVariant.Denominator).Single().Rows.Where(r => r[4] != null).Select(r => r);

                var Ratios = Nominators.Join(Denominators,
                                             n => n[0], d => d[0],
                                             (n, d) => new
                                             {
                                                 FileName = n[0],
                                                 PeptideRatio = ConvertUtil.doubleTryParse(n[4]) / ConvertUtil.doubleTryParse(d[4])
                                             });

                if (SelectedRatios.Contains(ratioVariant.RatioName))
                {
                    graphPane.AddBar(ratioVariant.RatioName, null, Ratios.Select(r => r.PeptideRatio).ToArray(), Color.Green);

                    graphPane.XAxis.Scale.TextLabels = Ratios.Select(f => (f.FileName as String).Substring(1, 25)).ToArray();
                }
            }

            graphPane.XAxis.MajorTic.IsBetweenLabels = true;
            graphPane.XAxis.Type = ZedGraph.AxisType.Text;

            graph.AxisChange();
            graph.Refresh();
        }

    }
}
