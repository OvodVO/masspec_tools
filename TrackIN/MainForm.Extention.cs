﻿using System;
using System.Configuration;
using System.Collections.Generic;
using System.Drawing;
using System.Deployment;
using System.Deployment.Application;
using System.IO;
using System.Linq;
using System.Reflection;
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

        private string _defaultPeptideRatioName; // = Properties.Settings.Default.PeptideRatio;
        private Dictionary<string, double> _peptideIntStdConcentrations = new Dictionary<string, double>()
            {
                {"Aβ42", 1},
                {"Aβ40", 10},
                {"Aβ38", 1.5},
                {"Aβ[Total]", 12.5}
            };
        private Graph _graphPeptideRatios;
        private AnalysisResults _analysisResults;
        private SkylineToolClient _toolClient;
        private string _skylineConnection
        {
            get { return SkylineArgs["SkylineConnection"]; }
            set { SkylineArgs["SkylineConnection"] = value; }
        }

        public Dictionary<string, string> SkylineArgs;
        private bool HasPeptideRatiosTabActivated = false;
        public bool IsConnectedToSkylineDoc { get; set; } = false;

        public void GetSkylineArgs(string[] args)
        {
            char[] separator = { '=' };
            SkylineArgs = args.Select(a => new { argName = a.Split(separator)[0], argValue = a.Split(separator)[1] }).ToDictionary(di => di.argName, di => di.argValue);
        }

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
                                 select Tuple.Create(Peptide.GetPeptideShortName(peptideN),
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

                if (variant.Item3 == _defaultPeptideRatioName) mnuItem.Checked = true;

                mnuItems.Add(mnuItem);
            }

            //if (mnuRatioSelection.DropDownItems.Cast<ToolStripMenuItem>().Where(m => m.Checked == true).Count() < 1)
            //  mnuRatioSelection.DropDownItems.Cast<ToolStripMenuItem>().FirstOrDefault().Checked = true;
        }

        private void DynamicMenuItemClicked(object sender, EventArgs e)
        {
            var item = (ToolStripMenuItem)sender;
            item.Checked = !item.Checked;
            BuildPeptideRatiosGraph();
        }

        private void ActivatePeptideRatiosTab()
        {
            _defaultPeptideRatioName = Properties.Settings.Default.PeptideRatio;

            if (IsConnectedToSkylineDoc && !HasPeptideRatiosTabActivated)
            {
                BuildPeptideRatiosMenuStrips();
                _graphPeptideRatios = new RatioGraph(graphPeptideRatios, "MS Runs", "Peptide Ratio");
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

            Color[] graphColors = new Color[] { Color.Red, Color.Blue, Color.Green, Color.Gray };

            Queue<Color> queColors = new Queue<Color>(graphColors);


            var PeptideList = reportPeptideRatios.Cells.Where(p => p[3] != null).Select(p => p[3]).Distinct();

            var PossibleRatios = from peptideN in PeptideList.Select(p => Peptide.GetPeptideShortName(p))
                                 from peptideD in PeptideList.Select(p => Peptide.GetPeptideShortName(p))
                                 where peptideN != peptideD
                                       && _peptideIntStdConcentrations.ContainsKey(peptideN)
                                       && _peptideIntStdConcentrations.ContainsKey(peptideD)
                                 select new
                                 {
                                     Nominator = peptideN,
                                     Denominator = peptideD,
                                     RatioName = String.Format("{0}/{1}", peptideN, peptideD),
                                     CorCoef = _peptideIntStdConcentrations[peptideN] / _peptideIntStdConcentrations[peptideD]
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
                                                 PeptideRatio = ConvertUtil.doubleTryParse(n[4]) / ConvertUtil.doubleTryParse(d[4]) * ratioVariant.CorCoef
                                             });

                if (SelectedRatios.Contains(ratioVariant.RatioName))
                {
                    graphPane.AddBar(ratioVariant.RatioName, null, Ratios.Select(r => r.PeptideRatio).ToArray(), queColors.Dequeue());

                    graphPane.XAxis.Scale.TextLabels = Ratios.Select(f => AnalysisResults.GetMSRunShorten(f.FileName, "0, 5")).ToArray();
                }
            }

            graphPane.XAxis.MajorTic.IsBetweenLabels = true;
            graphPane.XAxis.Type = ZedGraph.AxisType.Text;

            graph.AxisChange();
            graph.Refresh();
        }

        public static void CreateSkylineLinker()
        {
            string linkFileName = "TrackIN";
            string linkExtention = "cmd";
            string publisherName = "Bateman Lab";
            List<string> commands = new List<string>();

            var linkFilePath = String.Format("{0}\\{1}\\{2}.{3}",
                                              Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                              linkFileName,
                                              linkFileName, linkExtention);

            var command_1 = String.Format("\"{0}\\{1}\\{2}.appref-ms\" %*",
                                          Environment.GetFolderPath(Environment.SpecialFolder.Programs),
                                          publisherName,
                                          linkFileName);
            commands.Add(command_1);
            string path = Path.GetDirectoryName(linkFilePath);
            if ( !Directory.Exists(path) )
            {
                Directory.CreateDirectory(path);
            }
            File.WriteAllLines(linkFilePath, commands);
        }

    }
}
