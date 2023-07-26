using System;
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
using WashU.BatemanLab.MassSpec.Tools.Analysis;
using WashU.BatemanLab.Common;

namespace WashU.BatemanLab.MassSpec.TrackIN
{

    partial class MainForm
    {
        private string _defaultPeptideRatioName;
        private string _defaultPeptideName;
        private Dictionary<string, double> _peptideIntStdConcentrations = new Dictionary<string, double>()
            {
                {"Aβ42", 1},
                {"Aβ40", 10},
                {"Aβ38", 1.5},
                {"Aβ[Total]", 12.5}
            };

        public List<string> SelectedMeasures
        {
            get
            {
                return (from mnuItem in mnuNAMeasure.DropDownItems.Cast<ToolStripMenuItem>()
                       where mnuItem.Checked
                       select mnuItem.Text).ToList();
            }
        }

        private Graph _graphPeptideRatios;
        private NAmsrunsTabControl _tabControlMsRuns;
        private AnalysisResults _analysisResults;
        private SkylineToolClient _toolClient;

        private string _skylineConnection
        {
            get { return SkylineArgs["SkylineConnection"]; }
            set { SkylineArgs["SkylineConnection"] = value; }
        }
        public Dictionary<string, string> SkylineArgs;
        private bool HasPeptideRatiosTabActivated = false;
        private bool HasNoiseAnalysisTabActivated = false;
        private bool HasPrepForSkylineTabActivated = false;

        public bool IsConnectedToSkylineDoc { get; set; } = false;

        public void GetSkylineArgs(string[] args)
        {
            char[] separator = { '=' };
            SkylineArgs = args.Select(a => new { argName = a.Split(separator)[0], argValue = a.Split(separator)[1] }).ToDictionary(di => di.argName, di => di.argValue);
        }
        
        private List<string> GetPeptideListFromAnalysis()
        {
            var _peptides = (from _protein in _analysisResults.AnalysisTargets.Proteins
                             from _peptide in _protein.Peptides
                             select _peptide.Name.Trim()).Distinct();
            return _peptides.ToList();
        }

        private void CompleteAnalysisTasks()
        {
            BuildShownPeptideMenuStrips();
            Invoke( new Action(() => BuildNoiseAnalysisPlots()) );
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
            var mnuItems = mnuPRratioSelection.DropDownItems;
            mnuItems.Clear();
            foreach (var variant in GetPossibleRatios())
            {
                var mnuItem = new ToolStripMenuItem(variant.Item3);
                mnuItem.Click += new EventHandler(PRDynamicMenuItemClicked);
                if (variant.Item3 == _defaultPeptideRatioName) mnuItem.Checked = true;
                mnuItems.Add(mnuItem);
            }
        }

        private void BuildShownPeptideMenuStrips()
        {
            var mnuItems = mnuNAPeptide.DropDownItems;
            mnuItems.Clear();
            foreach (var peptide in GetPeptideListFromAnalysis())
            {
                var mnuItem = new ToolStripMenuItem(peptide);
                mnuItem.Click += new EventHandler(NADynamicMenuItemClicked);
                if (peptide == _defaultPeptideRatioName) mnuItem.Checked = true;
                mnuItems.Add(mnuItem);
            }
        }

        private void BuildShownMeasuresMenuStrips()
        {
            var mnuItems = mnuNAMeasure.DropDownItems;
            mnuItems.Clear();
            foreach (var measure in Chromatogramm.MeasuresNameList)
            {
                var mnuItem = new ToolStripMenuItem(measure);
                mnuItem.Click += new EventHandler(NADynamicMenuItemClicked);
                if (Chromatogramm.MeasuresNameListByDefault.Contains( measure )) mnuItem.Checked = true;
                mnuItems.Add(mnuItem);
            }
        }

        private void PRDynamicMenuItemClicked(object sender, EventArgs e)
        {
            var item = (ToolStripMenuItem)sender;
            item.Checked = !item.Checked;
            BuildPeptideRatiosGraph();
        }

        private void NADynamicMenuItemClicked(object sender, EventArgs e)
        {
            var item = (ToolStripMenuItem)sender;
            item.Checked = !item.Checked;
            ModifyNoiseAnalysisPlots();
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

        private void ActivateNoiseAnalysisTab()
        {
            _defaultPeptideName = Properties.Settings.Default.PeptideName;
            if ( !HasNoiseAnalysisTabActivated )
            {
                BuildShownMeasuresMenuStrips();
                HasNoiseAnalysisTabActivated = true;
            }
        }

        private void ActivateLCTracersTab()
        {
            lbTargetDir.Text = @"r:\_Projects\Aß\Plasma\Clinical Studies\BioFINDER\LC Data source BF-1";

          //  _defaultPeptideName = Properties.Settings.Default.PeptideName;
          //  if (!HasNoiseAnalysisTabActivated)
          //  {
          //      BuildShownMeasuresMenuStrips();
          //      HasNoiseAnalysisTabActivated = true;
          //  }
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
            var SelectedRatios = from mnuItem in mnuPRratioSelection.DropDownItems.Cast<ToolStripMenuItem>()
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

        private List<Protein> GetProteinsFromSkyline()
        {
            var result = new List<Protein>();
            IReport reportTrackINTargets = _toolClient.GetReport("BLR TrackIN Targets");
            var ProteinsQ =  from reportRow in reportTrackINTargets.Cells
                             where string.IsNullOrEmpty(reportRow[0]) != true
                             group reportRow by reportRow[0] into ProteintGroup
                             select new
                             {
                                 Protein = ProteintGroup.Key,
                                 Peptides = from reportRow in ProteintGroup
                                            group reportRow by reportRow[1] into PeptideGroup
                                            select new
                                            {
                                                Peptide = PeptideGroup.Key,
                                                Precursors = from reportRow in PeptideGroup
                                                             group reportRow by new { Isotope = reportRow[2], Precursor = reportRow[3] } into PrecursorGroup
                                                             select new
                                                             {
                                                                 Isotope = PrecursorGroup.Key.Isotope,
                                                                 PrecursorMZ = PrecursorGroup.Key.Precursor,
                                                                 ProductMZ = from reportRow in PrecursorGroup
                                                                             select reportRow[4]
                                                             }
                                            }
                             };
            foreach (var prot in ProteinsQ)
            {
                Protein protein = new Protein();
                protein.Name = prot.Protein;
                foreach (var pept in prot.Peptides)
                {
                    Peptide peptide = new Peptide();
                    peptide.Name = pept.Peptide;
                    foreach (var prec in pept.Precursors)
                    {
                        Precursor precursor = new Precursor();
                        precursor.IsotopeLabelType = prec.Isotope;
                        precursor.PrecursorMZ = Convert.ToDouble(prec.PrecursorMZ);
                        foreach (var prod in prec.ProductMZ)
                        {
                            precursor.Products.Add(Convert.ToDouble(prod));
                        }
                        peptide.Precursors.Add(precursor);
                    }
                    protein.Peptides.Add(peptide);
                }
                result.Add(protein);
            }
            return result;
        }

        private string[] GetMSFilesFromSkyline()
        {
            IReport reportMSRunsforTrackIN = _toolClient.GetReport("BLR MS Runs for TrackIN analysis");
            var MSRanNames = from reportRow in reportMSRunsforTrackIN.Cells
                             where string.IsNullOrEmpty(reportRow[0]) != true
                             select String.Format("{0}*{1}", reportRow[0], reportRow[1]);
            return MSRanNames.ToArray();
        }

        private async Task ReadAndAnalyzeSetAsync(string[] files)
        {
            if (this.IsConnectedToSkylineDoc)
            { _analysisResults.AnalysisTargets.Proteins = AnalysisTargets.GetDefaultProteins(); }  //GetProteinsFromSkyline(); }
            else
            { _analysisResults.AnalysisTargets.Proteins = AnalysisTargets.GetDefaultProteins(); }

            var ImportsAsync = files.Select(filename => Task.Factory.StartNew(async () =>
            {
                try
                {
                    await _analysisResults.ReadAndAnalyzeMSFile(filename);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + " Cannot import file - " + filename);
                }
                }));
                await Task.Factory.ContinueWhenAll(ImportsAsync.ToArray(), results => CompleteAnalysisTasks());
        }

        private void BuildNoiseAnalysisPlots1()
        {
            _tabControlMsRuns = new NAmsrunsTabControl();
            foreach (var result in _analysisResults.Results)
            {
                TabPage tabMsRunPage = new TabPage();
                tabMsRunPage.Name = result.MSrunFileName;
                tabMsRunPage.Text = result.MSrunFileName;
                NApeptidesTabControl _tabControlPeptide = new NApeptidesTabControl();

                foreach (var peptide in GetPeptideListFromAnalysis())
                {
                    TabPage tabPeptidePage = new TabPage();
                    tabPeptidePage.Name = peptide;
                    tabPeptidePage.Text = Peptide.GetPeptideShortName(peptide);

                    SplitContainer splPeptidePage = new SplitContainer();
                    splPeptidePage.Dock = DockStyle.Fill;
                    splPeptidePage.Orientation = Orientation.Horizontal;
                    splPeptidePage.BorderStyle = BorderStyle.None;

                    NoiseMasspecDataGraph graphAnalyte = new NoiseMasspecDataGraph("Retention Time", "Intensity", "Time, miliS");
                    graphAnalyte.GraphControl.Dock = DockStyle.Fill;

                    NoiseMasspecDataGraph graphISTD = new NoiseMasspecDataGraph("Retention Time", "Intensity", "Time, miliS");
                    graphISTD.GraphControl.Dock = DockStyle.Fill;

                    var analyteChroms = (from chrom in result.Chromatograms
                                         where chrom.Peptide == peptide && chrom.IsotopeLabelType == "light"
                                         select chrom).Single();

                    var istdChroms = (from chrom in result.Chromatograms
                                      where chrom.Peptide == peptide && chrom.IsotopeLabelType == "N15 ISTD"
                                      select chrom).Single();


                    foreach (var _measure in Chromatogramm.MeasuresDictionary)
                    {
                        var lineAnalyte = graphAnalyte.GraphControl.GraphPane.AddCurve
                                 (_measure.Key, analyteChroms.RetentionTimes,
                                  analyteChroms.GetMeasureByName(_measure.Key), _measure.Value);

                        var lineISTD = graphISTD.GraphControl.GraphPane.AddCurve
                                 (_measure.Key, istdChroms.RetentionTimes,
                                 istdChroms.GetMeasureByName(_measure.Key), _measure.Value);

                        lineAnalyte.Symbol.IsVisible = false;
                        lineAnalyte.Line.Width = 3.0F;
                        lineISTD.Symbol.IsVisible = false;
                        lineISTD.Line.Width = 3.0F;

                        if (_measure.Key == "IIT")
                        {
                            lineAnalyte.IsY2Axis = true; lineISTD.IsY2Axis = true;
                        }

                        if (SelectedMeasures.Contains(_measure.Key))
                        {
                            if (_measure.Key == "IIT")
                            {
                                graphAnalyte.GraphControl.GraphPane.Y2Axis.IsVisible = true;
                                graphISTD.GraphControl.GraphPane.Y2Axis.IsVisible = true;
                            }
                            lineAnalyte.IsVisible = true; lineAnalyte.Label.IsVisible = true;
                            lineISTD.IsVisible = true; lineISTD.Label.IsVisible = true;
                        }
                        else
                        {
                            if (_measure.Key == "IIT")
                            {
                                graphAnalyte.GraphControl.GraphPane.Y2Axis.IsVisible = false;
                                graphISTD.GraphControl.GraphPane.Y2Axis.IsVisible = false;
                            }
                            lineAnalyte.IsVisible = false; lineAnalyte.Label.IsVisible = false;
                            lineISTD.IsVisible = false; lineISTD.Label.IsVisible = false;
                        }
                    }

                    graphAnalyte.GraphControl.GraphPane.XAxis.Type = ZedGraph.AxisType.Linear;
                    graphISTD.GraphControl.GraphPane.XAxis.Type = ZedGraph.AxisType.Linear;
                    splPeptidePage.Panel1.Controls.Add(graphAnalyte.GraphControl);
                    graphAnalyte.GraphControl.AxisChange();
                    graphAnalyte.GraphControl.Refresh();
                    splPeptidePage.Panel1.Controls.Add(graphISTD.GraphControl);
                    graphISTD.GraphControl.AxisChange(); graphISTD.GraphControl.Refresh(); //graphISTD.GraphControl.
                    tabPeptidePage.Controls.Add(splPeptidePage);
                    _tabControlPeptide.TabPages.Add(tabPeptidePage);
                }
                tabMsRunPage.Controls.Add(_tabControlPeptide);
                _tabControlMsRuns.TabPages.Add(tabMsRunPage);
            }
            splNAtab.Panel2.Controls.Add(_tabControlMsRuns);
        }

        private void BuildNoiseAnalysisPlots()
        {
            _tabControlMsRuns = new NAmsrunsTabControl();
            foreach (var result in _analysisResults.Results)
            {
                TabPage tabMsRunPage = new TabPage();
                tabMsRunPage.Name = result.MSrunFileName;
                tabMsRunPage.Text = result.MSrunFileName;
                NApeptidesTabControl _tabControlPeptide = new NApeptidesTabControl();

                foreach (var peptide in GetPeptideListFromAnalysis())
                {
                    TabPage tabPeptidePage = new TabPage();
                    tabPeptidePage.Name = peptide;
                    tabPeptidePage.Text = Peptide.GetPeptideShortName(peptide);

                    SplitContainer splPeptidePage = new SplitContainer();
                    splPeptidePage.Dock = DockStyle.Fill;
                    splPeptidePage.Orientation = Orientation.Horizontal;
                    splPeptidePage.BorderStyle = BorderStyle.None;

                    splPeptidePage.Panel2Collapsed = true;

                    NoiseMasspecDataGraph graphAnalyte = new NoiseMasspecDataGraph("Retention Time", "Intensity", "Time, miliS");
                    graphAnalyte.GraphControl.Dock = DockStyle.Fill;

                    //ZedGraph.GraphPane ISTDGraphPane = new ZedGraph.GraphPane();


                    var ISTDGraphPane = graphAnalyte.GraphControl.MasterPane.PaneList["ISTD"];

                    //graphAnalyte.GraphControl.MasterPane.Add(ISTDGraphPane);

                    NoiseMasspecDataGraph graphISTD = new NoiseMasspecDataGraph("Retention Time", "Intensity", "Time, miliS");
                    graphISTD.GraphControl.Dock = DockStyle.Fill;

                    var analyteChroms = (from chrom in result.Chromatograms
                                         where chrom.Peptide == peptide && chrom.IsotopeLabelType == "light"
                                         select chrom).Single();

                    var istdChroms = (from chrom in result.Chromatograms
                                      where chrom.Peptide == peptide && chrom.IsotopeLabelType == "N15 ISTD"
                                      select chrom).Single();


                    foreach (var _measure in Chromatogramm.MeasuresDictionary)
                    {
                        var lineAnalyte = graphAnalyte.GraphControl.GraphPane.AddCurve
                                 (_measure.Key, analyteChroms.RetentionTimes,
                                  analyteChroms.GetMeasureByName(_measure.Key), _measure.Value);

                        

                        var lineISTD = ISTDGraphPane.AddCurve
                                 (_measure.Key, istdChroms.RetentionTimes,
                                 istdChroms.GetMeasureByName(_measure.Key), _measure.Value);

                        lineAnalyte.Symbol.IsVisible = false;
                        lineAnalyte.Line.Width = 3.0F;
                        lineISTD.Symbol.IsVisible = false;
                        lineISTD.Line.Width = 3.0F;

                        if (_measure.Key == "IIT")
                        {
                            lineAnalyte.IsY2Axis = true; lineISTD.IsY2Axis = true;
                        }

                        if (SelectedMeasures.Contains(_measure.Key))
                        {
                            if (_measure.Key == "IIT")
                            {
                                graphAnalyte.GraphControl.GraphPane.Y2Axis.IsVisible = true;

                                ISTDGraphPane.Y2Axis.IsVisible = true;

                                graphISTD.GraphControl.GraphPane.Y2Axis.IsVisible = true;
                            }
                            lineAnalyte.IsVisible = true; lineAnalyte.Label.IsVisible = true;
                            lineISTD.IsVisible = true; lineISTD.Label.IsVisible = true;
                        }
                        else
                        {
                            if (_measure.Key == "IIT")
                            {
                                graphAnalyte.GraphControl.GraphPane.Y2Axis.IsVisible = false;

                                ISTDGraphPane.Y2Axis.IsVisible = false;

                                //graphAnalyte.GraphControl.MasterPane.PaneList["ISTD"].Y2Axis.IsVisible = false;

                                graphISTD.GraphControl.GraphPane.Y2Axis.IsVisible = false;
                            }
                            lineAnalyte.IsVisible = false; lineAnalyte.Label.IsVisible = false;
                            lineISTD.IsVisible = false; lineISTD.Label.IsVisible = false;

                            //ISTDGraphPane.Y2Axis.IsVisible = false;
                        }
                    }

                    graphAnalyte.GraphControl.GraphPane.XAxis.Type = ZedGraph.AxisType.Linear;
                    graphISTD.GraphControl.GraphPane.XAxis.Type = ZedGraph.AxisType.Linear;
                    splPeptidePage.Panel1.Controls.Add(graphAnalyte.GraphControl);

                    graphAnalyte.GraphControl.MasterPane.SetLayout(splPeptidePage.Panel1.CreateGraphics(), ZedGraph.PaneLayout.SingleColumn);


                    graphAnalyte.GraphControl.AxisChange();
                    graphAnalyte.GraphControl.Refresh();
                    splPeptidePage.Panel2.Controls.Add(graphISTD.GraphControl);
                    graphISTD.GraphControl.AxisChange(); graphISTD.GraphControl.Refresh(); //graphISTD.GraphControl.
                    tabPeptidePage.Controls.Add(splPeptidePage);
                    _tabControlPeptide.TabPages.Add(tabPeptidePage);
                }
                tabMsRunPage.Controls.Add(_tabControlPeptide);
                _tabControlMsRuns.TabPages.Add(tabMsRunPage);
            }
            splNAtab.Panel2.Controls.Add(_tabControlMsRuns);
        }
        private void ModifyNoiseAnalysisPlots1()
        {
            var AnalyteGraphsForMsRuns = from _msrunTabPage in _tabControlMsRuns.TabPages.Cast<TabPage>()
                                         from _petideTabContolPage in _msrunTabPage.Controls.OfType<NApeptidesTabControl>()
                                         from _peptideTabPage in _petideTabContolPage.TabPages.Cast<TabPage>()
                                         from _splitContainer in _peptideTabPage.Controls.OfType<SplitContainer>()
                                         from _graph in _splitContainer.Panel1.Controls.OfType<ZedGraph.ZedGraphControl>()
                                         select _graph;

            var ISTDGraphsForMsRuns = from _msrunTabPage in _tabControlMsRuns.TabPages.Cast<TabPage>()
                                      from _petideTabContolPage in _msrunTabPage.Controls.OfType<NApeptidesTabControl>()
                                      from _peptideTabPage in _petideTabContolPage.TabPages.Cast<TabPage>()
                                      from _splitContainer in _peptideTabPage.Controls.OfType<SplitContainer>()
                                      from _graph in _splitContainer.Panel2.Controls.OfType<ZedGraph.ZedGraphControl>()
                                      select _graph;

            var AllGraphsForMsRuns = AnalyteGraphsForMsRuns.Concat(ISTDGraphsForMsRuns);


            foreach (var GraphOfRun in AllGraphsForMsRuns.Cast<ZedGraph.ZedGraphControl>())
            {
                var CurvesOfGraph = from _curve in GraphOfRun.GraphPane.CurveList
                                    select _curve;


                foreach (var Curve in CurvesOfGraph)
                {
                    if (SelectedMeasures.Contains(Curve.Label.Text))
                    {
                        Curve.IsVisible = true; Curve.Label.IsVisible = true;
                        if (Curve.Label.Text=="IIT") { GraphOfRun.GraphPane.Y2Axis.IsVisible = true; }
                    }
                    else
                    {
                        Curve.IsVisible = false; Curve.Label.IsVisible = false;
                        if (Curve.Label.Text == "IIT") { GraphOfRun.GraphPane.Y2Axis.IsVisible = false; }
                    }
                }

                GraphOfRun.AxisChange(); GraphOfRun.Refresh();

            }

        }

        private void ModifyNoiseAnalysisPlots()
        {
            var AnalyteGraphsForMsRuns = from _msrunTabPage in _tabControlMsRuns.TabPages.Cast<TabPage>()
                                         from _petideTabContolPage in _msrunTabPage.Controls.OfType<NApeptidesTabControl>()
                                         from _peptideTabPage in _petideTabContolPage.TabPages.Cast<TabPage>()
                                         from _splitContainer in _peptideTabPage.Controls.OfType<SplitContainer>()
                                         from _graph in _splitContainer.Panel1.Controls.OfType<ZedGraph.ZedGraphControl>()
                                         select _graph;

            var ISTDGraphsForMsRuns = from _msrunTabPage in _tabControlMsRuns.TabPages.Cast<TabPage>()
                                      from _petideTabContolPage in _msrunTabPage.Controls.OfType<NApeptidesTabControl>()
                                      from _peptideTabPage in _petideTabContolPage.TabPages.Cast<TabPage>()
                                      from _splitContainer in _peptideTabPage.Controls.OfType<SplitContainer>()
                                      from _graph in _splitContainer.Panel2.Controls.OfType<ZedGraph.ZedGraphControl>()
                                      select _graph;

            var AllGraphsForMsRuns = AnalyteGraphsForMsRuns.Concat(ISTDGraphsForMsRuns);


            foreach (var GraphOfRun in AllGraphsForMsRuns.Cast<ZedGraph.ZedGraphControl>())
            {
                var CurvesOfGraph  = from _pane in GraphOfRun.MasterPane.PaneList
                                     from _curve in _pane.CurveList
                                     select _curve;

                foreach (var Curve in CurvesOfGraph)
                {
                    if (SelectedMeasures.Contains(Curve.Label.Text))
                    {
                        Curve.IsVisible = true; Curve.Label.IsVisible = true;
                        if (Curve.Label.Text == "IIT") { GraphOfRun.MasterPane.PaneList.ForEach((p) => p.Y2Axis.IsVisible = true); }
                    }
                    else
                    {
                        Curve.IsVisible = false; Curve.Label.IsVisible = false;
                        if (Curve.Label.Text == "IIT") { GraphOfRun.MasterPane.PaneList.ForEach((p) => p.Y2Axis.IsVisible = false); }
                    }
                }
                GraphOfRun.AxisChange(); GraphOfRun.Refresh();

            }

        }
    }
}
