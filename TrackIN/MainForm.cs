using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SkylineTool;
using WashU.BatemanLab.MassSpec.Tools.ProcessRawData;
using WashU.BatemanLab.MassSpec.Tools.Analysis;
using WashU.BatemanLab.Common;

namespace WashU.BatemanLab.MassSpec.TrackIN
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            _analysisResults = new AnalysisResults();
        }

        public MainForm(string[] args) : this()
        {
            GetSkylineArgs(args);
            _toolClient = new SkylineTool.SkylineToolClient(_skylineConnection, "TrackIN");
            _toolClient.DocumentChanged += OnDocumentChanged;
            _toolClient.SelectionChanged += OnSelectionChanged;
            IsConnectedToSkylineDoc = true;
        }

        MsDataFileImplExtAgg _msdatafile;

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);

            try
            {
                _toolClient.DocumentChanged -= OnDocumentChanged;
                _toolClient.SelectionChanged -= OnSelectionChanged;
                _toolClient.Dispose();
            }
            catch
            { }

            _toolClient = null;
        }

        private void OnDocumentChanged(object sender, EventArgs eventArgs)
        {
            // Create graph on UI thread.
            // Invoke(new Action(CreateGraph));
        }

        private void OnSelectionChanged(object sender, EventArgs eventArgs)
        {
            var ReplicateName = _toolClient.GetReplicateName();
            var DocumentLocation = _toolClient.GetDocumentLocation();
            var DocumentLocationName = _toolClient.GetDocumentLocationName();
            // Invoke(new Action(() => MessageBox.Show(ReplicateName, "GetReplicateName()")));
            // Invoke(new Action(() => MessageBox.Show(DocumentLocationName, "DocumentLocationName()")));
        }

        private void btnTEST_Click(object sender, EventArgs e)
        {
            Stopwatch watch = new Stopwatch();
            TimeSpan[] TimesToPerform = new TimeSpan[3];
            double Toleranse = 0.1;
            OpenFileDialog _openDlg = new OpenFileDialog();
            _openDlg.Filter = "Thermo(*.raw,|*.raw;)";

            if (_openDlg.ShowDialog() == DialogResult.OK)
            {
                watch.Start();
                _msdatafile = new MsDataFileImplExtAgg(_openDlg.FileName);
                TimesToPerform[0] = TimeSpan.FromMilliseconds(watch.ElapsedMilliseconds);
            }
            else return;
            watch.Reset();
            watch.Start();
            _msdatafile.GetMsDataSpectrums();
            TimesToPerform[1] = TimeSpan.FromMilliseconds(watch.ElapsedMilliseconds);
            watch.Reset();
            watch.Start();
            _msdatafile.GetChromatograms(GetProteinsFromSkyline(), Toleranse);
            TimesToPerform[2] = TimeSpan.FromMilliseconds(watch.ElapsedMilliseconds);
            MessageBox.Show(String.Format("Read: {0}:{1}; GetSpectrums: {2}:{3}; GetChromatograms: {4}:{5} ",
                                           TimesToPerform[0].Minutes, TimesToPerform[0].Seconds,
                                           TimesToPerform[1].Minutes, TimesToPerform[1].Seconds,
                                           TimesToPerform[2].Minutes, TimesToPerform[2].Seconds));
            foreach (var chromatogram in _msdatafile.Chromatograms)
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

        private void btnTEST2_Click(object sender, EventArgs e)
        {
            // Test only
        }

        private void tabPeptideRatios_Enter(object sender, EventArgs e)
        {
            ActivatePeptideRatiosTab();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Properties.Settings.Default.PeptideRatio != _defaultPeptideRatioName)
                Properties.Settings.Default.PeptideRatio = _defaultPeptideRatioName;
            Properties.Settings.Default.Save();
        }

        private async void mnuAnalizeFromSkyline_Click(object sender, EventArgs e)
        {
            await ReadAndAnalyzeSet(GetMSFilesFromSkyline().Select(f => f.Split('*')[1]).ToArray());
        }

        private async void mnuNASelectMSRunsToAnalyze_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDlg = new OpenFileDialog();
            openDlg.Filter = "Thermo(*.raw,|*.raw;)";
            openDlg.Multiselect = true;
            if (openDlg.ShowDialog() == DialogResult.OK)
            {
               await ReadAndAnalyzeSet(openDlg.FileNames);
            }
        }

        private void btnMakeLinkerToSkyline_Click(object sender, EventArgs e)
        {
            CreateSkylineLinker();
        }
    }
}
  