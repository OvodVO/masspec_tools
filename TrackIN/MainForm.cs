using System;
using System.IO;
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
            var t = GetMSFilesFromSkyline().Select(f => f.Split('*')[1]).ToArray();
            MessageBox.Show(t.Count().ToString());
            listBox1.Items.AddRange(GetMSFilesFromSkyline().Select(f => f.Split('*')[1]).ToArray());

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

        private void tabMainForm_Enter(object sender, EventArgs e)
        {
            ActivateNoiseAnalysisTab();
        }

        private void btnPasteFromExcel_Click(object sender, EventArgs e)
        {
            string s = Clipboard.GetText();

            string[] lines = s.Replace("\n", "").Split('\r');

            string[] columns = lines[0].Split('\t');
            foreach (string column in columns )
            {
                dgSampleList.Columns.Add( column, column );
            }

            dgSampleList.Rows.Add(lines.Length - 2);
            string[] fields;
            int row = 0;
            int col = 0;

            foreach (string item in lines.Skip(1))
            {
                fields = item.Split('\t');
                foreach (string f in fields)
                {
                    //Console.WriteLine(f);
                    dgSampleList[col, row].Value = f;
                    col++;
                }
                row++;
                col = 0;
            }

        }

        private void btnGenSequence_Click(object sender, EventArgs e)
        {
            string f_date = "20180717";
            string s_pos = ""; string b_pos = "1:F,8";
            string s_batch = "Batch_01";

            Queue<string> tray = new Queue<string>(new string[] { "1:A,1", "1:A,2", "1:A,3", "1:A,4", "1:A,5", "1:A,6", "1:A,7", "1:A,8",
                                                                  "1:B,1", "1:B,2", "1:B,3", "1:B,4", "1:B,5", "1:B,6", "1:B,7", "1:B,8",
                                                                  "1:C,1", "1:C,2", "1:C,3", "1:C,4", "1:C,5", "1:C,6", "1:C,7", "1:C,8",
                                                                  "1:D,1", "1:D,2", "1:D,3", "1:D,4", "1:D,5", "1:D,6", "1:D,7", "1:D,8",
                                                                  "1:E,1", "1:E,2", "1:E,3", "1:E,4", "1:E,5", "1:E,6", "1:E,7", "1:E,8",
                                                                  "1:F,1", "1:F,2", "1:F,3", "1:F,4", "1:F,5", "1:F,6", "1:F,7", "1:F,8",});

            string s_inst_meth = "D:\\OTFLumos\\Plasma-Aß\\_methods\\20180703_Ab-noC13_HSS-T3-75x100_AVG-mass_300ulpmin_VO5";
            string s_path = @"D:\OTFLumos\Plasma-Aß\20180718_ADRC-1^2_hPlasma_Dyna_2mL_OTL_TS_VO";
            string s_sample_type = "Unknown";

            List<string> _sequemce = new List<string>();
            _sequemce.Add("Bracket Type=4");
            _sequemce.Add("File Name,Comment,Position,Inj Vol,Instrument Method,Path,Sample Type,Sample ID,Sample Name,Dil Factor,Sample Vol,Process Method,Level,Calibration File,Sample Wt,ISTD Amt,L1 Study,L2 Client,L3 Laboratory,L4 Company,L5 Phone,");

            for (int i=1; i<=5; i++)
            {
                string s_prtc = String.Format("{0}_PRTC_VO_B{1},VO {2},\"2:F,7\",4.500,D:\\OTFLumos\\Plasma-Aß\\_methods\\20170202_PRTC_PRM_VO5,D:\\OTFLumos\\_QC_PRTC,Unknown,,,1.000,0.000,,,,0.000,0.000,,,,,,",
                                              f_date,
                                              i.ToString().PadLeft(2, '0'),
                                              "07/18/2018");
                _sequemce.Add(s_prtc);
            }

            foreach (DataGridViewRow row in dgSampleList.Rows)
            {
                if (row.Cells[1].Value != null)
                {
                    string filename_blank = String.Format("{0}_RC_RS_N_N_{1}^SN{2}-{3}",
                                                    f_date,
                                                    row.Cells[1].Value,
                                                    row.Cells[0].Value.ToString().PadLeft(3, '0'),
                                                    "a");
                    string filename_a = String.Format("{0}_ADRC-{1}_Plasma_HJ5p1-IP_ISTD_SN{2}-{3}",
                                                    f_date,
                                                    row.Cells[1].Value,
                                                    row.Cells[0].Value.ToString().PadLeft(3, '0'),
                                                    "a");
                    string filename_b = String.Format("{0}_ADRC-{1}_Plasma_HJ5p1-IP_ISTD_SN{2}-{3}",
                                                    f_date,
                                                    row.Cells[1].Value,
                                                    row.Cells[0].Value.ToString().PadLeft(3, '0'),
                                                    "b");
                    
                    string blank_row1 = String.Format("{0},{1},\"{2}\",{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},",
                                                    filename_blank,
                                                    "", //Comment,
                                                    b_pos, //Position,
                                                    "4.500", //Inj Vol,
                                                    s_inst_meth, //Instrument Method,
                                                    s_path, //Path,
                                                    s_sample_type, //Sample Type,
                                                    "", //Sample ID,
                                                    "", //Sample Name,
                                                    "1.000", //Dil Factor,
                                                    "0.000", //Sample Vol,
                                                    "", //Process Method,
                                                    "", //Level,
                                                    "", //Calibration File,
                                                    "0.000", //Sample Wt,
                                                    "0.000", //ISTD Amt,
                                                    "", //L1 Study,
                                                    "", //L2 Client,
                                                    "", //L3 Laboratory,
                                                    "", //L4 Company,
                                                    ""); //L5 Phone

                    s_pos = tray.Dequeue();

                    string sample_row1 = String.Format("{0},{1},\"{2}\",{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},",
                                                    filename_a,
                                                    s_batch, //Comment,
                                                    s_pos, //Position,
                                                    "4.500", //Inj Vol,
                                                    s_inst_meth, //Instrument Method,
                                                    s_path, //Path,
                                                    s_sample_type, //Sample Type,
                                                    "", //Sample ID,
                                                    "", //Sample Name,
                                                    "1.000", //Dil Factor,
                                                    "0.000", //Sample Vol,
                                                    "", //Process Method,
                                                    "", //Level,
                                                    "", //Calibration File,
                                                    "0.000", //Sample Wt,
                                                    "0.000", //ISTD Amt,
                                                    "", //L1 Study,
                                                    "", //L2 Client,
                                                    "", //L3 Laboratory,
                                                    "", //L4 Company,
                                                    ""); //L5 Phone

                    string sample_row2 = String.Format("{0},{1},\"{2}\",{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},",
                                                    filename_b,
                                                    s_batch, //Comment,
                                                    s_pos, //Position,
                                                    "4.500", //Inj Vol,
                                                    s_inst_meth, //Instrument Method,
                                                    s_path, //Path,
                                                    s_sample_type, //Sample Type,
                                                    "", //Sample ID,
                                                    "", //Sample Name,
                                                    "1.000", //Dil Factor,
                                                    "0.000", //Sample Vol,
                                                    "", //Process Method,
                                                    "", //Level,
                                                    "", //Calibration File,
                                                    "0.000", //Sample Wt,
                                                    "0.000", //ISTD Amt,
                                                    "", //L1 Study,
                                                    "", //L2 Client,
                                                    "", //L3 Laboratory,
                                                    "", //L4 Company,
                                                    ""); //L5 Phone

                    _sequemce.Add(blank_row1);
                    _sequemce.Add(sample_row1);
                    _sequemce.Add(sample_row2);
                }
            }
            
            File.WriteAllLines(@"D:\test.csv", _sequemce, Encoding.GetEncoding(1250));
        }

        private void button5_Click(object sender, EventArgs e)
        {
            File.WriteAllText(@"D:\test5.csv", "Plasma-Aß", Encoding.GetEncoding(1250));
        }
    }
}
  