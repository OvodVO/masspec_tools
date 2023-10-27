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
using System.Xml.Linq;
using SkylineTool;
using WashU.BatemanLab.MassSpec.Tools.ProcessRawData;
using WashU.BatemanLab.MassSpec.Tools.Analysis;
using WashU.BatemanLab.Common;
using MSFileReader = MSFileReaderLib;

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

            //MessageBox.Show(DocumentLocation.ToString());
            Invoke(new Action(() => MessageBox.Show(ReplicateName, "GetReplicateName()")));
            Invoke(new Action(() => MessageBox.Show(DocumentLocationName, "DocumentLocationName()")));
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
            _msdatafile.GetChromatograms(GetProteinsFromSkyline(), Toleranse, Toleranse);
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

            await ReadAndAnalyzeSetAsync(GetMSFilesFromSkyline().Select(f => f.Split('*')[1]).ToArray());
        }

        private async void mnuNASelectMSRunsToAnalyze_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDlg = new OpenFileDialog();
            openDlg.Filter = "Thermo(*.raw,|*.raw;)";
            openDlg.Multiselect = true;
            if (openDlg.ShowDialog() == DialogResult.OK)
            {
                await ReadAndAnalyzeSetAsync(openDlg.FileNames);
            }
        }

        private void btnMakeLinkerToSkyline_Click(object sender, EventArgs e)
        {
            CreateSkylineLinker();
        }

        private void tabMainForm_Enter(object sender, EventArgs e)
        {
            ActivateNoiseAnalysisTab();
            ActivateLCTracersTab();
        }

        private void btnPasteFromExcel_Click(object sender, EventArgs e)
        {
            string s = Clipboard.GetText();

            string[] lines = s.Replace("\n", "").Split('\r');

            string[] columns = lines[0].Split('\t');
            foreach (string column in columns)
            {
                dgSampleList.Columns.Add(column, column);
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
            string f_date = "20180818";
            string s_pos = ""; string b_pos = "1:F,8";
            string s_batch = "Batch_09 | ";

            Queue<string> tray = new Queue<string>(new string[] { "1:A,1", "1:A,2", "1:A,3", "1:A,4", "1:A,5", "1:A,6", "1:A,7", "1:A,8",
                                                                  "1:B,1", "1:B,2", "1:B,3", "1:B,4", "1:B,5", "1:B,6", "1:B,7", "1:B,8",
                                                                  "1:C,1", "1:C,2", "1:C,3", "1:C,4", "1:C,5", "1:C,6", "1:C,7", "1:C,8",
                                                                  "1:D,1", "1:D,2", "1:D,3", "1:D,4", "1:D,5", "1:D,6", "1:D,7", "1:D,8",
                                                                  "1:E,1", "1:E,2", "1:E,3", "1:E,4", "1:E,5", "1:E,6", "1:E,7", "1:E,8",
                                                                  "1:F,1", "1:F,2", "1:F,3", "1:F,4", "1:F,5", "1:F,6", "1:F,7", "1:F,8",});

            string s_inst_meth = "D:\\OTFLumos\\Plasma-Aß\\_methods\\20180812_Ab-noC13_HSS-T3-75x100_AVG-mass_300ulpmin_VO5";
            string s_path = @"D:\OTFLumos\Plasma-Aß\20180818_ADRC-9^_hPlasma_Dyna_2mL_OTL_TS_VO";
            string s_sample_type = "Unknown";

            List<string> _sequemce = new List<string>();
            _sequemce.Add("Bracket Type=4");
            _sequemce.Add("File Name,Comment,Position,Inj Vol,Instrument Method,Path,Sample Type,Sample ID,Sample Name,Dil Factor,Sample Vol,Process Method,Level,Calibration File,Sample Wt,ISTD Amt,L1 Study,L2 Client,L3 Laboratory,L4 Company,L5 Phone,");

            for (int i = 1; i <= 5; i++)
            {
                string s_prtc = String.Format("{0}_PRTC_VO_B{1},VO {2},\"2:F,7\",4.500,D:\\OTFLumos\\Plasma-Aß\\_methods\\20170202_PRTC_PRM_VO5,D:\\OTFLumos\\_QC_PRTC,Unknown,,,1.000,0.000,,,,0.000,0.000,,,,,,",
                                              f_date,
                                              i.ToString().PadLeft(2, '0'),
                                              "08/12/2018");
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

                    var s_comm = s_batch + "SN" + row.Cells[0].Value.ToString().PadLeft(3, '0');

                    string sample_row1 = String.Format("{0},{1},\"{2}\",{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},",
                                                    filename_a,
                                                    s_comm, //Comment,
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
                                                    s_comm, //Comment,
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

            SaveFileDialog _saveDlg = new SaveFileDialog();
            if (_saveDlg.ShowDialog() == DialogResult.OK)
            {
                File.WriteAllLines(_saveDlg.FileName, _sequemce, Encoding.GetEncoding(1250));
            }


        }

        private void button5_Click(object sender, EventArgs e)
        {
            File.WriteAllText(@"D:\test5.csv", "Plasma-Aß", Encoding.GetEncoding(1250));
        }

        private void btSelectFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDlg = new FolderBrowserDialog();

            //folderBrowserDlg.SelectedPath = @"d:\_TEMP\mzXMLTEST\";

            if (folderBrowserDlg.ShowDialog() == DialogResult.OK)
            {
                lbWorkingFolder.Text = folderBrowserDlg.SelectedPath;

                string[] mzXMLfiles = Directory.GetFiles(folderBrowserDlg.SelectedPath, "*.mzXML");

                cblSelectedFiles.Items.AddRange(mzXMLfiles);

                CheckUnprocessedmzXML();


            }




        }

        private void button9_Click(object sender, EventArgs e)
        {
            //string currentPath = @"o:\Plasma-Aß\20190722_MAPT-9_hPlasma_Dyna_0^5mL_OTL_NO-KF_VO\";
            //string currentPath = @"o:\Plasma-Aß\20190724_MAPT-10_hPlasma_Dyna_0^5mL_OTL_NO-KF_VO\";
            string currentPath = @"o:\Plasma-Aß\20190918_MAYO-1r_hPlasma_Dyna_0^5mL_OTL_NO-KF_PL\";

            string[] mzXMLfiles = Directory.GetFiles(currentPath, "*.mzXML");

            cblSelectedFiles.Items.AddRange(mzXMLfiles);

            CheckUnprocessedmzXML();
        }

        private void InitlVSubsForSkyline()
        {
            lVSubsForSkyline.View = View.Details;
            lVSubsForSkyline.GridLines = true;
            lVSubsForSkyline.FullRowSelect = true;

            lVSubsForSkyline.CheckBoxes = true;

            lVSubsForSkyline.Columns.Add("", 20);
            lVSubsForSkyline.Columns.Add("What", 50);
            lVSubsForSkyline.Columns.Add("Find", 300); lVSubsForSkyline.Columns.Add("Replace", 300);
        }

        private void GetMSLevelSubstitutionRow()
        {
            string[] replacement = new string[4];
            ListViewItem itm;

            replacement[0] = "true";
            replacement[1] = "MS Level";
            replacement[2] = "msLevel=\"3\"";
            replacement[3] = "msLevel=\"2\"";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);
        }
        private void GetAbetaSubstitutionList()
        {
            string[] replacement = new string[4];
            ListViewItem itm;

            replacement[0] = "true";
            replacement[1] = "Ab40 N14";
            replacement[2] = ">891</precursorMz>";
            replacement[3] = ">607.7778</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "Ab40 N15";
            replacement[2] = ">901.5</precursorMz>";
            replacement[3] = ">614.7317</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "Ab42 N14";
            replacement[2] = ">1096.5</precursorMz>";
            replacement[3] = ">699.8963</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "Ab42 N15";
            replacement[2] = ">1109.5</precursorMz>";
            replacement[3] = ">707.8436</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "Ab38 N14";
            replacement[2] = ">768.5</precursorMz>";
            replacement[3] = ">508.6460</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "Ab38 N15";
            replacement[2] = ">777.5</precursorMz>";
            replacement[3] = ">514.6065</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "AbMD N14";
            replacement[2] = ">1028.5</precursorMz>";
            replacement[3] = ">663.7446</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "AbMD N15";
            replacement[2] = ">1038.5</precursorMz>";
            replacement[3] = ">670.6985</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);


            foreach (ListViewItem item in lVSubsForSkyline.Items)
            {
                item.Checked = true;
            }


        }



        private void GetTauSubstitutionList()
        {
            string[] replacement = new string[4];
            ListViewItem itm;

            replacement[0] = "true";
            replacement[1] = "p0N pT111 14N";
            replacement[2] = ">967</precursorMz>";
            replacement[3] = ">835.3694</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "pT153 14N";
            replacement[2] = ">524</precursorMz>";
            replacement[3] = ">319.1571</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "pT231 14N";
            replacement[2] = ">660</precursorMz>";
            replacement[3] = ">523.7915</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "pT231 AQUA";
            replacement[2] = ">660.1</precursorMz>";
            replacement[3] = ">527.7986</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "pT175 14N";
            replacement[2] = ">493.7572</precursorMz>";
            replacement[3] = ">367.2018</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "pT175 AQUA";
            replacement[2] = ">497.7643</precursorMz>";
            replacement[3] = ">369.8733</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "TPSL 14N";
            replacement[2] = ">668.3726</precursorMz>";
            replacement[3] = ">533.7982</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "TPSL AQUA";
            replacement[2] = ">678.3809</precursorMz>";
            replacement[3] = ">538.8023</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "TPSL 15N";
            replacement[2] = ">677.3459</precursorMz>";
            replacement[3] = ">540.2789</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "pT217 pS214 14N";
            replacement[2] = ">713</precursorMz>";
            replacement[3] = ">573.7814</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "pT217 AQUA";
            replacement[2] = ">758.34</precursorMz>";
            replacement[3] = ">578.7855</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "pT181 14N";
            replacement[2] = ">735.3554</precursorMz>";
            replacement[3] = ">556.6062</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);

            replacement[0] = "true";
            replacement[1] = "pT181 AQUA";
            replacement[2] = ">739.3625</precursorMz>";
            replacement[3] = ">559.2776</precursorMz>";

            itm = new ListViewItem(replacement);
            lVSubsForSkyline.Items.Add(itm);


            foreach (ListViewItem item in lVSubsForSkyline.Items)
            {
                item.Checked = true;
            }


        }
        private void tabPrepForSkyline_Enter(object sender, EventArgs e)
        {
            if (!HasPrepForSkylineTabActivated)
            {
                lbWorkingFolder.Text = "";
                InitlVSubsForSkyline();
                HasPrepForSkylineTabActivated = true;
            }
        }

        private void SetlVSubsForSkyline()
        {
            lVSubsForSkyline.Items.Clear();

            if (clbReplacementsSelection.CheckedItems.Count > 0)

            {
                GetMSLevelSubstitutionRow();

                foreach (var item in clbReplacementsSelection.CheckedItems)
                {
                    switch (item.ToString())
                    {
                        case "Abeta":
                            GetAbetaSubstitutionList();
                            break;
                        case "Tau":
                            GetTauSubstitutionList();
                            break;
                    }
                }
            }

        }


        private void btReplaceAll_Click(object sender, EventArgs e)
        {
            foreach (object drMzFile in cblSelectedFiles.CheckedItems)
            {
                string _fileName = drMzFile.ToString();
                string _Rawfilename = Path.ChangeExtension(_fileName, ".raw");
                string _mzXMLfile = File.ReadAllText(_fileName);

                foreach (ListViewItem itemToSub in lVSubsForSkyline.CheckedItems)
                {
                    _mzXMLfile = _mzXMLfile.Replace(itemToSub.SubItems[2].Text, itemToSub.SubItems[3].Text);
                }

                string _directoryNameOnly = Path.GetDirectoryName(_fileName);
                string _newDirectoryName = Path.Combine(_directoryNameOnly, "Skyline");

                if (!Directory.Exists(_newDirectoryName)) Directory.CreateDirectory(_newDirectoryName);

                string _fileNameOnly = Path.GetFileName(_fileName);

                string _newFileName = _newDirectoryName + Path.DirectorySeparatorChar + _fileNameOnly;

                File.WriteAllText(_newFileName, _mzXMLfile);

                XDocument _mzXMLdoc;
                try
                {
                    _mzXMLdoc = XDocument.Load(_newFileName);
                }
                catch (Exception _XMLExeption)
                {
                    MessageBox.Show(_XMLExeption.Message);
                    return;
                }


                var indexOffsetElement = _mzXMLdoc.Root.Descendants().SingleOrDefault(p => p.Name.LocalName == "indexOffset");
                var indexElement = _mzXMLdoc.Root.Descendants().SingleOrDefault(p => p.Name.LocalName == "index");


                indexOffsetElement.Remove();
                indexElement.Remove();

                _mzXMLdoc.Save(_newFileName);

                File.SetCreationTime(_newFileName, GetRawCreationDate(_Rawfilename));
                File.SetLastWriteTime(_newFileName, GetRawCreationDate(_Rawfilename));
                //File.SetLastWriteTime(_newFileName, GetRawCreationDate(_Rawfilename));

            }

            MessageBox.Show("Completed");

        }

        private void CheckAllmzXML()
        {
            for (int i = 0; i < cblSelectedFiles.Items.Count; i++)
            {
                cblSelectedFiles.SetItemChecked(i, true);
            }
        }
        private void CheckUnprocessedmzXML()
        {
            for (int i = 0; i < cblSelectedFiles.Items.Count; i++)
            {
                string _fileName = cblSelectedFiles.Items[i].ToString();

                string _directoryNameOnly = Path.GetDirectoryName(_fileName);

                string _newDirectoryName = Path.Combine(_directoryNameOnly, "Skyline");

                string _fileNameOnly = Path.GetFileName(_fileName);

                string _newFileName = _newDirectoryName + Path.DirectorySeparatorChar + _fileNameOnly;

                if (File.Exists(_newFileName))
                { cblSelectedFiles.SetItemChecked(i, false); }
                else { cblSelectedFiles.SetItemChecked(i, true); }
            }
        }
        private void btCheckUnprocessed_Click(object sender, EventArgs e)
        {
            CheckUnprocessedmzXML();
        }

        private void btCheckALL_Click(object sender, EventArgs e)
        {
            CheckAllmzXML();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            foreach (object drMzFile in cblSelectedFiles.CheckedItems)
            {
                string _fileName = drMzFile.ToString();

                string _fileNameRAW = _fileName.Replace(".mxXML", ".raw");

                var msrun = new MsDataFileImplExtAgg(_fileNameRAW);

                // msrun.MsDataFile.RunStartTime

                MessageBox.Show(msrun.MsDataFile.IsThermoFile.ToString(), "Is Thermo File");
                MessageBox.Show(msrun.MsDataFile.IsWatersFile.ToString(), "Is Waters File");

                MessageBox.Show(msrun.MsDataFile.RunStartTime.HasValue.ToString(), "RunStartTime.HasValue");

                //MessageBox.Show(msrun.MsDataFile. .HasValue.ToString(), "RunStartTime.HasValue");

                //string _mzXMLfile = File.ReadAllText(_fileName);

                //foreach (ListViewItem itemToSub in lVSubsForSkyline.CheckedItems)
                //{
                //    _mzXMLfile = _mzXMLfile.Replace(itemToSub.SubItems[2].Text, itemToSub.SubItems[3].Text);
                //}

                //string _directoryNameOnly = Path.GetDirectoryName(_fileName);
                //string _newDirectoryName = Path.Combine(_directoryNameOnly, "Skyline");

                //if (!Directory.Exists(_newDirectoryName)) Directory.CreateDirectory(_newDirectoryName);

                //string _fileNameOnly = Path.GetFileName(_fileName);

                //string _newFileName = _newDirectoryName + Path.DirectorySeparatorChar + _fileNameOnly;

                //File.WriteAllText(_newFileName, _mzXMLfile);

            }

            MessageBox.Show("Completed");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            OpenFileDialog _openDlg = new OpenFileDialog();
            _openDlg.Filter = "RAW|*.RAW";
            _openDlg.Multiselect = true;

            if (_openDlg.ShowDialog() == DialogResult.OK)
            {
                DateTime exported = GetRawCreationDate(_openDlg.FileName);
            }
        }


        public static DateTime GetRawCreationDate(string RawFileName)
        {
            MSFileReader.IXRawfile _rawfile = new MSFileReader.MSFileReader_XRawfile();

            _rawfile.Open(RawFileName);
            _rawfile.SetCurrentController(0, 1);

            //  string AquDate = null;
            //_rawfile.GetInstSerialNumber(ref AquDate);

            // _rawfile.GetSeqRowComment(ref AquDate);
            //  MessageBox.Show(AquDate, "AquDate");

            DateTime pCreationDate = new DateTime();
            _rawfile.GetCreationDate(ref pCreationDate);

            /* string AcquisitionFileName = null;
             _rawfile.GetAcquisitionFileName(ref AcquisitionFileName); MessageBox.Show(AcquisitionFileName, "AcquisitionFileName()");*/

            return pCreationDate;
        }

        private void clbReplacementsSelection_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetlVSubsForSkyline();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void saveAnalysisToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //_analysisResults.SaveToXML(@"D:\test.xml");
            _analysisResults.SaveToBin(@"D:\test.tra");
        }

        private void openAnalysisToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog _openDlg = new OpenFileDialog();
            _openDlg.Filter = "TRA/XML |*.tra; *.xml";

            _openDlg.Multiselect = false;

            if (_openDlg.ShowDialog() == DialogResult.OK)
            {
                switch (Path.GetExtension(_openDlg.FileName))
                {
                    case ".tra":
                        _analysisResults = AnalysisResults.OpenFromBinFile(_openDlg.FileName);
                        break;
                    case ".xml":
                        _analysisResults = AnalysisResults.OpenFromXMLFile(_openDlg.FileName);
                        break;
                    default: MessageBox.Show("Unknown extention" + Path.GetExtension(_openDlg.FileName)); break;
                }

                BuildNoiseAnalysisPlots();
            }


            //_analysisResults = AnalysisResults.OpenFromXMLFile(@"D:\test.xml");
            //_analysisResults = AnalysisResults.OpenFromBinFile(@"D:\test.tra");
            //MessageBox.Show("Opened");

        }

        private void btnConvertLCTracerFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog _openDlg = new OpenFileDialog();
            _openDlg.Filter = "LC Tracers (*.csv) |*.csv";

            _openDlg.Multiselect = false;

            if (_openDlg.ShowDialog() == DialogResult.OK)
            {
                int reduction_factor = 2000;
                char _delimiter = ',';
                string[] OriginalLines = File.ReadAllLines(_openDlg.FileName);
                List<string> ModifiedLines = new List<string>();


                DateTime StartTime = DateTime.Parse(OriginalLines[0].Split(_delimiter)[0].Split('=')[1]);
                DateTime EndTime = StartTime;

                string[] Fields = OriginalLines[1].Split(_delimiter);

                ModifiedLines.Add(OriginalLines[1]);

                for (int i = 2; i < OriginalLines.Length; i++)
                {
                    string[] _lcTracers = OriginalLines[i].Split(_delimiter);

                    int MStime = Convert.ToInt32(_lcTracers[0]);

                    if (MStime % reduction_factor == 0)
                    {
                        _lcTracers[0] = StartTime.AddMilliseconds(MStime).ToString();

                        EndTime = StartTime.AddMilliseconds(MStime);

                        string ModifiedLine = EndTime.ToString();

                        for (byte j = 1; j < Fields.Length; j++)
                        {
                            ModifiedLine += "," + _lcTracers[j];
                        }

                        ModifiedLines.Add(ModifiedLine);

                    }
                }



                string _targetPath = Path.GetDirectoryName(_openDlg.FileName);

                if (cbUseTargetDir.Checked && lbTargetDir.Text != "")
                {
                    _targetPath = lbTargetDir.Text;
                }

                string ModifiedFileName = _targetPath + Path.DirectorySeparatorChar
                    + "Lumos_LC_Tracers_" + StartTime.ToString("yyyy-MM-dd HH^mm") + "--" + EndTime.ToString("yyyy-MM-dd HH^mm") + ".csv";

                File.WriteAllLines(ModifiedFileName, ModifiedLines.ToArray());

                MessageBox.Show("Done");

            }

        }

        private void btnSelectTargetDir_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog _folderBrowserDlg = new FolderBrowserDialog();

            // _folderBrowserDlg.RootFolder = Environment.SpecialFolder. @"r:\_Projects\Aß\Plasma\Clinical Studies\BioFINDER\LC Data source BF-1";

            _folderBrowserDlg.SelectedPath = @"r:\_Projects\Aß\Plasma\Clinical Studies\BioFINDER\LC Data source BF-1";

            if (_folderBrowserDlg.ShowDialog() == DialogResult.OK)
            {

                // MessageBox.Show("User click OK");
                lbTargetDir.Text = _folderBrowserDlg.SelectedPath;
            }
        }

        private void btnSelectTableauFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog _openDlg = new OpenFileDialog();
            _openDlg.Filter = "Tableau workbook |*.twb; *.twbx";

            _openDlg.Multiselect = false;

            if (_openDlg.ShowDialog() == DialogResult.OK)
            {
                tbTableuFileName.Text = _openDlg.FileName;

                try
                {
                    _twbXdoc = XDocument.Load(tbTableuFileName.Text);
                }
                catch (Exception _XDocExeption)
                {
                    MessageBox.Show(_XDocExeption.Message);
                    return;
                }

                _workBook = _twbXdoc.Root;

                _dataSources = _workBook.Element("datasources");


                _dsParameters = _dataSources.Elements("datasource").Where(ds => ds.Attribute("name").Value == "Parameters").Single();

                _dsMS_DataUnion = _dataSources.Elements("datasource").Where(ds => ds.Attribute("caption")?.Value.ToString() == "DS: MS Data union with Sample Info").Single();

            }
        }

        private void btnSelectPrecursorNotesCSV_Click(object sender, EventArgs e)
        {
            OpenFileDialog _openDlg = new OpenFileDialog();
            _openDlg.Filter = "Skyline Precursor Notes Reportr |*.csv";

            _openDlg.Multiselect = false;

            if (_openDlg.ShowDialog() == DialogResult.OK)
            {
                tbSelectPrecursorNotesCSV.Text = _openDlg.FileName;

            }
        }

        private void btnPasteFromClipboard_Click(object sender, EventArgs e)
        {
            dgCalcFieldToMake.Rows.Clear();
            dgCalcFieldToMake.Columns.Clear();

            string s = Clipboard.GetText();

            string[] lines = s.Replace("\n", "").Split('\r');

            lines = lines.Take(lines.Count() - 1).ToArray();


            string[] columns = lines[0].Split('\t');
            foreach (string column in columns)
            {
                dgCalcFieldToMake.Columns.Add(column, column);
            }


            dgCalcFieldToMake.Rows.Add(lines.Length - 2);
            string[] fields;
            int row = 0;
            int col = 0;

            foreach (string item in lines.Skip(1))
            {

                fields = item.Split('\t');
                foreach (string f in fields)
                {

                    dgCalcFieldToMake[col, row].Value = f;
                    col++;
                }
                row++;
                col = 0;
            }

            dgCalcFieldToMake.AutoResizeColumns();
            //MessageBox.Show(dgCalcFieldToMake.RowCount.ToString());
        }

        XDocument _twbXdoc;
        XElement _workBook;
        XElement _dataSources;
        XElement _dsParameters;
        XElement _dsMS_DataUnion;


        private void btnModifyTableau_Click(object sender, EventArgs e)
        {


            // MessageBox.Show(GetCalcFieldNameByCaption(_dsMS_DataUnion, "Technical replicate ID"), "Technical replicate ID");

            // MessageBox.Show(GetCalcFieldNameByCaption(_dsMS_DataUnion, "Technical replicate ID5"), "Technical replicate ID5");

            // MessageBox.Show(GetParameterNameByCaption(_dsParameters, "p 13C15N 2N4R Amount, ng"), "p 13C15N 2N4R Amount, ng");

            // MessageBox.Show(GetParameterNameByCaption(_dsParameters, "p 13C15N 2N4R Amount, ng5"), "p 13C15N 2N4R Amount, ng5");



            int i = 10; int j = 10;

            foreach (DataGridViewRow row in dgCalcFieldToMake.Rows)
            {
                String[] arrPrecursorAttrs = row.Cells[0].Value.ToString().Split(' ');

                bool ifPreclongName = false;

                if (arrPrecursorAttrs.Length > 2) { ifPreclongName = true; }

                String precursorName = String.Format("{1} {2} {0}", arrPrecursorAttrs[0], arrPrecursorAttrs[1].Trim(' '), ifPreclongName ? arrPrecursorAttrs[2].Trim(' ') : null);

                String _caption_Param = String.Format("p Skyline Inx (prec): {0}", row.Cells[0].Value);
                String _caption_CalcF = String.Format("Area: {1} {2}({0}, ToUse)", arrPrecursorAttrs[0], arrPrecursorAttrs[1].Trim(' '), ifPreclongName ? arrPrecursorAttrs[2].Trim(' ') : null);

                // MessageBox.Show(_caption_CalcF);


                XElement _foundParameter = _dsParameters.Descendants("column").Where(p => p.Attribute("caption") != null).SingleOrDefault(p => (string)p.Attribute("caption").Value == _caption_Param);

                String _curParameterName = String.Format("[Parameter {0}]", i);

                string paramvalue = row.Cells[1].Value.ToString() != "" ? row.Cells[1].Value.ToString() : "0";

                if (_foundParameter == null)
                {
                    //MessageBox.Show("Not Found " + _caption + " parameter");


                    XElement _newParameter = new XElement("column",
                                                          new XAttribute("caption", _caption_Param),
                                                          new XAttribute("datatype", "integer"),
                                                          new XAttribute("name", _curParameterName),
                                                          new XAttribute("param-domain-type", "range"),
                                                          new XAttribute("role", "measure"),
                                                          new XAttribute("type", "quantitative"),
                                                          new XAttribute("value", paramvalue),
                                                          new XElement("calculation",
                                                                          new XAttribute("class", "tableau"),
                                                                          new XAttribute("formula", "1")),
                                                          new XElement("range",
                                                                          new XAttribute("granularity", "1"),
                                                                          new XAttribute("max", "3"),
                                                                          new XAttribute("min", "0"))
                                                         );

                    _dsParameters.Descendants("column").Last().AddAfterSelf(_newParameter);

                    i++;
                }
                else
                {
                    _curParameterName = _foundParameter.Attribute("name").Value.ToString();

                    _foundParameter.Element("range").Attribute("min").Value = "0";
                    _foundParameter.Attribute("value").Value = paramvalue;
                    // MessageBox.Show("Found " + _caption + " parameter");
                    // MessageBox.Show(_foundParameter.ToString());
                }


                
                XElement _foundCalcField = _dsMS_DataUnion.Elements("column").Where(p => p.Attribute("caption") != null).SingleOrDefault(p => (string)p.Attribute("caption").Value == _caption_CalcF);

                

                string _techreplicate_fieldname = GetCalcFieldNameByCaption(_dsMS_DataUnion, "Technical replicate ID");
                string _skylineInx_fieldname = GetCalcFieldNameByCaption(_dsMS_DataUnion, "Skyline File Inx");
                string _ifPassedQCprec_fieldname = GetCalcFieldNameByCaption(_dsMS_DataUnion, "if Passed QC (prec)");

                string _showfailed_paramname = GetParameterNameByCaption(_dsParameters, "p If Show FAILED QC");

                if (_foundCalcField == null)
                {
                    XElement _newCalcField = new XElement("column",
                                                          new XAttribute("caption", _caption_CalcF),
                                                          new XAttribute("datatype", "real"),
                                                          new XAttribute("name", String.Format("[Calculation_{0}]", j)),
                                                          new XAttribute("role", "measure"),
                                                          new XAttribute("type", "quantitative"),
                                                          new XElement("calculation",
                                                                          new XAttribute("class", "tableau"),
                                                                          new XAttribute("formula", @String.Format(@tbFormulaString.Text,
                                                                                                                            row.Cells[0].Value,
                                                                                                                            _curParameterName,
                                                                                                                            _techreplicate_fieldname,
                                                                                                                            _skylineInx_fieldname,
                                                                                                                            _ifPassedQCprec_fieldname,
                                                                                                                            _showfailed_paramname
                                                                                                                                         )))
                                                         );

                    _dsMS_DataUnion.Elements("column").Last().AddAfterSelf(_newCalcField);

                    j++;

                }
                else
                {
                    // MessageBox.Show(_caption_CalcF + " already exists");
                    _foundCalcField.Element("calculation").Attribute("formula").Value = @String.Format(@tbFormulaString.Text,
                                                                                                                row.Cells[0].Value,
                                                                                                                _curParameterName,
                                                                                                                _techreplicate_fieldname,
                                                                                                                _skylineInx_fieldname,
                                                                                                                _ifPassedQCprec_fieldname,
                                                                                                                _showfailed_paramname);
                }

            }

            string _NewFileName = tbTableuFileName.Text.Replace(Path.GetFileNameWithoutExtension(@tbTableuFileName.Text), Path.GetFileNameWithoutExtension(@tbTableuFileName.Text) + "_new");

            _twbXdoc.Save(@_NewFileName);

        }

        private void btnPasteFromClipboardToFormulaBox_Click(object sender, EventArgs e)
        {
            tbFormulaString.Clear();
            tbFormulaString.Text = Clipboard.GetText();

        }

        private void btnPasteFromClipboardToGrid2_Click(object sender, EventArgs e)
        {
            dgCalcFieldToMake2.Rows.Clear();
            dgCalcFieldToMake2.Columns.Clear();

            string s = Clipboard.GetText();

            string[] lines = s.Replace("\n", "").Split('\r');

            lines = lines.Take(lines.Count() - 1).ToArray();


            string[] columns = lines[0].Split('\t');
            foreach (string column in columns)
            {
                dgCalcFieldToMake2.Columns.Add(column, column);
            }


            dgCalcFieldToMake2.Rows.Add(lines.Length - 2);
            string[] fields;
            int row = 0;
            int col = 0;

            foreach (string item in lines.Skip(1))
            {

                fields = item.Split('\t');
                foreach (string f in fields)
                {

                    dgCalcFieldToMake2[col, row].Value = f;
                    col++;
                }
                row++;
                col = 0;
            }

            dgCalcFieldToMake2.AutoResizeColumns();
            //MessageBox.Show(dgCalcFieldToMake.RowCount.ToString());

        }

        private void btnPasteFromClipboardToFormulaBox2_Click(object sender, EventArgs e)
        {
            tbFormulaString2.Clear();
            tbFormulaString2.Text = Clipboard.GetText();
        }

        private void btnModifyTableauLevels_Click(object sender, EventArgs e)
        {

            int i = 10; int j = 10;

            foreach (DataGridViewRow row in dgCalcFieldToMake2.Rows)
            {
                String[] arrPrecursor1Attrs = row.Cells[5].Value.ToString().Split(' ');
                String[] arrPrecursor2Attrs = row.Cells[10].Value.ToString().Split(' ');

                bool ifPrec1longName = false;
                bool ifPrec2longName = false;

                if (arrPrecursor1Attrs.Length > 2) { ifPrec1longName = true; }
                if (arrPrecursor2Attrs.Length > 2) { ifPrec2longName = true; }


                String _caption_Precursor1 = String.Format("Area: {1} {2}({0}, ToUse)", arrPrecursor1Attrs[0], arrPrecursor1Attrs[1].Trim(' '), ifPrec1longName ? arrPrecursor1Attrs[2].Trim(' ') : null);
                String _caption_Precursor2 = String.Format("Area: {1} {2}({0}, ToUse)", arrPrecursor2Attrs[0], arrPrecursor2Attrs[1].Trim(' '), ifPrec2longName ? arrPrecursor2Attrs[2].Trim(' ') : null);

                string precursor1_fieldname = GetCalcFieldNameByCaption(_dsMS_DataUnion, _caption_Precursor1);

                string precursor2_fieldname = GetCalcFieldNameByCaption(_dsMS_DataUnion, _caption_Precursor2);


                string _analyte = row.Cells[1].Value.ToString().Replace(" level", "");
                String _caption_CalcField_AreaRatio = String.Format("Area Ratio: {0} (N14/{1})", _analyte, row.Cells[0].Value);
                String _caption_CalcField_Level = String.Format("Level: {0} (ng, by {1})", _analyte, row.Cells[0].Value);
                String _caption_CalcField_Conc = String.Format("Conc.: {0} (ng/mL, by {1})", _analyte, row.Cells[0].Value);

                if (precursor1_fieldname == "Not Found" || precursor2_fieldname == "Not Found")
                {
                    MessageBox.Show(String.Format("Either {0} or {1} not exists. Not possible to create formula for {2}.",
                                                   _caption_Precursor1,
                                                   _caption_Precursor2,
                                                   _caption_CalcField_AreaRatio), "Warning");
                    continue;
                }

                XElement _foundCalcField_AreaRatio = _dsMS_DataUnion.Elements("column").Where(p => p.Attribute("caption") != null).SingleOrDefault(p => (string)p.Attribute("caption").Value == _caption_CalcField_AreaRatio);
                if (_foundCalcField_AreaRatio == null)
                {


                    XElement _newCalcField_AreaRatio = new XElement("column",
                                                          new XAttribute("caption", _caption_CalcField_AreaRatio),
                                                          new XAttribute("datatype", "real"),
                                                          new XAttribute("name", String.Format("[Calculation_AR{0}]", j)),
                                                          new XAttribute("role", "measure"),
                                                          new XAttribute("type", "quantitative"),
                                                          new XElement("calculation",
                                                                          new XAttribute("class", "tableau"),
                                                                          new XAttribute("formula", @String.Format("{0} / {1}", precursor1_fieldname, precursor2_fieldname)))
                                                         );

                    _dsMS_DataUnion.Elements("column").Last().AddAfterSelf(_newCalcField_AreaRatio);

                    j++;

                }
                else
                {
                    _foundCalcField_AreaRatio.Element("calculation").Attribute("formula").Value = @String.Format("{0} / {1}", precursor1_fieldname, precursor2_fieldname);
                }


                string arearatio_fieldname = GetCalcFieldNameByCaption(_dsMS_DataUnion, _caption_CalcField_AreaRatio);
                string isamount_fieldname = GetCalcFieldNameByCaption(_dsMS_DataUnion, row.Cells[13].Value.ToString());

                if (isamount_fieldname == "Not Found")
                {
                    XElement _newCalcField_ISamount = new XElement("column",
                                                         new XAttribute("caption", row.Cells[13].Value.ToString()),
                                                         new XAttribute("datatype", "real"),
                                                         new XAttribute("name", String.Format("[Calculation_IS{0}]", j)),
                                                         new XAttribute("role", "measure"),
                                                         new XAttribute("type", "quantitative"),
                                                         new XElement("calculation",
                                                                         new XAttribute("class", "tableau"),
                                                                         new XAttribute("formula", @String.Format("1.000")))
                                                        );

                    isamount_fieldname = String.Format("[Calculation_IS{0}]", j);

                    _dsMS_DataUnion.Elements("column").Last().AddAfterSelf(_newCalcField_ISamount);

                }

                if (arearatio_fieldname == "Not Found")
                {
                    MessageBox.Show(String.Format("Either {0} or {1} not exists. Not possible to create formula for {2}.",
                                                   _caption_CalcField_AreaRatio,
                                                   row.Cells[13].Value.ToString(),
                                                   _caption_CalcField_Level), "Warning");
                    continue;
                }

                XElement _foundCalcField_Level = _dsMS_DataUnion.Elements("column").Where(p => p.Attribute("caption") != null).SingleOrDefault(p => (string)p.Attribute("caption").Value == _caption_CalcField_Level);
                if (_foundCalcField_Level == null)
                {


                    XElement _newCalcField_Level = new XElement("column",
                                                          new XAttribute("caption", _caption_CalcField_Level),
                                                          new XAttribute("datatype", "real"),
                                                          new XAttribute("name", String.Format("[Calculation_LE{0}]", j)),
                                                          new XAttribute("role", "measure"),
                                                          new XAttribute("type", "quantitative"),
                                                          new XElement("calculation",
                                                                          new XAttribute("class", "tableau"),
                                                                          new XAttribute("formula", @String.Format("{0} * {1}", arearatio_fieldname, isamount_fieldname)))
                                                         );

                    _dsMS_DataUnion.Elements("column").Last().AddAfterSelf(_newCalcField_Level);

                    j++;

                }
                else
                {
                    _foundCalcField_Level.Element("calculation").Attribute("formula").Value = @String.Format("{0} * {1}", arearatio_fieldname, isamount_fieldname);
                }

                string level_fieldname = GetCalcFieldNameByCaption(_dsMS_DataUnion, _caption_CalcField_Level);

                string samplevolume_fieldname = GetCalcFieldNameByCaption(_dsMS_DataUnion, "Sample Volume");

                if (samplevolume_fieldname == "Not Found")
                {
                    XElement _newCalcField_SampleVolume = new XElement("column",
                                                         new XAttribute("caption", "Sample Volume"),
                                                         new XAttribute("datatype", "real"),
                                                         new XAttribute("name", String.Format("[Calculation_VO{0}]", j)),
                                                         new XAttribute("role", "measure"),
                                                         new XAttribute("type", "quantitative"),
                                                         new XElement("calculation",
                                                                         new XAttribute("class", "tableau"),
                                                                         new XAttribute("formula", @String.Format("[Volume]")))
                                                        );

                    samplevolume_fieldname = String.Format("[Calculation_VO{0}]", j);

                    _dsMS_DataUnion.Elements("column").Last().AddAfterSelf(_newCalcField_SampleVolume);

                }

                XElement _foundCalcField_Conc = _dsMS_DataUnion.Elements("column").Where(p => p.Attribute("caption") != null).SingleOrDefault(p => (string)p.Attribute("caption").Value == _caption_CalcField_Conc);
                if (_foundCalcField_Conc == null)
                {


                    XElement _newCalcField_Conc = new XElement("column",
                                                          new XAttribute("caption", _caption_CalcField_Conc),
                                                          new XAttribute("datatype", "real"),
                                                          new XAttribute("name", String.Format("[Calculation_CO{0}]", j)),
                                                          new XAttribute("role", "measure"),
                                                          new XAttribute("type", "quantitative"),
                                                          new XElement("calculation",
                                                                          new XAttribute("class", "tableau"),
                                                                          new XAttribute("formula", @String.Format("{0} / {1}", level_fieldname, samplevolume_fieldname)))
                                                         );

                    _dsMS_DataUnion.Elements("column").Last().AddAfterSelf(_newCalcField_Conc);

                    j++;

                }
                else
                {
                    _foundCalcField_Conc.Element("calculation").Attribute("formula").Value = @String.Format("{0} / {1}", level_fieldname, samplevolume_fieldname);
                }

            }

            string _NewFileName = tbTableuFileName.Text.Replace(Path.GetFileNameWithoutExtension(@tbTableuFileName.Text), Path.GetFileNameWithoutExtension(@tbTableuFileName.Text) + "_new");

            _twbXdoc.Save(@_NewFileName);


        }

        public static string GetCalcFieldNameByCaption(XElement datasource, string caption)
        {

            XElement _foundElement = datasource.Elements("column").Where(p => p.Attribute("caption") != null).SingleOrDefault(p => (string)p.Attribute("caption").Value == caption);

            return (_foundElement != null) ? _foundElement.Attribute("name").Value.ToString() : "Not Found";

        }
        public static string GetParameterNameByCaption(XElement datasource, string caption)
        {

            XElement _foundElement = datasource.Elements("column").Where(p => p.Attribute("caption") != null).SingleOrDefault(p => (string)p.Attribute("caption").Value == caption);

            return (_foundElement != null) ? "[Parameters]." + _foundElement.Attribute("name").Value.ToString() : "Not Found";

        }

        private void btnPasteFromClipboardToGrid3_Click(object sender, EventArgs e)
        {
            dgCalcFieldToMake3.Rows.Clear();
            dgCalcFieldToMake3.Columns.Clear();

            string s = Clipboard.GetText();

            string[] lines = s.Replace("\n", "").Split('\r');

            lines = lines.Take(lines.Count() - 1).ToArray();


            string[] columns = lines[0].Split('\t');
            foreach (string column in columns)
            {
                dgCalcFieldToMake3.Columns.Add(column, column);
            }


            dgCalcFieldToMake3.Rows.Add(lines.Length - 2);
            string[] fields;
            int row = 0;
            int col = 0;

            foreach (string item in lines.Skip(1))
            {

                fields = item.Split('\t');
                foreach (string f in fields)
                {

                    dgCalcFieldToMake3[col, row].Value = f;
                    col++;
                }
                row++;
                col = 0;
            }

            dgCalcFieldToMake3.AutoResizeColumns();
            //MessageBox.Show(dgCalcFieldToMake.RowCount.ToString());
        }

        private void btnModifyTableauPTau_Click(object sender, EventArgs e)
        {

            int i = 0; int j = 10;

            string _caption_Param_pTautau_Aqua_ratio = "p ptau/tau AQUA ratio";
            string _name_Param_pTautau_Aqua_ratio = String.Format("[Parameter G{0}]", 1);
            string paramvalue = "0.1";

            XElement _foundParameter_pTautau_Aqua_ratio = _dsParameters.Descendants("column").Where(p => p.Attribute("caption") != null).SingleOrDefault(p => (string)p.Attribute("caption").Value == _caption_Param_pTautau_Aqua_ratio);

            if (_foundParameter_pTautau_Aqua_ratio == null)
            {

                XElement _newParameter_pTautau_Aqua_ratio = new XElement("column",
                                                      new XAttribute("caption", _caption_Param_pTautau_Aqua_ratio),
                                                      new XAttribute("datatype", "integer"),
                                                      new XAttribute("name", _name_Param_pTautau_Aqua_ratio),
                                                      new XAttribute("param-domain-type", "range"),
                                                      new XAttribute("role", "measure"),
                                                      new XAttribute("type", "quantitative"),
                                                      new XAttribute("value", paramvalue),
                                                      new XElement("calculation",
                                                                      new XAttribute("class", "tableau"),
                                                                      new XAttribute("formula", "1")),
                                                      new XElement("range",
                                                                      new XAttribute("granularity", "1"),
                                                                      new XAttribute("max", "3"),
                                                                      new XAttribute("min", "0"))
                                                     );

                _dsParameters.Descendants("column").Last().AddAfterSelf(_newParameter_pTautau_Aqua_ratio);

                i++;
            }
            else
            {
                _name_Param_pTautau_Aqua_ratio = _foundParameter_pTautau_Aqua_ratio.Attribute("name").Value.ToString();

                _foundParameter_pTautau_Aqua_ratio.Element("range").Attribute("min").Value = "0";
                _foundParameter_pTautau_Aqua_ratio.Attribute("value").Value = paramvalue;
                // MessageBox.Show("Found " + _caption + " parameter");
                // MessageBox.Show(_foundParameter.ToString());
            }

            foreach (DataGridViewRow row in dgCalcFieldToMake3.Rows)
            {
                i++;
                String[] arrPrecursor1Attrs = row.Cells[5].Value.ToString().Split(' ');
                String[] arrPrecursor2Attrs = row.Cells[10].Value.ToString().Split(' ');
                String[] arrPrecursor3Attrs = row.Cells[22].Value.ToString().Split(' ');
                String[] arrPrecursor4Attrs = row.Cells[27].Value.ToString().Split(' ');


                bool ifPrec1longName = false;
                bool ifPrec2longName = false;
                bool ifPrec3longName = false;
                bool ifPrec4longName = false;

                if (arrPrecursor1Attrs.Length > 2) { ifPrec1longName = true; }
                if (arrPrecursor2Attrs.Length > 2) { ifPrec2longName = true; }
                if (arrPrecursor3Attrs.Length > 2) { ifPrec3longName = true; }
                if (arrPrecursor4Attrs.Length > 2) { ifPrec4longName = true; }


                String _caption_Precursor1 = String.Format("Area: {1} {2}({0}, ToUse)", arrPrecursor1Attrs[0], arrPrecursor1Attrs[1].Trim(' '), ifPrec1longName ? arrPrecursor1Attrs[2].Trim(' ') : null);
                String _caption_Precursor2 = String.Format("Area: {1} {2}({0}, ToUse)", arrPrecursor2Attrs[0], arrPrecursor2Attrs[1].Trim(' '), ifPrec2longName ? arrPrecursor2Attrs[2].Trim(' ') : null);
                String _caption_Precursor3 = String.Format("Area: {1} {2}({0}, ToUse)", arrPrecursor3Attrs[0], arrPrecursor3Attrs[1].Trim(' '), ifPrec3longName ? arrPrecursor3Attrs[2].Trim(' ') : null);
                String _caption_Precursor4 = String.Format("Area: {1} {2}({0}, ToUse)", arrPrecursor4Attrs[0], arrPrecursor4Attrs[1].Trim(' '), ifPrec4longName ? arrPrecursor4Attrs[2].Trim(' ') : null);

                string precursor1_fieldname = GetCalcFieldNameByCaption(_dsMS_DataUnion, _caption_Precursor1);

                string precursor2_fieldname = GetCalcFieldNameByCaption(_dsMS_DataUnion, _caption_Precursor2);

                string precursor3_fieldname = GetCalcFieldNameByCaption(_dsMS_DataUnion, _caption_Precursor3);

                string precursor4_fieldname = GetCalcFieldNameByCaption(_dsMS_DataUnion, _caption_Precursor4);

                /*
                MessageBox.Show(String.Format("i: {2}; Prec1: {0}; Prec2: {1}", _caption_Precursor1, _caption_Precursor2, i), "Caption pTau");
                MessageBox.Show(String.Format("i: {2}; Prec3: {0}; Prec4: {1}", _caption_Precursor3, _caption_Precursor4, i), "Caption pTau Norm");

                MessageBox.Show(String.Format("i: {2}; Prec1: {0}; Prec2: {1}", precursor1_fieldname, precursor2_fieldname, i), "Field pTau");
                MessageBox.Show(String.Format("i: {2}; Prec3: {0}; Prec4: {1}", precursor3_fieldname, precursor4_fieldname, i), "Field pTau Norm");
                */

                               
                                string _analyte = row.Cells[1].Value.ToString().Replace("R ", "");
                                
                                String _caption_CalcField_pTautau = String.Format("pTauTau: {0} (Area Ratio)", _analyte, row.Cells[0].Value);
                                String _caption_CalcField_pTautauNorm = String.Format("pTauTau: {0} (Norm Ratio)", _analyte, row.Cells[0].Value);


                                if (precursor1_fieldname == "Not Found" || precursor2_fieldname == "Not Found" || precursor3_fieldname == "Not Found" || precursor4_fieldname == "Not Found")
                                {
                                    MessageBox.Show(String.Format("Either {0} or {1} or {2} or {3} not exists. Not possible to create formula for {4}.",
                                                                   _caption_Precursor1,
                                                                   _caption_Precursor2,
                                                                   _caption_Precursor3,
                                                                   _caption_Precursor4,
                                                                   _analyte), "Warning");
                                    continue;
                                }

                                XElement _foundCalcField_pTautau = _dsMS_DataUnion.Elements("column").Where(p => p.Attribute("caption") != null).SingleOrDefault(p => (string)p.Attribute("caption").Value == _caption_CalcField_pTautau);
                                if (_foundCalcField_pTautau == null)
                                {


                                    XElement _newCalcField_pTautau = new XElement("column",
                                                                          new XAttribute("caption", _caption_CalcField_pTautau),
                                                                          new XAttribute("datatype", "real"),
                                                                          new XAttribute("name", String.Format("[Calculation_pT{0}]", j)),
                                                                          new XAttribute("role", "measure"),
                                                                          new XAttribute("type", "quantitative"),
                                                                          new XElement("calculation",
                                                                                          new XAttribute("class", "tableau"),
                                                                                          new XAttribute("formula", @String.Format("{0} / {1}", precursor1_fieldname, precursor2_fieldname)))
                                                                         );

                                    _dsMS_DataUnion.Elements("column").Last().AddAfterSelf(_newCalcField_pTautau);

                                    j++;

                                }
                                else
                                {
                                   _foundCalcField_pTautau.Element("calculation").Attribute("formula").Value = @String.Format("{0} / {1}", precursor1_fieldname, precursor2_fieldname);
                                }


                                string pTautau_fieldname = GetCalcFieldNameByCaption(_dsMS_DataUnion, _caption_CalcField_pTautau);


                                XElement _foundCalcField_pTautauNorm = _dsMS_DataUnion.Elements("column").Where(p => p.Attribute("caption") != null).SingleOrDefault(p => (string)p.Attribute("caption").Value == _caption_CalcField_pTautauNorm);
                                if (_foundCalcField_pTautauNorm == null)
                                {


                                    XElement _newCalcField_pTautauNorm = new XElement("column",
                                                                          new XAttribute("caption", _caption_CalcField_pTautauNorm),
                                                                          new XAttribute("datatype", "real"),
                                                                          new XAttribute("name", String.Format("[Calculation_pTN{0}]", j)),
                                                                          new XAttribute("role", "measure"),
                                                                          new XAttribute("type", "quantitative"),
                                                                          new XElement("calculation",
                                                                                          new XAttribute("class", "tableau"),
                                                                                          new XAttribute("formula", @String.Format("{0} * {1} / {2} * [Parameters].{3}", pTautau_fieldname,
                                                                                                                                                      precursor3_fieldname,
                                                                                                                                                      precursor4_fieldname,
                                                                                                                                                      _name_Param_pTautau_Aqua_ratio)))
                                                                         );

                                    _dsMS_DataUnion.Elements("column").Last().AddAfterSelf(_newCalcField_pTautauNorm);

                                    j++;

                                }
                                else
                                {
                                    _foundCalcField_pTautauNorm.Element("calculation").Attribute("formula").Value = @String.Format("{0} * {1} / {2} * [Parameters].{3}", pTautau_fieldname,
                                                                                                                                                      precursor3_fieldname,
                                                                                                                                                      precursor4_fieldname,
                                                                                                                                                      _name_Param_pTautau_Aqua_ratio);
                                }

            }


                        string _NewFileName = tbTableuFileName.Text.Replace(Path.GetFileNameWithoutExtension(@tbTableuFileName.Text), Path.GetFileNameWithoutExtension(@tbTableuFileName.Text) +"_"+ DateTime.Today.ToLongDateString() );

                        _twbXdoc.Save(@_NewFileName); 


        }
                
    }
}
  