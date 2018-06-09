﻿using System;
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
            IsConnectedToSkylineDoc = true;
        }

        MsDataFileImplExtAgg _msdatafile;


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
     
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IReport reportPeptideRatios = _toolClient.GetReport("BLR Peptide Ratios");

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

                foreach (var ratio in Ratios)
                {
                    listBox1.Items.Add(String.Format("Ratio {0}: File - {1}: Value: {2};", ratioVariant.RatioName, ratio.FileName, ratio.PeptideRatio));
                }

                if (ratioVariant.RatioName == "Aβ42/Aβ40")
                {
                    zedGraphControlTest.GraphPane.AddBar(ratioVariant.RatioName, null, Ratios.Select(r => r.PeptideRatio).ToArray(), Color.Green);

                    zedGraphControlTest.GraphPane.XAxis.Scale.TextLabels = Ratios.Select(f => (f.FileName as String).Substring(10, 10)).ToArray();
                }

            }

            zedGraphControlTest.GraphPane.XAxis.MajorTic.IsBetweenLabels = true;
            zedGraphControlTest.GraphPane.XAxis.Type = ZedGraph.AxisType.Text;
            //zedGraphControlTest.GraphPane.XAxis.Scale.FontSpec = ZedGraph.FontSpec.
            zedGraphControlTest.AxisChange();
            zedGraphControlTest.Refresh();

            //  listBox1.Items.Add(ratios.Count());

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

        private void button2_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(Environment.GetFolderPath(Environment.SpecialFolder.Programs));
            sb.Append("\\");
            //publisher name is Nightbird
            sb.Append("Bateman Lab");
            sb.Append("\\");
            //product name is TestRunningWithArgs
            sb.Append("TrackIN.appref-ms ");
            string shortcutPath = sb.ToString();

            string argsToPass = "duzhe,klassno,ha,ho";

            System.Diagnostics.Process.Start(shortcutPath, argsToPass);

            MessageBox.Show(shortcutPath);
            //string[] activationData =
            //    AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string[] activationData =
                  AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData;

            char[] myComma = { ',' };
            foreach (string arg in activationData)
            {
                //string[] myList = activationData[0].Split(myComma, StringSplitOptions.RemoveEmptyEntries);

//                foreach (string)

                MessageBox.Show(arg);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            CreateSkylineLinker();
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            List<Protein> test = GetProteinsFromSkyline();

            foreach (var testpro in test)
            {
                MessageBox.Show(String.Format("Prot - {0}", testpro.Name));
            }
            
            var proteins = from prot in test
                     from pept in prot.Peptides
                     from prec in pept.Precursors
                     from prod in prec.Products
                     select prot;

            foreach (var pr in test)
                foreach (var pe in pr.Peptides)
                    foreach (var prec in pe.Precursors)
                        foreach (var prod in prec.Products)
                            listBox1.Items.Add(String.Format("Prot - {0}; Pept - {1}; Isotop - {2}; PrecMZ - {3}; ProdMZ - {4}",
                                                             pr.Name,
                                                             pe.Name,
                                                             prec.IsotopeLabelType,
                                                             prec.PrecursorMZ,
                                                             prod
                                                             )
                                              );
        }

        private void mnuAnalizeFromSkyline_Click(object sender, EventArgs e)
        {
            PlotChromatograms(zedGraphControlTest);
        }

        private async void mnuNASelectMSRunsToAnalyze_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDlg = new OpenFileDialog();
            openDlg.Filter = "Thermo(*.raw,|*.raw;)";
            openDlg.Multiselect = true;
            var TargetedProteins = GetProteinsFromSkyline();
            if (openDlg.ShowDialog() == DialogResult.OK)
            {
               await ReadAndAnalyzeSet(openDlg.FileNames);
            }
        }
    }
}
  