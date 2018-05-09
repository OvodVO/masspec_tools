using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WashU.BatemanLab.MassSpec.Tools.ProcessRawData;
using WashU.BatemanLab.MassSpec.Tools.TargetAnalysis;

namespace TrackIN
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }
                
        MsDataFileImplExtAgg _msdatafile;
        
       
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog _openDlg = new OpenFileDialog();
            _openDlg.Filter = "Thermo(*.raw,|*.raw;)";

            double[] times, TICs, IITs, precursorMZ;
            //byte[] msLevels;

            if (_openDlg.ShowDialog() == DialogResult.OK)
            {
                _msdatafile = new MsDataFileImplExtAgg(_openDlg.FileName);

                //_msdatafile.MsDataFile.GetSpectrumsInfo(CancellationToken.None, out times, out msLevels, out TICs, out IITs);

              /*  double[] lev = new double[msLevels.Length];
                for (int i = 1; i < msLevels.Length; i++)
                    lev[i] = Convert.ToDouble(msLevels[i]); */


                //_msdatafile.MsDataFile.GetSpectrumsInfo(out times, out precursorMZ);
                _msdatafile.MsDataFile.GetSpectrumsInfo(out times, out TICs, out IITs, out precursorMZ);

                MessageBox.Show("Ready");


              

                zedGraphControl_TIC.GraphPane.AddCurve("TIC", times, TICs, Color.Red);
                zedGraphControl_IIT.GraphPane.AddCurve("IIT", times, IITs, Color.DarkBlue);
               // zedGraphControl_PrecursorMZ.GraphPane.AddCurve("Precursors", times, precursorMZ, Color.Green);
                var PrecursorMSline = zedGraphControl_PrecursorMZ.GraphPane.AddCurve("Precursors", times, precursorMZ, Color.Green, ZedGraph.SymbolType.Star );
                PrecursorMSline.Line.IsVisible = false;


                zedGraphControl_TIC.AxisChange(); zedGraphControl_IIT.AxisChange(); zedGraphControl_PrecursorMZ.AxisChange();
                zedGraphControl_TIC.Refresh(); zedGraphControl_IIT.Refresh(); zedGraphControl_PrecursorMZ.Refresh();

            }
        }



        private void button2_Click(object sender, EventArgs e)
        {            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog _openDlg = new OpenFileDialog();
            _openDlg.Filter = "Thermo(*.raw,|*.raw;)";

            double[] times; // = new double[];
            double[] IITs; // = new double[];
            double[] TICs;  //, precursorMZ;

            double[] posEICs;// = new double[];
            double[] negEICs;// = new double[];
            double[] AllMzs; 

            //byte[] msLevels;

            double Toleranse = 0.1;

            if (_openDlg.ShowDialog() == DialogResult.OK)
            {
                _msdatafile = new MsDataFileImplExtAgg(_openDlg.FileName);

                 MessageBox.Show("Starting GetmsDataSpectrumList()");

                //_msdatafile.MsDataFile.GetMsDataSpectrums();

                _msdatafile.GetMsDataSpectrum();
                
                MessageBox.Show("Starting chrom extract");
                
                var proteins = TargetAnalysis.GetDefaultProteins();

                var Targets = from protein in TargetAnalysis.GetDefaultProteins()
                              from peptide in protein.Peptides
                              from precursor in peptide.Precursors
                              select new { ProteinName = protein.Name,
                                  PeptideName = peptide.Name,
                                  PrecursorIsoform = precursor.IsotopeLabelType, PrecursorMZ = precursor.PrecursorMZ, Products = precursor.Products };
                              
                foreach (var target in Targets)
                {

                    var Chrom = from mzspectrum in _msdatafile.MsDataFile.MsDataSpectrums
                                where ProcessRawDataTools.InMZTolerance(mzspectrum.PrecursorMZ.GetValueOrDefault(0), target.PrecursorMZ, 0.1) == true
                                select new
                                {
                                    mzspectrum.RetentionTime,
                                    mzspectrum.IonIT,
                                    mzspectrum.TIC,
                                    AllMz = mzspectrum.Intensities.Sum(),
                                    Pos = ProcessRawDataTools.AggIonCounts(mzspectrum.Mzs, mzspectrum.Intensities, target.Products, Toleranse)[0],
                                    Neg = ProcessRawDataTools.AggIonCounts(mzspectrum.Mzs, mzspectrum.Intensities, target.Products, Toleranse)[1]
                                };

                    
                    times = Chrom.Select(x => x.RetentionTime.GetValueOrDefault(0)).ToArray<double>();
                    IITs  = Chrom.Select(x => x.IonIT.GetValueOrDefault(0)).ToArray<double>();
                    TICs  = Chrom.Select(x => x.TIC.GetValueOrDefault(0)).ToArray<double>();
                    AllMzs  = Chrom.Select(x => x.AllMz).ToArray<double>();
                    posEICs = Chrom.Select(x => x.Pos).ToArray<double>(); negEICs = Chrom.Select(x => x.Neg).ToArray<double>();

                    
                    var Testline = zedGraphControlTest.GraphPane.AddCurve(String.Concat("XIC pos ", target.PrecursorIsoform).ToString(), times, posEICs, Color.Red);
                                       
                    if (target.PrecursorIsoform.Contains("N15")) Testline.Line.Color = Color.Blue;
                    else Testline.Line.Color = Color.Green;

                   Testline.Symbol.IsVisible = false;

                    var PrecursorMSline = zedGraphControl_PrecursorMZ.GraphPane.AddCurve("XIC neg", times, negEICs,
                                       //Chrom.Select(x => x.RetentionTime.GetValueOrDefault(0)).ToArray<double>(),
                                       //Chrom.Select(x => x.IonIT.GetValueOrDefault(0)).ToArray<double>(),
                                       Color.Red);
                    PrecursorMSline.Symbol.IsVisible = false;

                    var PrecursorMSline1 = zedGraphControl_PrecursorMZ.GraphPane.AddCurve("TIC ", times, TICs,
                                     //Chrom.Select(x => x.RetentionTime.GetValueOrDefault(0)).ToArray<double>(),
                                     //Chrom.Select(x => x.IonIT.GetValueOrDefault(0)).ToArray<double>(),
                                     Color.Brown);
                    PrecursorMSline1.Symbol.IsVisible = false;

                    var PrecursorMSline2 = zedGraphControl_PrecursorMZ.GraphPane.AddCurve("XIC all", times, AllMzs,
                                     //Chrom.Select(x => x.RetentionTime.GetValueOrDefault(0)).ToArray<double>(),
                                     //Chrom.Select(x => x.IonIT.GetValueOrDefault(0)).ToArray<double>(),
                                     Color.Blue);
                    PrecursorMSline2.Symbol.IsVisible = false;



                }


                /*
                                zedGraphControl_IIT.GraphPane.AddCurve("IIT", times, IITs,
                                                                       //Chrom.Select(x => x.RetentionTime.GetValueOrDefault(0)).ToArray<double>(),
                                                                       //Chrom.Select(x => x.IonIT.GetValueOrDefault(0)).ToArray<double>(),
                                                                       Color.Red);

                                zedGraphControl_TIC.GraphPane.AddCurve("TIC", times, TICs,
                                                       //Chrom.Select(x => x.RetentionTime.GetValueOrDefault(0)).ToArray<double>(),
                                                       //Chrom.Select(x => x.IonIT.GetValueOrDefault(0)).ToArray<double>(),
                                                       Color.Red);

                */
                //zedGraphControl_IIT.GraphPane.AddCurve()

                //Exception a = new Exception("Stop");


                MessageBox.Show("Ending chrom extract");

                /*
                zedGraphControl_IIT.GraphPane.AddCurve("IIT", times, IITs, Color.DarkBlue);
                // zedGraphControl_PrecursorMZ.GraphPane.AddCurve("Precursors", times, precursorMZ, Color.Green);
                var PrecursorMSline = zedGraphControl_PrecursorMZ.GraphPane.AddCurve("Precursors", times, precursorMZ, Color.Green, ZedGraph.SymbolType.Star);
                PrecursorMSline.Line.IsVisible = false;

            */
                //zedGraphControl_TIC.AxisChange(); zedGraphControl_TIC.Refresh();
                //zedGraphControl_IIT.AxisChange();  zedGraphControl_IIT.Refresh();
                zedGraphControlTest.AxisChange(); zedGraphControlTest.Refresh();
                zedGraphControl_PrecursorMZ.AxisChange(); zedGraphControl_PrecursorMZ.Refresh();

                MessageBox.Show("done");

            }

        }
    }
}
  