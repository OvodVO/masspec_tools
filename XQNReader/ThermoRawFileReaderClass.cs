using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
//using XRAW = MSFileReaderLib; 
//using XRAW = MSFileReaderNET;
using XRAWLib = XRAWFILE2Lib;
using Microsoft.Win32;

namespace WashU.BatemanLab.Tools.ThermoRawFileReader
{
    class ThermoRawFileReaderClass
    {
        public class IonData 
        {
            public double _MassOverCharge;
            public double _Intensity;
            
            public double Intensity
            {
                get {  return _Intensity; }
                set { _Intensity = value; }
            }

            public IonData(double initMassOverCharge, double initIntensity )
            {
                _MassOverCharge = initMassOverCharge;
                _Intensity = Intensity;
            }
        }

        public struct CompoundInjectTime
        {
            public string _Compound;
            public float _InjectTime;
            public int _ScanNum;
        }

        public class PeptidePeaksExtra
        {
            public List<CompoundInjectTime> _allInjectTimes;
            public string _Peptide;
            public float _AvgC13_InjT;
            public int _NumSpecC13_InjT;
            public float _AvgC12_InjT;
            public int _NumSpecC12_InjT;
            public float _AvgN15_InjT;
            public int _NumSpecN15_InjT;
            public int _TotScans;

        }

        public static List<double> IonstForQuant = new List<double>(new double[] { 816.4648, 915.5332, 1029.5761, 1128.6445, 1227.7130, /*Abeta42 C12*/
                                                                                   822.4849, 921.5533, 1035.5963, 1134.6647, 1233.7331, /*Abeta42 C13*/
                                                                                   825.4381, 925.5035, 1041.5405, 1141.6060, 1241.6714, /*Abeta42 N15*/
                                                                                   /*816.4648, 915.5332,*/ 972.5574, /*1029.5761, 1128.6445,*/  /*Abeta40 C12*/
                                                                                   /*822.4849, 921.5533,*/ 978.5748, /*1035.5963, 1134.6647,*/  /*Abeta40 C13*/
                                                                                   /*825.4381, 925.5035,*/ 983.5220, /*1041.5405, 1141.6060,*/  /*Abeta40 N15*/
                                                                                   653.4345, /*816.4648, 915.5332,*/                        /*Abeta38 C12*/
                                                                                   659.4546, /*822.4849, 921.5533,*/                        /*Abeta38 C13*/
                                                                                   661.4108, /*825.4391, 925.5035 */                        /*Abeta38 N15*/
                                                                                   /*Need to remove duplicate*/
                                                                                   177.1100,
                                                                                   715.8957, 718.8732, 723.8430, /*Contaminant for 42 ?*/
                                                                                   623.78, 626.75, 630.73        /*Contaminant for 40 ?*/
        
        });

        public static List<IonData> InitIonsForQuantList()
        {
            List<IonData> newList = new List<IonData>();
            foreach (double Ion in IonstForQuant)
            {
                newList.Add(new IonData(Ion, 0.0));
            }
            return newList;
        }

        public static int GetScanNum(string RawFileName)
        {
            XRAWLib.XRawfile _rawfile = new XRAWLib.XRawfile();
            
            _rawfile.Open(RawFileName);
            _rawfile.SetCurrentController(0, 1);

            int /*pnFirstSpectrum = 0,*/ pnLastSpectrum = 0;
            _rawfile.GetLastSpectrumNumber(ref pnLastSpectrum);

            //MessageBox.Show(pnLastSpectrum.ToString());

            _rawfile.Close();

            return pnLastSpectrum;
        }

        public static void DumpScanHeaderData(string RawFileName, string RawFileOut, bool IfHeader)
        {
            List<string> _strExportRAW = new List<string>();

            XRAWLib.XRawfile _rawfile = new XRAWLib.XRawfile();
            //XRAWFILE2Lib.XRawfile _rawfile = new XRawfile();

            _rawfile.Open(RawFileName);
            _rawfile.SetCurrentController(0, 1);

            int testSTLog = 0, testTuneData = 0, testerror = 0;
            _rawfile.GetNumStatusLog(ref testSTLog); MessageBox.Show("testSTLog - " + testSTLog.ToString());
            _rawfile.GetNumStatusLog(ref testTuneData); MessageBox.Show("testTuneData - " + testTuneData.ToString());
            _rawfile.GetNumStatusLog(ref testerror); MessageBox.Show("testerror - " + testerror.ToString());

            string pbstrFileName = null, pbstrInstMethod = null;

            int pnNumPackets = 0;
            double dStartTime = 0.0, pdRT = 0.0, pdStatusLogRT = 0.0;
            double dLowMass = 0.0;
            double dHighMass = 0.0;
            double dTIC = 0.0;
            double dBasePeakMass = 0.0;
            double dBasePeakIntensity = 0.0;
            int nChannels = 0;
            int bUniformTime = 0;
            double dFrequency = 0.0;

            int pnFirstSpectrum = 0, pnLastSpectrum = 0, pnArraySizeTE = 0, pnArraySizeSL = 0;

            DateTime pCreationDate = DateTime.Now;

            _rawfile.GetFileName(ref pbstrFileName); _rawfile.GetInstMethod(1, ref pbstrInstMethod); _rawfile.GetCreationDate(ref pCreationDate);

            _rawfile.GetFirstSpectrumNumber(ref pnFirstSpectrum); _rawfile.GetLastSpectrumNumber(ref pnLastSpectrum);

            int count = 0;

            for (int nScanNumber = pnFirstSpectrum; nScanNumber <= pnLastSpectrum; nScanNumber++)  // all spectrums
            {
                string pbstrFilter = null, pbstrScanEvent = null;
                object pvarLabelsTE = null, pvarValuesTE = null;
                object pvarLabelsSL = null, pvarValuesSL = null;
                count++;

                _rawfile.GetScanHeaderInfoForScanNum(nScanNumber, ref pnNumPackets, ref dStartTime, ref dLowMass, ref dHighMass, ref dTIC,
                                                     ref dBasePeakMass, ref dBasePeakIntensity, ref nChannels, ref bUniformTime, ref dFrequency);
                _rawfile.GetTrailerExtraForScanNum(nScanNumber, ref pvarLabelsTE, ref pvarValuesTE, ref pnArraySizeTE);

                _rawfile.GetFilterForScanNum(nScanNumber, ref pbstrFilter); //MessageBox.Show(pbstrFilter);

                _rawfile.RTFromScanNum(nScanNumber, ref pdRT);

                (_rawfile as XRAWLib.IXRawfile5).GetScanEventForScanNum(nScanNumber, ref pbstrScanEvent);

                _rawfile.GetStatusLogForScanNum(nScanNumber, ref pdStatusLogRT, ref pvarLabelsSL, ref pvarValuesSL, ref pnArraySizeSL); MessageBox.Show(pnArraySizeSL.ToString());

                string[] strTrailerLabels = (string[])pvarLabelsTE; string[] strTrailerValues = (string[])pvarValuesTE;
                string[] strStatusLabels = (string[])pvarLabelsSL; string[] strStatusValues = (string[])pvarValuesSL;

                string trailerExtra = "", trailerHeader = ""; string statusLog = "", statusHeader = "";

                for (int i = 0; i < pnArraySizeTE; i++)  // all TrailerExtra
                {
                    trailerExtra += strTrailerValues[i]; trailerExtra += "; ";

                    if (count == 1)
                    {
                        trailerHeader += strTrailerLabels[i]; trailerHeader += "; ";
                    }

                }

                for (int i = 0; i < pnArraySizeSL; i++)  // all StatusLog
                {
                    statusLog += strStatusValues[i]; statusLog += "; "; MessageBox.Show(strStatusLabels[i] + " - " + strStatusValues[i]);

                    if (count == 1)
                    {
                        statusHeader += strStatusLabels[i]; statusHeader += "; "; 
                    }

                }

                if (count == 1)
                {
                    _strExportRAW.Add("pbstrFileName; pCreationDate; nScanNumber; pbstrFilter; pdRT; pbstrScanEvent; pnNumPackets; dStartTime; dLowMass; " +
                                      "dHighMass; dTIC; dBasePeakMass; dBasePeakIntensity; nChannels; bUniformTime; dFrequency; " + trailerHeader);
                }

                _strExportRAW.Add(
                String.Format("{0}; {1}; {2}; {3}; {4}; {5}; {6}; {7}; {8}; {9}; {10}; {11}; {12}; {13}; {14}; {15}; {16}",
                               pbstrFileName, pCreationDate,
                               nScanNumber, pbstrFilter, pdRT, pbstrScanEvent,
                               pnNumPackets, dStartTime, dLowMass, dHighMass, dTIC, dBasePeakMass, dBasePeakIntensity, nChannels, bUniformTime, dFrequency,
                               trailerExtra)
                               );

            }

            if (IfHeader)
            {
                File.WriteAllLines(RawFileOut, _strExportRAW.ToArray());
            }
            else
            {
                File.AppendAllLines(RawFileOut, _strExportRAW.ToArray());
            }


            _rawfile.Close();

        }

        public static int DumpInstrumentMethod(string RawFileName)
        {
            List<string> _strExportRAW = new List<string>();

            XRAWLib.XRawfile _rawfile = new XRAWLib.XRawfile();
            //XRAWFILE2Lib.XRawfile _rawfile = new XRawfile();

            _rawfile.Open(RawFileName);
            _rawfile.SetCurrentController(0, 1);

            //_rawfile.GetInstMethod()
            /* 
             int testSTLog = 0, testTuneData = 0, testerror = 0;
             _rawfile.GetNumStatusLog(ref testSTLog); MessageBox.Show("testSTLog - " + testSTLog.ToString());
             _rawfile.GetNumTuneData(ref testTuneData); MessageBox.Show("testTuneData - " + testTuneData.ToString());
             _rawfile.GetNumErrorLog(ref testerror); MessageBox.Show("testerror - " + testerror.ToString()); */

            string pbstrFileName = null, pbstrInstMethod0 = null, pbstrInstMethod1 = null ;

            int pnNumPackets = 0;
            double dStartTime = 0.0, pdRT = 0.0, pdStatusLogRT = 0.0;
            double dLowMass = 0.0;
            double dHighMass = 0.0;
            double dTIC = 0.0;
            double dBasePeakMass = 0.0;
            double dBasePeakIntensity = 0.0;
            int nChannels = 0;
            int bUniformTime = 0;
            double dFrequency = 0.0;

            int pnFirstSpectrum = 0, pnLastSpectrum = 0, pnArraySizeTE = 0, pnArraySizeSL = 0;

            DateTime pCreationDate = DateTime.Now;

            _rawfile.GetFileName(ref pbstrFileName); 
            _rawfile.GetInstMethod(0, ref pbstrInstMethod0);  _rawfile.GetInstMethod(1, ref pbstrInstMethod1); 
            _rawfile.GetCreationDate(ref pCreationDate);

            
            File.WriteAllText(Path.ChangeExtension(RawFileName, "inm"), pbstrInstMethod0 + pbstrInstMethod1);

            //MessageBox.Show(pbstrInstMethod0);

            string AquDate = null;

            _rawfile.GetInstSerialNumber (ref AquDate);

            MessageBox.Show(pCreationDate.ToString(), "Creatio");

            MessageBox.Show(AquDate, "AquDate");

            return 1;

        }


        public static int DumpScanHeaderDataForCompound(string RawFileName, string RawFileOut, bool IfHeader, string Compound, double PeakLeftRT, double PeakRightRT, double Tolerance)
        {
            List<string> _strExportRAW = new List<string>();

            XRAWLib.XRawfile _rawfile = new XRAWLib.XRawfile();
            //XRAWFILE2Lib.XRawfile _rawfile = new XRawfile();

            _rawfile.Open(RawFileName);
            _rawfile.SetCurrentController(0, 1);

            //_rawfile.GetInstMethod()
           /* 
            int testSTLog = 0, testTuneData = 0, testerror = 0;
            _rawfile.GetNumStatusLog(ref testSTLog); MessageBox.Show("testSTLog - " + testSTLog.ToString());
            _rawfile.GetNumTuneData(ref testTuneData); MessageBox.Show("testTuneData - " + testTuneData.ToString());
            _rawfile.GetNumErrorLog(ref testerror); MessageBox.Show("testerror - " + testerror.ToString()); */

            string pbstrFileName = null, pbstrInstMethod = null;

            int pnNumPackets = 0;
            double dStartTime = 0.0, pdRT = 0.0, pdStatusLogRT = 0.0;
            double dLowMass = 0.0;
            double dHighMass = 0.0;
            double dTIC = 0.0;
            double dBasePeakMass = 0.0;
            double dBasePeakIntensity = 0.0;
            int nChannels = 0;
            int bUniformTime = 0;
            double dFrequency = 0.0;

            int pnFirstSpectrum = 0, pnLastSpectrum = 0, pnArraySizeTE = 0, pnArraySizeSL = 0;

            DateTime pCreationDate = DateTime.Now;

            _rawfile.GetFileName(ref pbstrFileName); _rawfile.GetInstMethod(1, ref pbstrInstMethod); _rawfile.GetCreationDate(ref pCreationDate);

            pbstrFileName = Path.GetFileNameWithoutExtension(pbstrFileName);
            
            _rawfile.GetFirstSpectrumNumber(ref pnFirstSpectrum); _rawfile.GetLastSpectrumNumber(ref pnLastSpectrum);

            int count = 0;

            for (int nScanNumber = pnFirstSpectrum; nScanNumber <= pnLastSpectrum; nScanNumber++)  // all spectrums
            {
                List<IonData> _ScanMSDataList = InitIonsForQuantList();
                double _tii = GetTotalIntensityOfIonFromScanNumber(_rawfile, nScanNumber, _ScanMSDataList, Tolerance);

                string pbstrFilter = null, pbstrScanEvent = null;
                object pvarLabelsTE = null, pvarValuesTE = null;
                object pvarLabelsSL = null, pvarValuesSL = null;
                

                _rawfile.GetScanHeaderInfoForScanNum(nScanNumber, ref pnNumPackets, ref dStartTime, ref dLowMass, ref dHighMass, ref dTIC,
                                                     ref dBasePeakMass, ref dBasePeakIntensity, ref nChannels, ref bUniformTime, ref dFrequency);
                _rawfile.GetTrailerExtraForScanNum(nScanNumber, ref pvarLabelsTE, ref pvarValuesTE, ref pnArraySizeTE);

                _rawfile.GetFilterForScanNum(nScanNumber, ref pbstrFilter);

                string compoundName = CompoundFromFilter(pbstrFilter);

                if (!IfCompound(compoundName, Compound)) { continue; }

                count++;
               
                bool _ifPeak = false;

                _rawfile.RTFromScanNum(nScanNumber, ref pdRT); if (pdRT > PeakLeftRT && pdRT < PeakRightRT) { _ifPeak = true; }

                (_rawfile as XRAWLib.IXRawfile5).GetScanEventForScanNum(nScanNumber, ref pbstrScanEvent);

                _rawfile.GetStatusLogForScanNum(nScanNumber, ref pdStatusLogRT, ref pvarLabelsSL, ref pvarValuesSL, ref pnArraySizeSL); //MessageBox.Show(pnArraySizeSL.ToString());

                string[] strTrailerLabels = (string[])pvarLabelsTE; string[] strTrailerValues = (string[])pvarValuesTE;
                string[] strStatusLabels = (string[])pvarLabelsSL; string[] strStatusValues = (string[])pvarValuesSL;

                string trailerExtra = "", trailerHeader = ""; string statusLog = "", statusHeader = ""; string quantIons = "",  quantIonsHeader = "";

                for (int i = 0; i < pnArraySizeTE; i++)  // all TrailerExtra
                {
                    trailerExtra += strTrailerValues[i]; trailerExtra += "; ";

                    if (count == 1 && IfHeader)
                    {
                        trailerHeader += strTrailerLabels[i]; trailerHeader += "; ";
                    }

                }

                foreach (IonData _ion in _ScanMSDataList)  // all Quant MS ions-intens.
                {
                    quantIons += _ion._Intensity.ToString(); quantIons += "; ";

                    if (count == 1 && IfHeader)
                    {
                        quantIonsHeader += _ion._MassOverCharge.ToString(); quantIonsHeader += "; ";
                    }

                }

                for (int i = 0; i < pnArraySizeSL; i++)  // all StatusLog
                {
                    statusLog += strStatusValues[i]; statusLog += "; "; //MessageBox.Show(strStatusLabels[i] + " - " + strStatusValues[i]);

                    if (count == 1 && IfHeader)
                    {
                        statusHeader += strStatusLabels[i]; statusHeader += "; ";
                    }

                }

                if (count == 1 && IfHeader)
                {
                   
                    _strExportRAW.Add("pbstrFileName; pCreationDate; nScanNumber; pbstrFilter; pdRT; pbstrScanEvent; pnNumPackets; dStartTime; dLowMass; " +
                                      "dHighMass; dTIC; dBasePeakMass; dBasePeakIntensity; nChannels; bUniformTime; dFrequency; ifPeak; compound; TII; PeakLeftRT; PeakRightRT; " 
                                      + quantIonsHeader + trailerHeader);
                }

                _strExportRAW.Add(
                String.Format("{0}; {1}; {2}; {3}; {4}; {5}; {6}; {7}; {8}; {9}; {10}; {11}; {12}; {13}; {14}; {15}; {16}; {17}; {18}; {19}; {20}; {21} {22}",
                               pbstrFileName, pCreationDate,
                               nScanNumber, pbstrFilter, pdRT, pbstrScanEvent,
                               pnNumPackets, dStartTime, dLowMass, dHighMass, dTIC, dBasePeakMass, dBasePeakIntensity, nChannels, bUniformTime, dFrequency,
                               _ifPeak, compoundName, _tii, PeakLeftRT, PeakRightRT,
                               quantIons,
                               trailerExtra)
                               );

            }

            if (IfHeader)
            {
                File.WriteAllLines(RawFileOut, _strExportRAW.ToArray());
            }
            else
            {
                File.AppendAllLines(RawFileOut, _strExportRAW.ToArray());
            }


            _rawfile.Close();

            return count;

        }


        public static PeptidePeaksExtra GetPeptideExtrasForCompound(string RawFileName, string Compound, double PeakLeftRT, double PeakRightRT)
        {
            PeptidePeaksExtra _peptideExtras = new PeptidePeaksExtra();

            _peptideExtras._allInjectTimes = new List<CompoundInjectTime>();

            _peptideExtras._Peptide = PeptideFromCompound(Compound);

            //List<CompoundInjectTime> _listCompVSInj = new List<CompoundInjectTime>();

            XRAWLib.XRawfile _rawfile = new XRAWLib.XRawfile();

            _rawfile.Open(RawFileName);
            _rawfile.SetCurrentController(0, 1);

            

            int pnFirstSpectrum = 0, pnLastSpectrum = 0, pnArraySizeTE = 0;

            _rawfile.ScanNumFromRT(PeakLeftRT, ref pnFirstSpectrum); _rawfile.ScanNumFromRT(PeakRightRT, ref pnLastSpectrum); 

            int count = 0;

            

            for (int nScanNumber = pnFirstSpectrum; nScanNumber <= pnLastSpectrum; nScanNumber++)  // all spectrums
            {
                string pbstrFilter = null, pbstrScanEvent = null;

                object pvarLabelsTE = null, pvarValuesTE = null;

                _rawfile.GetFilterForScanNum(nScanNumber, ref pbstrFilter);

                

               // MessageBox.Show(nScanNumber.ToString() +" - "+ pbstrFilter+" peakleft " + PeakLeftRT);

               // if (pbstrFilter == null) { MessageBox.Show(RawFileName + "have null filter"); continue; }

                string compoundName = CompoundFromFilter(pbstrFilter);
                string peptideName = PeptideFromFilter(pbstrFilter);

                if (!IfCompound(peptideName, Compound)) { continue; }

                count++;

                (_rawfile as XRAWLib.IXRawfile5).GetScanEventForScanNum(nScanNumber, ref pbstrScanEvent);

                _rawfile.GetTrailerExtraForScanNum(nScanNumber, ref pvarLabelsTE, ref pvarValuesTE, ref pnArraySizeTE);
                

                string[] strTrailerLabels = (string[])pvarLabelsTE; string[] strTrailerValues = (string[])pvarValuesTE;
              

                string trailerExtra = "";   //, trailerHeader = ""; string statusLog = "", statusHeader = ""; string quantIons = "", quantIonsHeader = "";


                int _injectIndex = 2;

                for (int i = 0; i < strTrailerLabels.Length ; i++ )
                {
                   
                    if (strTrailerLabels[i].StartsWith("Ion Injection Time"))
                    {
                        trailerExtra = strTrailerValues[i];
                        _injectIndex = i; 
                    }
                }

                CompoundInjectTime _compInj = new CompoundInjectTime();
                
                _compInj._ScanNum = nScanNumber;
                _compInj._Compound = compoundName;

                try
                {
                    _compInj._InjectTime = float.Parse(strTrailerValues[_injectIndex]);
                }
                catch (Exception _ex)
                {
                    MessageBox.Show("GetPeptideExtrasForCompound():float.Parse", "Cannot parse " + strTrailerValues[_injectIndex] + " to float;");
                }

                _peptideExtras._allInjectTimes.Add(_compInj);
            }
           
            _rawfile.Close();


            var AvgInjTimeQuery = from pair in _peptideExtras._allInjectTimes
                                  group pair by pair._Compound into gr
                                  select new { Compound = gr.Key, AvgInjTime = gr.Average(a => a._InjectTime), NumSpec = gr.Count() };
                                 
                                 

            foreach (var _compound in AvgInjTimeQuery)
            {

               if (_compound.Compound.Contains("C13"))
               {
                   _peptideExtras._AvgC13_InjT = _compound.AvgInjTime; _peptideExtras._NumSpecC13_InjT = _compound.NumSpec;
               }

               if (_compound.Compound.Contains("C12"))
               {
                   _peptideExtras._AvgC12_InjT = _compound.AvgInjTime; _peptideExtras._NumSpecC12_InjT = _compound.NumSpec;
               }

               if (_compound.Compound.Contains("N15"))
               {
                   _peptideExtras._AvgN15_InjT = _compound.AvgInjTime; _peptideExtras._NumSpecN15_InjT = _compound.NumSpec;
               }
            }

            _peptideExtras._TotScans = count;

            return _peptideExtras;

        }


        public static double GetTotalIntensityOfIonFromScanNumber(XRAWLib.XRawfile Rawfile, int nScanNumber, List<IonData> IonMassList, double Tolerance)
        {
            double totalIonIntensity = 0.0;
            object pvarMassList = null; object pvarPeakFlags = null; int pnArraySize = 0; double pdCentroidWidth = 0.0;
            Rawfile.GetMassListFromScanNum(nScanNumber, null, 0, 0, 0, 1, ref pdCentroidWidth, ref pvarMassList, ref pvarPeakFlags, ref pnArraySize);
            double[,] scanMassList = (double[,])pvarMassList;
            for (int i = 0; i < pnArraySize; i++ )
            {
                totalIonIntensity += scanMassList[1, i];
                for (int j = 0; j < IonMassList.Count; j++ )
                {
                    if ( scanMassList[0, i] > ( IonMassList[j]._MassOverCharge - Tolerance ) && scanMassList[0, i] < (  IonMassList[j]._MassOverCharge + Tolerance ) )
                    {
                         IonMassList[j].Intensity += scanMassList[1, i];
                    }
                }
               // MessageBox.Show(String.Format("Mass - {0}, Intensity - {1} ", mas[0, i], mas[1, i]) );
            }
            return totalIonIntensity;
        }
        public static bool IfCompound(string CompoundName, string XQNCompound)
        {
            if (XQNCompound.Contains(CompoundName)) 
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static string CompoundFromFilterN(string MSFilter)
        {
            string compound = "unknown";

            if (MSFilter.Contains("715.")) { compound = "Abeta42_N14"; }
            if (MSFilter.Contains("718.")) { compound = "Abeta42_C13"; }
            if (MSFilter.Contains("723.")) { compound = "Abeta42_N15"; }
            if (MSFilter.Contains("623.")) { compound = "Abeta40_N14"; }
            if (MSFilter.Contains("626.")) { compound = "Abeta40_C13"; }
            if (MSFilter.Contains("630.")) { compound = "Abeta40_N15"; }
            if (MSFilter.Contains("524.")) { compound = "Abeta38_N14"; } 
            if (MSFilter.Contains("527.")) { compound = "Abeta38_C13"; }
            if (MSFilter.Contains("531.")) { compound = "Abeta38_N15"; }

            return compound;
        }

        public static string PeptideFromFilter(string MSFilter)
        {
            string peptide = "unknown";

            if (MSFilter.Contains("699.")) { peptide = "Aß42"; }
            if (MSFilter.Contains("702.")) { peptide = "Aß42"; }
            if (MSFilter.Contains("707.")) { peptide = "Aß42"; }

            if (MSFilter.Contains("607.")) { peptide = "Aß40"; }
            if (MSFilter.Contains("610.")) { peptide = "Aß40"; }
            if (MSFilter.Contains("614.")) { peptide = "Aß40"; }

            if (MSFilter.Contains("508.")) { peptide = "Aß38"; }
            if (MSFilter.Contains("511.")) { peptide = "Aß38"; }
            if (MSFilter.Contains("514.")) { peptide = "Aß38"; }

            if (MSFilter.Contains("663.")) { peptide = "AßMD"; }
            if (MSFilter.Contains("666.")) { peptide = "AßMD"; }
            if (MSFilter.Contains("670.")) { peptide = "AßMD"; }

            return peptide;
        }

        public static string PeptideFromCompound(string Compound)
        {
            string peptide = "unknown";

            if (Compound.Contains("Aß42")) { peptide = "Aß42"; }
            if (Compound.Contains("Aß40")) { peptide = "Aß40"; }
            if (Compound.Contains("Aß38")) { peptide = "Aß38"; }
            if (Compound.Contains("AßMD")) { peptide = "AßMD"; }

            return peptide;
        }

        public static string CompoundFromFilter(string MSFilter)
        {
            string compound = "unknown";

              if (MSFilter.Contains("699.")) { compound = "Aß42_C12"; }
              if (MSFilter.Contains("702.")) { compound = "Aß42_C13"; }
              if (MSFilter.Contains("707.")) { compound = "Aß42_N15"; }
            
              if (MSFilter.Contains("607.")) { compound = "Aß40_C12"; }
              if (MSFilter.Contains("610.")) { compound = "Aß40_C13"; }
              if (MSFilter.Contains("614.")) { compound = "Aß40_N15"; }

              if (MSFilter.Contains("508.")) { compound = "Aß38_C12"; }
              if (MSFilter.Contains("511.")) { compound = "Aß38_C13"; }
              if (MSFilter.Contains("514.")) { compound = "Aß38_N15"; }

              if (MSFilter.Contains("663.")) { compound = "AßMD_C12"; }
              if (MSFilter.Contains("666.")) { compound = "AßMD_C13"; }
              if (MSFilter.Contains("670.")) { compound = "AßMD_N15"; }

            return compound;
        }

                
        /*
        public static void ProcessRAWFiles()
        {
            OpenFileDialog _openFileDlg = new OpenFileDialog();
            _openFileDlg.Multiselect = true;

            _openFileDlg.Filter = "raw files (*.raw)|*.raw";
            bool _ifHeader = true;
            if (_openFileDlg.ShowDialog() == true)
            {

                if (_openFileDlg.FileNames.Count() > 1)
                {
                    SaveFileDialog _saveFileDlg = new SaveFileDialog();
                    _saveFileDlg.Filter = "csv files (*.csv)|*.csv";
                    if (_saveFileDlg.ShowDialog() == true)
                    {
                        foreach (string _rawFile in _openFileDlg.FileNames)
                        {
                            DumpScanHeaderData(_rawFile, _saveFileDlg.FileName, _ifHeader);
                            _ifHeader = false;
                        }
                    }
                }
                else
                {
                    DumpScanHeaderData(_openFileDlg.FileName, System.IO.Path.ChangeExtension(_openFileDlg.FileName, "csv"), _ifHeader);

                }
            }

        } */
    }
}
