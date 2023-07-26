using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using pwiz.ProteowizardWrapper;
using WashU.BatemanLab.MassSpec.Tools.Analysis;


namespace WashU.BatemanLab.MassSpec.Tools.ProcessRawData
{
    #region MsDataFileImplExtAgg definition
    [Serializable]
    public class MsDataFileImplExtAgg
    {
        [NonSerialized]
        private MsDataFileImpl _msDataFileImpl;
        private bool _HasMzDataSpectrums = false;
        private bool _HasChromatograms = false;
        private bool _HasBeenRead = false;
        private List<Chromatogramm> _chromatograms;
        
        public bool HasMzDataSpectrums { get { return _HasMzDataSpectrums; } }
        public bool HasChromatograms { get { return _HasChromatograms; } }
        public bool HasBeenRead { get { return _HasBeenRead; } }

        public string _msrunFileName;

        public string MSrunFileName { get { return _msrunFileName; } }
        

        public MsDataFileImplExtAgg()
        {
            _chromatograms = new List<Chromatogramm>();
        }

        public MsDataFileImpl MsDataFile
        {
            get { return _msDataFileImpl; }
        }

        public List<Chromatogramm> Chromatograms
        {
            get { return _chromatograms; }
        }

        public MsDataFileImplExtAgg(string path)
        {
            try
            {
                _msDataFileImpl = new MsDataFileImpl(path);
                _msrunFileName = Path.GetFileName(path);
            }
            catch (Exception e)
            { }

            _HasBeenRead = true;
        }

        public void GetMsDataSpectrums(ReadAndAnalyzeProgressInfo progress)
        {
            try
            {
                _msDataFileImpl.GetMsDataSpectrums(progress);
            }
            catch (Exception e)
            { }

            _HasMzDataSpectrums = true;
        }
        public void GetMsDataSpectrums()
        {
            try
            {
                _msDataFileImpl.GetMsDataSpectrums();
            }
            catch (Exception e)
            { }

            _HasMzDataSpectrums = true;
        }

        public void SortMsDataSpectrum()
        {
            try
            {
                _msDataFileImpl.SortMsDataSpectrums();
            }
            catch (Exception e)
            { }

            _HasMzDataSpectrums = true;
        }

        public void TESTMS3()
        {
            List<string> _strOutput = new List<string>();
            for (int s = 1; s < _msDataFileImpl.SpectrumCount; s++)
            {
                _strOutput.Add("Index" + _msDataFileImpl.MsDataSpectrums[s].Index.ToString() +
                               "; MsLevel - " + _msDataFileImpl.MsDataSpectrums[s].Level.ToString() +
                               "; Precursor - " + _msDataFileImpl.MsDataSpectrums[s].PrecursorMZ.ToString() );
                                
            }

            File.WriteAllLines(@"d:\_TEMP\2019-12-09\out.txt", _strOutput.ToArray());
        }
        public void ITFtoMS2(double ITFprecursor, double MS2precursor, double tolerance)
        {
            List<string> _strOutput = new List<string>();
            for (int s = 1; s < _msDataFileImpl.SpectrumCount; s++)
            {
                if (_msDataFileImpl.MsDataSpectrums[s].PrecursorMZ > ITFprecursor - tolerance &&
                    _msDataFileImpl.MsDataSpectrums[s].PrecursorMZ < ITFprecursor + tolerance)
                {
                    _msDataFileImpl.MsDataSpectrums[s].PrecursorMZ = MS2precursor;
                }


            }

            File.WriteAllLines(@"d:\_TEMP\2019-12-09\out.txt", _strOutput.ToArray());
        }

        public void CorrectForITF(Dictionary<double, double> correctionDictionary)
        {
            for (int spectrum = 0; spectrum < _msDataFileImpl.SpectrumCount; spectrum++)
            {
                if ( _msDataFileImpl.MsDataSpectrums[spectrum].PrecursorMZ.HasValue )
                {
                    double ITFprecursor = _msDataFileImpl.MsDataSpectrums[spectrum].PrecursorMZ.Value;
                    if ( correctionDictionary.ContainsKey( ITFprecursor ) )
                    {
                        _msDataFileImpl.MsDataSpectrums[spectrum].PrecursorMZ = correctionDictionary[ITFprecursor];

                    }


                }
            }
        }

        public void CorrectForITF1(Dictionary<double, double> correctionDictionary)
        {
            foreach ( var ITFtoMS2pair in correctionDictionary)
            {
                double ITFprecursor = ITFtoMS2pair.Key, MS2precursor = ITFtoMS2pair.Value;

                for (int spectrum = 0; spectrum < _msDataFileImpl.SpectrumCount; spectrum++)
                {
                    if (_msDataFileImpl.MsDataSpectrums[spectrum].PrecursorMZ == ITFprecursor)
                    {
                        _msDataFileImpl.MsDataSpectrums[spectrum].PrecursorMZ = MS2precursor;
                    }

                }
            }
        }

        public void GetChromatograms(double tolerance)
        {
            _chromatograms = new List<Chromatogramm>();
            var Channels = from spectrum in _msDataFileImpl.MsDataSpectrums
                           group spectrum by spectrum.PrecursorMZ into spectrumgroup
                           select new { ChannelMS1 = spectrumgroup.Key, ChannelSpectrums = spectrumgroup };
            var Proteins =  Analysis.AnalysisTargets.GetDefaultProteins();
            var Targets = from protein in Proteins
                          from peptide in protein.Peptides
                          from precursor in peptide.Precursors
                          select new
                          {
                              ProteinName = protein.Name,
                              PeptideName = peptide.Name,
                              PrecursorIsoform = precursor.IsotopeLabelType,
                              PrecursorMZ = precursor.PrecursorMZ,
                              Products = precursor.Products
                          };
            foreach (var Target in Targets)
            {
                var ExtraChromatograms = from mzspectrum in Channels
                                                         .Where(ch => ProcessRawDataTools.InMZTolerance(ch.ChannelMS1, Target.PrecursorMZ, tolerance))
                                                         .Select(s => s.ChannelSpectrums).FirstOrDefault()
                                         select new
                                         {
                                             mzspectrum.PrecursorMZ,
                                             mzspectrum.RetentionTime,
                                             mzspectrum.IonIT,
                                             mzspectrum.TIC,
                                             SumOfIntensities = mzspectrum.Intensities.Sum(),
                                             SumOfPositiveMatch = ProcessRawDataTools.AggIonCountsOld(mzspectrum.Mzs, mzspectrum.Intensities, Target.Products, tolerance)[0],
                                             SumOfNegativeMatch = ProcessRawDataTools.AggIonCountsOld(mzspectrum.Mzs, mzspectrum.Intensities, Target.Products, tolerance)[1]
                                         };

                _chromatograms.Add(new Chromatogramm()
                {
                    Protein = Target.ProteinName,
                    Peptide = Target.PeptideName,
                    IsotopeLabelType = Target.PrecursorIsoform,
                    PrecursorMZ = Target.PrecursorMZ,
                    RetentionTimes = ExtraChromatograms.Select(ec => ec.RetentionTime.GetValueOrDefault(0)).ToArray(),
                    IonInjectionTimes = ExtraChromatograms.Select(ec => ec.IonIT.GetValueOrDefault(0)).ToArray(),
                    SumOfIntensities = ExtraChromatograms.Select(ec => ec.SumOfIntensities).ToArray(),
                    SumOfPositiveMatch = ExtraChromatograms.Select(ec => ec.SumOfPositiveMatch).ToArray(),
                    SumOfNegativeMatch = ExtraChromatograms.Select(ec => ec.SumOfNegativeMatch).ToArray()
                });
            }
        }

        public void GetChromatograms(List<Protein> targets, double MS1tolerance, double MS2tolerance)
        {
            MsDataSpectrum defaultSpectrum = new MsDataSpectrum();
            MsDataSpectrum[] msDataSpectrums = new MsDataSpectrum[1];
            msDataSpectrums[0] = defaultSpectrum;
            _chromatograms = new List<Chromatogramm>();
            var Channels = (from spectrum in _msDataFileImpl.MsDataSpectrums
                           group spectrum by spectrum.PrecursorMZ into spectrumgroup
                           select new { ChannelMS1 = spectrumgroup.Key, ChannelSpectrums = spectrumgroup }).ToDictionary(di => di.ChannelMS1, di => di.ChannelSpectrums);
            var Proteins = targets;
            var Targets = from protein in Proteins
                          from peptide in protein.Peptides
                          from precursor in peptide.Precursors
                          select new
                          {
                              ProteinName = protein.Name,
                              PeptideName = peptide.Name,
                              PrecursorIsoform = precursor.IsotopeLabelType,
                              PrecursorMZ = precursor.PrecursorMZ,
                              Products = precursor.Products,
                              ProductsM1 = precursor.ProductsM1,
                              ProductsM2 = precursor.ProductsM2,
                              ProductsM3 = precursor.ProductsM3,
                              ProductsMneg1 = precursor.ProductsMneg1
                          };
            foreach (var Target in Targets)
            {
                var SpectrumsForChannel = from channel in Channels
                                          where ProcessRawDataTools.InMZTolerance(channel.Key, Target.PrecursorMZ, MS1tolerance) == true
                                          select channel.Value;
                if (SpectrumsForChannel.Any())
                {
                    var ExtraChromatograms = from mzspectrum in SpectrumsForChannel.Single()
                                             select new
                                             {
                                                 mzspectrum.PrecursorMZ,
                                                 mzspectrum.RetentionTime,
                                                 mzspectrum.IonIT,
                                                 mzspectrum.TIC,
                                                 SumOfIntensities = mzspectrum.Intensities.Sum(),
                                                 SumOfMatches = ProcessRawDataTools.AggIonCounts
                                                                                         (mzspectrum.Mzs, mzspectrum.Intensities, MS2tolerance,
                                                                                         Target.Products,
                                                                                         Target.ProductsM1, Target.ProductsM2, Target.ProductsM3, Target.ProductsMneg1)
                                                 //SumOfPositiveMatch = ProcessRawDataTools.AggIonCounts(mzspectrum.Mzs, mzspectrum.Intensities, Target.Products, tolerance)[0],
                                                 //SumOfNegativeMatch = ProcessRawDataTools.AggIonCounts(mzspectrum.Mzs, mzspectrum.Intensities, Target.Products, tolerance)[1]
                                             };
                    _chromatograms.Add(new Chromatogramm()
                    {
                        Protein = Target.ProteinName,
                        Peptide = Target.PeptideName,
                        IsotopeLabelType = Target.PrecursorIsoform,
                        PrecursorMZ = Target.PrecursorMZ,
                        RetentionTimes = ExtraChromatograms.Select(ec => ec.RetentionTime.GetValueOrDefault(0)).ToArray(),
                        TotalIonCurrents = ExtraChromatograms.Select(ec => ec.TIC.GetValueOrDefault(0)).ToArray(),
                        IonInjectionTimes = ExtraChromatograms.Select(ec => ec.IonIT.GetValueOrDefault(0)).ToArray(),
                        SumOfIntensities = ExtraChromatograms.Select(ec => ec.SumOfIntensities).ToArray(),

                        SumOfNegativeMatch =      ExtraChromatograms.Select(ec => ec.SumOfMatches[0]).ToArray(),
                        SumOfPositiveMatch_M0 =   ExtraChromatograms.Select(ec => ec.SumOfMatches[1]).ToArray(),
                        SumOfPositiveMatch_M1 =   ExtraChromatograms.Select(ec => ec.SumOfMatches[2]).ToArray(),
                        SumOfPositiveMatch_M2 =   ExtraChromatograms.Select(ec => ec.SumOfMatches[3]).ToArray(),
                        SumOfPositiveMatch_M3 =   ExtraChromatograms.Select(ec => ec.SumOfMatches[4]).ToArray(),
                        SumOfPositiveMatch_Mneg1= ExtraChromatograms.Select(ec => ec.SumOfMatches[5]).ToArray(),
                        SumOfPositiveMatch =      ExtraChromatograms.Select(ec => ec.SumOfMatches[6]).ToArray(),

                        
                    });
                }
            }
        }
    }
    #endregion

    #region MsDataFileImplExtInh definition
    public class MsDataFileImplExtInh : MsDataFileImpl
    {
        public MsDataFileImplExtInh(string path, int sampleIndex = 0, LockMassParameters lockmassParameters = null, bool simAsSpectra = false,
                                    bool srmAsSpectra = false, bool acceptZeroLengthSpectra = true, bool requireVendorCentroidedMS1 = false,
                                    bool requireVendorCentroidedMS2 = false, bool ignoreZeroIntensityPoints = false)
                             : base(path, sampleIndex = 0, lockmassParameters = null, simAsSpectra = false, srmAsSpectra = false, acceptZeroLengthSpectra = true,
                                    requireVendorCentroidedMS1 = false, requireVendorCentroidedMS2 = false, ignoreZeroIntensityPoints = false)
        { }
    }
    #endregion
}
