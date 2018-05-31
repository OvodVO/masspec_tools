using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using pwiz.ProteowizardWrapper;
using WashU.BatemanLab.MassSpec.Tools.AnalysisTargets;


namespace WashU.BatemanLab.MassSpec.Tools.ProcessRawData
{
    #region MsDataFileImplExtAgg definition
    public class MsDataFileImplExtAgg
    {
        private MsDataFileImpl _msDataFileImpl;
        private bool _HasMzDataSpectrums = false;
        private bool _HasChromatograms = false;
        private bool _HasBeenRead = false;
        private List<Chromatogram> _chromatograms;
        
        public bool HasMzDataSpectrums { get { return _HasMzDataSpectrums; } }
        public bool HasChromatograms { get { return _HasChromatograms; } }
        public bool HasBeenRead { get { return _HasBeenRead; } }

        public MsDataFileImplExtAgg()
        {
            _chromatograms = new List<Chromatogram>();
        }

        public MsDataFileImpl MsDataFile
        {
            get { return _msDataFileImpl; }
        }

        public List<Chromatogram> Chromatograms
        {
            get { return _chromatograms; }
        }

        public MsDataFileImplExtAgg(string path)
        {
            try
            {
                _msDataFileImpl = new MsDataFileImpl(path);
            }
            catch (Exception e)
            { }

            _HasBeenRead = true;
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

        public void GetChromatograms(double tolerance)
        {
            _chromatograms = new List<Chromatogram>();

            var Channels = from spectrum in _msDataFileImpl.MsDataSpectrums
                           group spectrum by spectrum.PrecursorMZ into spectrumgroup
                           select new { ChannelMS1 = spectrumgroup.Key, ChannelSpectrums = spectrumgroup };

            //Assing targets here
            var Proteins =  AnalysisTargets.AnalysisTargets.GetDefaultProteins();

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
                                             SumOfPositiveMatch = ProcessRawDataTools.AggIonCounts(mzspectrum.Mzs, mzspectrum.Intensities, Target.Products, tolerance)[0],
                                             SumOfNegativeMatch = ProcessRawDataTools.AggIonCounts(mzspectrum.Mzs, mzspectrum.Intensities, Target.Products, tolerance)[1]
                                         };

                _chromatograms.Add(new Chromatogram()
                {
                    Protein = Target.ProteinName,
                    Peptide = Target.ProteinName,
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

        public void GetChromatograms(List<Protein> targets, double tolerance)
        {
            MsDataSpectrum defaultSpectrum = new MsDataSpectrum();
            MsDataSpectrum[] msDataSpectrums = new MsDataSpectrum[1];
            msDataSpectrums[0] = defaultSpectrum;


            _chromatograms = new List<Chromatogram>();

            var Channels = (from spectrum in _msDataFileImpl.MsDataSpectrums
                           group spectrum by spectrum.PrecursorMZ into spectrumgroup
                           select new { ChannelMS1 = spectrumgroup.Key, ChannelSpectrums = spectrumgroup }).ToDictionary(di => di.ChannelMS1, di => di.ChannelSpectrums);

            //Assing targets here
            var Proteins = targets; //AnalysisTargets.AnalysisTargets.GetDefaultProteins();

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
                var SpectrumsForChannel = from channel in Channels
                                          where ProcessRawDataTools.InMZTolerance(channel.Key, Target.PrecursorMZ, tolerance) == true
                                          select channel.Value;

                if (!SpectrumsForChannel.Any()) { throw new System.Exception("Haha" + Target.PrecursorMZ.ToString() ); }

                var ExtraChromatograms = from mzspectrum in SpectrumsForChannel.Single()
                                                         
                                         select new
                                         {
                                             mzspectrum.PrecursorMZ,
                                             mzspectrum.RetentionTime,
                                             mzspectrum.IonIT,
                                             mzspectrum.TIC,
                                             SumOfIntensities = mzspectrum.Intensities.Sum(),
                                             SumOfPositiveMatch = ProcessRawDataTools.AggIonCounts(mzspectrum.Mzs, mzspectrum.Intensities, Target.Products, tolerance)[0],
                                             SumOfNegativeMatch = ProcessRawDataTools.AggIonCounts(mzspectrum.Mzs, mzspectrum.Intensities, Target.Products, tolerance)[1]
                                         };

                _chromatograms.Add(new Chromatogram()
                {
                    Protein = Target.ProteinName,
                    Peptide = Target.ProteinName,
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

        public void GetChromatogramsTEST(double tolerance)
        {
            _chromatograms = new List<Chromatogram>();

            var Channels = from spectrum in _msDataFileImpl.MsDataSpectrums
                           group spectrum by spectrum.PrecursorMZ into spectrumgroup
                           select new { ChannelMS1 = spectrumgroup.Key, ChannelSpectrums = spectrumgroup.ToArray() };

            List<MsDataSpectrum[]> testList = new List<MsDataSpectrum[]>(); //var test = from Ch in Channels.

            MsDataSpectrum[] test55 = new MsDataSpectrum[4000];

            Array.Copy(_msDataFileImpl.MsDataSpectrums, test55, 4000);


                foreach (var cha in Channels)
            {
             //   testList.Add(cha.ChannelSpectrums);
            }

            //Assing targets here
            var Proteins = AnalysisTargets.AnalysisTargets.GetDefaultProteins();

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

            var Channels1 = Channels.ToArray();

            foreach (var Target in Targets)
            {

                //var ExtraChromatograms = from mzspectrum in Channels1.Single(s => ProcessRawDataTools.InMZTolerance(s.ChannelMS1, Target.PrecursorMZ, tolerance) == true ).ChannelSpectrums

                                         var ExtraChromatograms = from mzspectrum in test55

                                                                  select new
                                         {
                                             mzspectrum.PrecursorMZ,
                                             mzspectrum.RetentionTime,
                                             mzspectrum.IonIT,
                                             mzspectrum.TIC,
                                             SumOfIntensities = mzspectrum.Intensities.Sum(),
                                             SumOfPositiveMatch = ProcessRawDataTools.AggIonCounts(mzspectrum.Mzs, mzspectrum.Intensities, Target.Products, tolerance)[0],
                                             SumOfNegativeMatch = ProcessRawDataTools.AggIonCounts(mzspectrum.Mzs, mzspectrum.Intensities, Target.Products, tolerance)[1]
                                         };

                _chromatograms.Add(new Chromatogram()
                {
                    Protein = Target.ProteinName,
                    Peptide = Target.ProteinName,
                    IsotopeLabelType = Target.PrecursorIsoform,
                    PrecursorMZ = Target.PrecursorMZ,
                    RetentionTimes = ExtraChromatograms.Select(ec => ec.RetentionTime.GetValueOrDefault(0)).ToArray(),
                 //   IonInjectionTimes = ExtraChromatograms.Select(ec => ec.IonIT.GetValueOrDefault(0)).ToArray(),
                 //   SumOfIntensities = ExtraChromatograms.Select(ec => ec.SumOfIntensities).ToArray(),
                    SumOfPositiveMatch = ExtraChromatograms.Select(ec => ec.SumOfPositiveMatch).ToArray(),
                 //   SumOfNegativeMatch = ExtraChromatograms.Select(ec => ec.SumOfNegativeMatch).ToArray()
                });
            }

        }

        public void GetChromatogramsTEST57(double tolerance)
        {
            _chromatograms = new List<Chromatogram>();

            var Channels = from spectrum in _msDataFileImpl.MsDataSpectrums
                           group spectrum by spectrum.PrecursorMZ into spectrumgroup
                           select new { ChannelMS1 = spectrumgroup.Key, ChannelSpectrums = spectrumgroup };

            //Assing targets here
            var Proteins = AnalysisTargets.AnalysisTargets.GetDefaultProteins();

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
                                             SumOfPositiveMatch = ProcessRawDataTools.AggIonCounts(mzspectrum.Mzs, mzspectrum.Intensities, Target.Products, tolerance)[0],
                                             SumOfNegativeMatch = ProcessRawDataTools.AggIonCounts(mzspectrum.Mzs, mzspectrum.Intensities, Target.Products, tolerance)[1]
                                         };

                _chromatograms.Add(new Chromatogram()
                {
                    Protein = Target.ProteinName,
                    Peptide = Target.ProteinName,
                    IsotopeLabelType = Target.PrecursorIsoform,
                    PrecursorMZ = Target.PrecursorMZ,
                    RetentionTimes = ExtraChromatograms.Select(ec => ec.RetentionTime.GetValueOrDefault(0)).ToArray(),
                    //  IonInjectionTimes = ExtraChromatograms.Select(ec => ec.IonIT.GetValueOrDefault(0)).ToArray(),
                    //  SumOfIntensities = ExtraChromatograms.Select(ec => ec.SumOfIntensities).ToArray(),
                    SumOfPositiveMatch = ExtraChromatograms.Select(ec => ec.SumOfPositiveMatch).ToArray(),
                    //  SumOfNegativeMatch = ExtraChromatograms.Select(ec => ec.SumOfNegativeMatch).ToArray()
                });
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
