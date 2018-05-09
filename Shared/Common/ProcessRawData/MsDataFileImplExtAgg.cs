using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using pwiz.ProteowizardWrapper;

namespace WashU.BatemanLab.MassSpec.Tools.ProcessRawData
{
    #region MsDataFileImplExtAgg definition
    public class MsDataFileImplExtAgg
    {
        private MsDataFileImpl _msDataFileImpl;
        private bool _HasMzDataSpectrums = false;
        private bool _HasChromatograms = false;
        private bool _HasBeenRead = false;

        public bool HasMzDataSpectrums { get { return _HasMzDataSpectrums; } }
        public bool HasChromatograms { get { return _HasChromatograms; } }
        public bool HasBeenRead { get { return _HasBeenRead; } }

        public MsDataFileImpl MsDataFile
        {
            get { return _msDataFileImpl; }
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

        public void GetMsDataSpectrum()
        {
            try
            {
                _msDataFileImpl.GetMsDataSpectrums();
            }
            catch (Exception e)
            { }

            _HasMzDataSpectrums = true;
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
