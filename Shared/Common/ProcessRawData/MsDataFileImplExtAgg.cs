using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using pwiz.ProteowizardWrapper;

namespace WashU.BatemanLab.MassSpec.Tools.ProcessRawData
{
    #region MsDataFileImplExtAgg definition
    class MsDataFileImplExtAgg
    {
        private MsDataFileImpl _msDataFileImpl;

        public MsDataFileImpl MsDataFile
        {
            get { return _msDataFileImpl; }
        }

        public MsDataFileImplExtAgg(string path)
        {
            _msDataFileImpl = new MsDataFileImpl(path);
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
