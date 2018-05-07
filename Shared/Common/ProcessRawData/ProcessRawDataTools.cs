using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using pwiz.ProteowizardWrapper;

namespace WashU.BatemanLab.MassSpec.Tools.ProcessRawData
{
    public class MzIntensityPair
    {
        public MzIntensityPair()
        { }

        public MzIntensityPair(double Mz, double Intensity)
        {
            mz = Mz;
            intensity = Intensity;
        }
               
        public double intensity { get; set; }

        public double mz { get; set; }
    }
        public class ProcessRawDataTools
    {
        public static bool InMZTolerance(double mz, double expectedMz, double tolerance)
        {
            if (mz >= (expectedMz - tolerance) && mz <= (expectedMz + tolerance)) return true;
            else return false;
        }

        public static bool InMZTolerance(double mz, double[] expectedMzArray, double tolerance)
        {
            bool _result = false;
            for (int i = 0; i <= expectedMzArray.Length; i++)
            {
                if (InMZTolerance(mz, expectedMzArray[i], tolerance)) _result = true;
            }
            return _result;
        }

        public static bool InMZTolerance(double mz, List<double> expectedMzList, double tolerance)
        {
            bool _result = false;
            foreach (double _mz in expectedMzList) 
            {
                if (InMZTolerance(mz, _mz, tolerance)) _result = true;
            }
            return _result;
        }

        public static MzIntensityPair[] test(double[] MZs, double[] Intensities)
        {
            var mzintensitiesPairs = MZs.Zip(Intensities, (m, i) => new MzIntensityPair { mz = m,  intensity = i });

            return mzintensitiesPairs.ToArray();
        }

        public static string testOpen(string path)
        {
            MsDataFileImpl testFI = new MsDataFileImpl(path);

            int i;

            string st = testFI.GetChromatogramId(0, out i);
            return st;  // testFI.ChromatogramCount;
        }

        public static double[] testGetScanTimes(string path)
        {
            MsDataFileImpl testFI = new MsDataFileImpl(path);

            return testFI.GetScanTimes();
        }

        public static double[] testGetTotalIonCurrent(string path)
        {
            MsDataFileImpl testFI = new MsDataFileImpl(path);

            return testFI.GetTotalIonCurrent();
        }

        public static double[] testGetTotalIonCurrentSel(string path)
        {
            MsDataFileImpl testFI = new MsDataFileImpl(path); 

            List<double> _listInt = new List<double>();


            for (int i = 0; i < testFI.SpectrumCount; i++)
            {
                var _spectrum = testFI.GetSpectrum(i);
                if (_spectrum.Precursors[0].IsolationWindowTargetMz > 707 && _spectrum.Precursors[0].IsolationWindowTargetMz < 708)
                {
                    double _inten = 0;
                    for (int j = 0; j < _spectrum.Intensities.Length; j++)

                    { _inten +=_spectrum.Intensities[j]; }

                    _listInt.Add(_inten);

                    
                }
            }

            double[] selSpecIonCounts = new double[_listInt.Count];

            selSpecIonCounts = _listInt.ToArray();

            return selSpecIonCounts;
        }

        public static double[] testGetTimeSel(string path)
        {
            MsDataFileImpl testFI = new MsDataFileImpl(path);

            List<double> _listRT = new List<double>();


            for (int i = 0; i < testFI.SpectrumCount; i++)
            {
                var _spectrum = testFI.GetSpectrum(i);
                if (_spectrum.Precursors[0].IsolationWindowTargetMz > 707 && _spectrum.Precursors[0].IsolationWindowTargetMz < 708)
                {
                    _listRT.Add((double)_spectrum.RetentionTime);
                }
            }

            double[] selRetTime = new double[_listRT.Count];

            selRetTime = _listRT.ToArray();

            return selRetTime;

            
        }

        public static double[] testGetIonInjectionTime77(string path)
        {
            MsDataFileImpl testFI = new MsDataFileImpl(path);

            List<double> _listInt = new List<double>();


            for (int i = 0; i < testFI.SpectrumCount; i++)
            {
                var _spectrum = testFI.GetSpectrum(i);

                if (_spectrum.Precursors[0].IsolationWindowTargetMz > 707 && _spectrum.Precursors[0].IsolationWindowTargetMz < 708)
                {
                    double _ion_inj_time = (double) testFI.GetIonInjectionTime(i);
                 
                    _listInt.Add(_ion_inj_time);


                }
            }

            double[] selSpecIonCounts = new double[_listInt.Count];

            selSpecIonCounts = _listInt.ToArray();

            return selSpecIonCounts;
        }

        public static double testGetIonInjectionTimeVO(string path)
        {
            MsDataFileImpl testFI = new MsDataFileImpl(path);

            
            return testFI.GetIonInjectionTime(5);
        }


        public static double testGetIonInjectionTime(string path)
        {
            MsDataFileImpl testFI = new MsDataFileImpl(path);
            return testFI.GetIonInjectionTime(9840);
        }
        
    }

    public class MsDataFileImplExtInh : MsDataFileImpl 
    {

        public MsDataFileImplExtInh(string path, int sampleIndex = 0, LockMassParameters lockmassParameters = null, bool simAsSpectra = false,
            bool srmAsSpectra = false, bool acceptZeroLengthSpectra = true, bool requireVendorCentroidedMS1 = false, bool requireVendorCentroidedMS2 = false,
            bool ignoreZeroIntensityPoints = false) 
            : base(path, sampleIndex = 0, lockmassParameters = null, simAsSpectra = false, srmAsSpectra = false, acceptZeroLengthSpectra = true, 
             requireVendorCentroidedMS1 = false, requireVendorCentroidedMS2 = false, ignoreZeroIntensityPoints = false)
        { }
    }

    public class MsDataFileImplExtAgg
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

        public void test()
        {
            // _msDataFileImpl.GetSpectrumsInfo();
        }

    }
}
