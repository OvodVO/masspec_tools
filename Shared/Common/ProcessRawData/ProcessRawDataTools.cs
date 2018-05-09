using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using pwiz.ProteowizardWrapper;
using WashU.BatemanLab.MassSpec.Tools.TargetAnalysis;

namespace WashU.BatemanLab.MassSpec.Tools.ProcessRawData
        
{    
    public class ProcessRawDataTools
    {
        public static bool InMZTolerance(double mz, double expectedMz, double tolerance)
        {
            if (mz >= (expectedMz - tolerance) && mz <= (expectedMz + tolerance)) return true;
            else return false;
        }

        public static bool InMZTolerance(double mz, double[] expectedMzArray, double tolerance)
        {
            bool result = false;
            for (int i = 0; i <= expectedMzArray.Length; i++)
            {
                if (InMZTolerance(mz, expectedMzArray[i], tolerance)) result = true;
            }
            return result;
        }

        public static bool InMZTolerance(double mz, List<double> expectedMzList, double tolerance)
        {
            bool result = false;
            foreach (double _mz in expectedMzList) 
            {
                if (InMZTolerance(mz, _mz, tolerance)) result = true;
            }
            return result;
        }

        public static MzIntensityPair[] PairMzIntensity(double[] MZs, double[] Intensities)
        {
            var result = MZs.Zip(Intensities, (m, i) => new MzIntensityPair { mz = m,  intensity = i });

            return result.ToArray();
        }
        
        public static double[] AggIonCounts(double[] MZs, double[] Intensities, List<double> expectedMzList, double tolerance)
        {
            double[] result = new double[2];
            var groupresult = (from pair in PairMzIntensity(MZs, Intensities)
                               group pair by InMZTolerance(pair.mz, expectedMzList, tolerance) into gr
                               select new { IsTarget = gr.Key, TotIntensity = gr.Sum(i => i.intensity) });

            result[0] = groupresult.Where(w => w.IsTarget == true ).Sum(s => s.TotIntensity);
            result[1] = groupresult.Where(w => w.IsTarget == false).Sum(s => s.TotIntensity);

            return result;
        }
    }

    public class MzIntensityPair
    {
        public MzIntensityPair() { }
        public double intensity { get; set; }
        public double mz { get; set; }
        public MzIntensityPair(double Mz, double Intensity)
        {
            mz = Mz;
            intensity = Intensity;
        }
    }

    

   
}
