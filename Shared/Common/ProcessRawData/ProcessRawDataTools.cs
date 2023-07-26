using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using pwiz.ProteowizardWrapper;
using WashU.BatemanLab.MassSpec.Tools.Analysis;

using System.IO;

namespace WashU.BatemanLab.MassSpec.Tools.ProcessRawData
        
{    
    public class ProcessRawDataTools
    {
        public static bool InMZTolerance(double mz, double expectedMz, double tolerance)
        {
            if (mz >= (expectedMz - tolerance) && mz <= (expectedMz + tolerance)) return true;
            else return false;
        }

        public static bool InMZTolerance(double? mz, double expectedMz, double tolerance)
        {
            if (mz.HasValue) return InMZTolerance(mz.Value, expectedMz, tolerance);
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

        public static MzIntensityPair[] PairMzIntensityByLINQ(double[] MZs, double[] Intensities)
        {
            var result = MZs.Zip(Intensities, (m, i) => new MzIntensityPair { mz = m,  intensity = i });

            return result.ToArray();
        }

        public static MzIntensityPair[] PairMzIntensity(double[] MZs, double[] Intensities)
        {
            MzIntensityPair[] result = new MzIntensityPair[MZs.Length];
            for (int i = 0; i <= MZs.Length; i++)
            {
                result[i] = new MzIntensityPair() { mz = MZs[i], intensity = Intensities[i] };
            }
            return result;
        }

        public static double[] AggIonCountsByLINQ(double[] MZs, double[] Intensities, List<double> expectedMzList, double tolerance)
        {
            double[] result = new double[2];
            var groupresult = (from pair in PairMzIntensity(MZs, Intensities)
                               group pair by InMZTolerance(pair.mz, expectedMzList, tolerance) into gr
                               select new { IsTarget = gr.Key, TotIntensity = gr.Sum(i => i.intensity) });

            result[0] = groupresult.Where(w => w.IsTarget == true ).Sum(s => s.TotIntensity);
            result[1] = groupresult.Where(w => w.IsTarget == false).Sum(s => s.TotIntensity);

            return result;
        }

        public static double[] AggIonCounts(double[] MZs, double[] Intensities, double tolerance,
                                            List<double> expectedM0List,
                                            List<double> expectedM1List,
                                            List<double> expectedM2List,
                                            List<double> expectedM3List,
                                            List<double> expectedMneg1List)
        {
            double[] result = new double[7];
            double posMatchSum_M0 = 0, posMatchSum_M1 = 0, posMatchSum_M2 = 0, posMatchSum_M3 = 0, posMatchSum_Mneg1 = 0;
            double negMatchSum = 0;

            for (int i = 0; i < MZs.Length; i++)
            {
                if (InMZTolerance(MZs[i], expectedM0List, tolerance))
                {
                    posMatchSum_M0 += Intensities[i];
                }
                else if ((InMZTolerance(MZs[i], expectedM1List, tolerance)))
                {
                    posMatchSum_M1 += Intensities[i];
                }
                else if ((InMZTolerance(MZs[i], expectedM2List, tolerance)))
                {
                    posMatchSum_M2 += Intensities[i];
                }
                else if ((InMZTolerance(MZs[i], expectedM3List, tolerance)))
                {
                    posMatchSum_M3 += Intensities[i];
                }
                else if ((InMZTolerance(MZs[i], expectedMneg1List, tolerance)))
                {
                    posMatchSum_Mneg1 += Intensities[i];
                }
                else
                {
                   negMatchSum += Intensities[i];
                }
            }

            result[0] = negMatchSum; 
            result[1] = posMatchSum_M0;
            result[2] = posMatchSum_M1;
            result[3] = posMatchSum_M2;
            result[4] = posMatchSum_M3;
            result[5] = posMatchSum_Mneg1;
            result[6] = posMatchSum_M0 + posMatchSum_M1 + posMatchSum_M2 + posMatchSum_M3 + posMatchSum_Mneg1;

            return result;
        }

        public static double[] AggIonCountsOld(double[] MZs, double[] Intensities, List<double> expectedMzList, double tolerance)
        {
            double[] result = new double[2];
            double posMatchSum = 0;
            double negMatchSum = 0;

            // temp

            // List<String> _deb = new List<string>();

            for (int i = 0; i < MZs.Length; i++)
            {
                //   string debstr = "i=" + i.ToString() + "  MZs[i]=" + MZs[i].ToString();

                if (InMZTolerance(MZs[i], expectedMzList, tolerance))
                {
                    posMatchSum += Intensities[i];
                    //     debstr += "Intensities - " + Intensities[i].ToString() + " true; and posMatchSum=" + posMatchSum.ToString();
                }
                else
                {
                    negMatchSum += Intensities[i];
                    //     debstr += "Intensities - " + Intensities[i].ToString() +  " false; and negMatchSum=" + negMatchSum.ToString();
                }
                // _deb.Add(debstr);
            }

            result[0] = posMatchSum;
            result[1] = negMatchSum;

            //_deb.Add("Finally result[0] is - " + result[0].ToString());
            //_deb.Add("Finally result[1] is - " + result[1].ToString());

            //File.WriteAllLines(@"d:\_TEMP\2019-12-09\debug.txt", _deb.ToArray<String>());

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
