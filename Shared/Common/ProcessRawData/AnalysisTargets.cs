using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using pwiz.ProteowizardWrapper;

namespace WashU.BatemanLab.MassSpec.Tools.Analysis
{
    [Serializable]
    public class AnalysisTargets
    {
        public static readonly double MS1Tolerance = 0.1;
        public string TargetAnalysisType { get; set; }
        public List<Protein> Proteins { get; set; }
        public static List<Protein> GetDefaultProteins()
        {
            var proteins = new List<Protein>()
            {
                new Protein()
                {
                    Name = "Ab38",
                    Peptides = new List<Peptide>()
                    {
                        new Peptide()
                        {
                            Name = "Ab38 Peptide",
                            Precursors = new List<Precursor>()
                            {
                                new Precursor()
                                {
                                    PrecursorMZ = 508.646041, IsotopeLabelType = "light",
                                    Products = new List<double>() { 653.434472, 784.474957, 883.543371 }
                                },
                                new Precursor()
                                {
                                    PrecursorMZ = 514.606495, IsotopeLabelType = "N15 ISTD",
                                    Products = new List<double>() { 661.410752, 793.448272, 893.513721 }
                                }
                            }
                        }
                    }
                },
                new Protein()
                {
                    Name = "Ab42",
                    Peptides = new List<Peptide>()
                    {
                        new Peptide()
                        {
                            Name = "Ab42 Peptide",
                            Precursors = new List<Precursor>()
                            {
                                new Precursor()
                                {
                                    PrecursorMZ = 699.896296, IsotopeLabelType = "light",
                                    Products = new List<double>() { 883.543371, 940.564835, 997.586298, 1096.654712, 1195.723126 }
                                },
                                new Precursor()
                                {
                                    PrecursorMZ = 707.843568, IsotopeLabelType = "N15 ISTD",
                                    Products = new List<double>() { 893.513721, 951.53222, 1009.550718, 1109.616167, 1209.681616 }
                                }
                            }
                        }
                    }
                }
            };
            return proteins;
        }
    }

    [Serializable]
    public class Protein
    {
        public string Name;
        public List<Peptide> Peptides;
        public Protein()
        {
            Peptides = new List<Peptide>();
        }
    }

    [Serializable]
    public class Peptide
    {
        public static string GetPeptideShortName(string peptide)
        {
            string result;
            switch (peptide)
            {
                case null: result = "";
                    break;
                case "KLVFFAEDVGSN": result = "Aβ[Total]";
                    break;
                case "KGAIIGLMVGG": result = "Aβ38";
                    break;
                case "KGAIIGLMVGGVV": result = "Aβ40";
                    break;
                case "KGAIIGLMVGGVVIA": result = "Aβ42";
                    break;
                default: result = peptide;
                    break;
            }
            return result;
        }

        public string Name;
        public string Sequence;
        public List<Precursor> Precursors { get; set; }
        public Peptide()
        {
            Precursors = new List<Precursor>();
        }
    }

    [Serializable]
    public class Precursor
    {
        public string IsotopeLabelType;
        public double PrecursorMZ { get; set; }
        public List<double> Products { get; set; }
        public List<double> ProductsM1
        { get
            {
                List<double> _M1productList = new List<double>();
                Products.ForEach((m) => { _M1productList.Add(m + 1.0); });
                return _M1productList;
            } 
        }
        public List<double> ProductsM2
        { get
            {
                List<double> _M2productList = new List<double>();
                Products.ForEach((m) => { _M2productList.Add(m + 2.0); });
                return _M2productList;
            } 
        }
        public List<double> ProductsM3
        {
            get
            {
                List<double> _M3productList = new List<double>();
                Products.ForEach((m) => { _M3productList.Add(m + 3.0); });
                return _M3productList;
            }
        }
        public List<double> ProductsMneg1
        {
            get
            {
                List<double> _Mn1productList = new List<double>();
                Products.ForEach((m) => { _Mn1productList.Add(m - 1.0); });
                return _Mn1productList;
            }
        }
        public Precursor()
        {
            Products = new List<double>();
        }
    }

    [Serializable]
    public class Chromatogramm
    {
        public string Protein { get; set; }
        public string Peptide { get; set; }
        public string IsotopeLabelType { get; set; }
        public double PrecursorMZ { get; set; }
        public double[] RetentionTimes { get; set; }

        public double[] IonInjectionTimes { get; set; }
        public double[] TotalIonCurrents { get; set; }
        public double[] SumOfIntensities { get; set; }
        public double[] SumOfPositiveMatch { get; set; }
        public double[] SumOfPositiveMatch_M0 { get; set; }
        public double[] SumOfPositiveMatch_M1 { get; set; }
        public double[] SumOfPositiveMatch_M2 { get; set; }
        public double[] SumOfPositiveMatch_M3 { get; set; }
        public double[] SumOfPositiveMatch_Mneg1 { get; set; }
        public double[] SumOfNegativeMatch { get; set; }

        public static List<string> MeasuresNameList
        {
            get
            {
                 return new List<string>
                          { "IIT", "TIC", "Sum of Intensities",
                            "M0", "M+1", "M+2", "M+3", "M-1",
                           "Total positive match", "Negative match/noise" };
            }
        }

        public static List<string> MeasuresNameListByDefault
        {
            get
            {
                return new List<string>
                          { "IIT", "M0", "Negative match/noise" };
            }
        }

        public static List<Color> MeasuresColorList
        {
            get
            {
                return new List<Color>
                          { Color.Black, Color.Chocolate, Color.Indigo,
                            Color.Blue, Color.DeepPink, Color.Silver, Color.Teal, Color.Gold,
                            Color.DarkGreen, Color.Red
                          };
            }
        }

        public static Dictionary<string, Color> MeasuresDictionary
        {
            get
            {
                return MeasuresNameList.Zip(MeasuresColorList, (n, k) => new { n, k })
                .ToDictionary(x => x.n, x => x.k);
            }
        }
        public static int MeasuresCount
        {
            get
            {
                return MeasuresNameList.Count;
            }
        }

        public double[] GetMeasureByName(String measure)
        {
            switch (measure)
            {
                case "IIT": return IonInjectionTimes; break;
                case "TIC": return TotalIonCurrents; break;
                case "Sum of Intensities": return SumOfIntensities; break;
                case "M0":  return SumOfPositiveMatch_M0; break;
                case "M+1": return SumOfPositiveMatch_M1; break;
                case "M+2": return SumOfPositiveMatch_M2; break;
                case "M+3": return SumOfPositiveMatch_M3; break;
                case "M-1": return SumOfPositiveMatch_Mneg1; break;
                case "Total positive match": return SumOfPositiveMatch; break;
                case "Negative match/noise": return SumOfNegativeMatch; break;
                default: return null;
                    break;
            }

        }
    }
}
