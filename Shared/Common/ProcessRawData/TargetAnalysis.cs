using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using pwiz.ProteowizardWrapper;

namespace WashU.BatemanLab.MassSpec.Tools.TargetAnalysis
{
    public class TargetAnalysis
    {
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

    public class Protein
    {
        public string Name;
        public List<Peptide> Peptides;
    }

    public class Peptide
    {
        public string Name;
        public string Sequence;
        public List<Precursor> Precursors { get; set; }
    }

    public class Precursor
    {
        public string IsotopeLabelType;
        public double PrecursorMZ { get; set; }
        public List<double> Products { get; set; }
    }
}
