using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using pwiz.ProteowizardWrapper;

namespace WashU.BatemanLab.MassSpec.Tools.TargetAnalysis
{
    class TargetAnalysis
    {
        public string TargetAnalysisType { get; set; }
        public List<Protein> Proteins { get; set; }
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
