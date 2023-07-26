using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.Threading.Tasks;
using WashU.BatemanLab.Common;
using pwiz.ProteowizardWrapper;
using WashU.BatemanLab.MassSpec.Tools.ProcessRawData;


namespace WashU.BatemanLab.MassSpec.Tools.Analysis
{
    [Serializable]
    public class AnalysisResults
    {
        private List<MsDataFileImplExtAgg> _analysisResults;
        private AnalysisTargets _analysisTargets;
        public AnalysisTargets AnalysisTargets { get { return _analysisTargets; } set { _analysisTargets = value; } }
        public List<MsDataFileImplExtAgg> Results { get { return _analysisResults; }
        }

        public AnalysisResults()
        {
            _analysisResults = new List<MsDataFileImplExtAgg>();
            _analysisTargets = new AnalysisTargets();
        }
        public static string GetMSRunShorten(string msRunName, string format)
        {
            string result = "";
            char separator = '_';
            string[] msRunSplit = msRunName.Split('_');
            string[] formatSplit = format.Split(',');
            int count = 0;
            for (int i = 0; i < formatSplit.Length; i++)
            {
                int index;
                if (int.TryParse(formatSplit[i], out index))
                {
                    if (index < msRunSplit.Length)
                    {
                        if (count > 0 ) { result += separator; }
                        result += msRunSplit[index];
                        count++;
                    }
                };
            }
            return result.Replace(".raw", "");
        }

        public void SaveToXML(string filePath)
        {
            XmlSerialization.WriteToXmlFile<AnalysisResults>(filePath, this, false);
        }

        public void SaveToBin(string filePath)
        {
            XmlSerialization.WriteToBinaryFile<AnalysisResults>(filePath, this, false);
        }

        public static AnalysisResults OpenFromXMLFile(string filePath)
        {
            return XmlSerialization.ReadFromXmlFile<AnalysisResults>(filePath);
        }

        public static AnalysisResults OpenFromBinFile(string filePath)
        {
            return XmlSerialization.ReadFromBinaryFile<AnalysisResults>(filePath);
        }
        
        private Dictionary<double, double> AbetaITFcorrection()
        {
            var _Dictionary = new Dictionary<double, double>
            {
                {    891, 607.7778 },
                {  901.5, 614.7317 },
                { 1096.5, 699.8963 },
                { 1109.5, 707.8436 },
                {  768.5, 508.6460 },
                {  777.5, 514.6065 },
                { 1028.5, 663.7446 },
                { 1038.5, 670.6985 }
            };
            return _Dictionary;
        }

        public Task<bool> ReadAndAnalyzeMSFile(string path)
        {
            var IsCompleted = false;
            var c = _analysisResults.Count;
            var progress = new ReadAndAnalyzeProgressInfo();
            var msrun = new MsDataFileImplExtAgg(path);
             msrun.GetMsDataSpectrums();

            //msrun.TESTMS3();
            //msrun.ITFtoMS2(891, 607.7778, 0.1);

            //List<String> _deb = new List<string>();

            //_deb.Add("starting modi" + path + "at" + DateTime.Now.ToString());

           // msrun.CorrectForITF1(AbetaITFcorrection());

            //_deb.Add("finish modi" + path + "at" + DateTime.Now.ToString());

            //File.WriteAllLines(@"d:\_TEMP\2019-12-09\" + Path.GetFileNameWithoutExtension(path) + ".dbj", _deb.ToArray<String>());

            msrun.GetChromatograms(_analysisTargets.Proteins, 0.01, 0.1);
             _analysisResults.Add(msrun);
            if (_analysisResults.Count - c == 1)
                IsCompleted = true;
            return Task.FromResult(IsCompleted); 
        }
    }
}
