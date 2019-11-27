using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using Microsoft.Win32;
using System.Xml.Linq;
using System.Xml;
using Office = Microsoft.Office.Core;

using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace QLDReader
{
    public partial class fmMain : Form
    {
        public fmMain()
        {
            InitializeComponent();
            _OutputFormat = new OutputFormatLayout();
        }
        //private DBNull

        private Excel.Application _xlApp;

        private string[] _FileList;

        private string _workFolder;

        private string _CustomDateFormat;

        private class OutputFormatLayout
        {
            public byte _titleShift = 4;

            public char _colLet_First = 'A';
            public byte _colNum_First = 1;

            public char _colLet_Last = 'O';
            public byte _colNum_Last = 15;

            public char _colLet_sampleid = 'C';
            public byte _colNum_sampleid = 3;

            public char _colLet_stdconc = 'D';
            public byte _colNum_stdconc = 4;

            public char _colLet_response = 'I';
            public byte _colNum_response = 9;

            public char _colLet_analconc = 'J';
            public byte _colNum_analconc = 10;

            public char _colLet_analarea = 'E';
            public byte _colNum_analarea = 5;

            public char _colLet_analISarea = 'F';
            public byte _colNum_analISarea = 6;

            public char _colLet_stddev = 'K';
            public byte _colNum_stddev = 11;

            public char _colLet_subject = 'N';
            public byte _colNum_subject = 14;


        }

        private OutputFormatLayout _OutputFormat;

        private class ExperimentInfoType
        {
            public string _Date = "noDate";
            public byte _DateInx = 0;
            public string _Subject = "noSubjectNo";
            public byte _SubjectInx = 1;
            public string _matrix = "noMatrix";
            public byte _matrixID;
            public byte _matrixInx = 2;
            public string _IP = "noIP";
            public byte _AbodyID;
            public byte _IPInx = 3;
            public string _Enzyme = "noEnzyme";
            public byte _EnzymeID;
            public byte _EnzymeInx = 4;
            public string _Instument = "noInstument";
            public byte _InstrumentID;
            public byte _InstumentInx = 5;
            public string _quanType = "noQuanType";
            public byte _quanTypeInx = 6;
            public String _CommonChartNameStr;
            public string _ProcessedBy = "";
            public byte _SampleProcessBy;
            public byte _DoneBy;
            public byte _QuantitatedBy;
            public string _fileName = "";
            public string _expName = "";
            public void Update()
            {
                _CommonChartNameStr = String.Format
                            ("{0} {1} {2} {3} {4} ({5})",
                                                _Subject, _matrix, _IP, _Enzyme, _Instument, _Date);
            }
        }

        private ExperimentInfoType ExperimentInfo;

        private void getExperimentInfoType(string _inputFileName, bool _parsefileName)
        {
            if (_inputFileName == "") { return; }

            ExperimentInfo = new ExperimentInfoType();

            if (_inputFileName != null)
            {
                ExperimentInfo._fileName = _inputFileName;
                ExperimentInfo._expName = _inputFileName.Replace("_REL", "").Replace("_ABS", "");
            }

            string[] _parameters = _inputFileName.Split('_');

            if (_parsefileName)
            {
                //string[] _parameters = _inputFileName.Split('_');

                try
                {
                    dateTimePicker_AssayDate.Value = DateTime.ParseExact(_parameters[ExperimentInfo._DateInx],
                        _CustomDateFormat, null);

                    comboBox_Fluid.SelectedIndex = -1;
                    comboBox_Abody.SelectedIndex = -1;
                    comboBox_Enzyme.SelectedIndex = -1;
                    comboBox_Instrument.SelectedIndex = -1;
                    comboBox_QuantType.SelectedIndex = -1;

                    comboBox_Fluid.SelectedItem = _parameters[ExperimentInfo._matrixInx];
                    comboBox_Abody.SelectedItem = _parameters[ExperimentInfo._IPInx];
                    comboBox_Enzyme.SelectedItem = _parameters[ExperimentInfo._EnzymeInx];
                    comboBox_Instrument.SelectedItem = _parameters[ExperimentInfo._InstumentInx];

                    if (_parameters[ExperimentInfo._quanTypeInx].Contains("REL") || _parameters[ExperimentInfo._quanTypeInx].Contains("ABS"))
                    {
                        comboBox_QuantType.SelectedItem = _parameters[ExperimentInfo._quanTypeInx];
                    }
                    else
                    {
                        comboBox_QuantType.SelectedIndex = 0;
                    }

                    textBox_subject.Text = _parameters[ExperimentInfo._SubjectInx];
                }

                catch (FormatException _excep)
                {
                    MessageBox.Show(_excep.Message, "change custom date format");
                }

                catch (Exception _excep)
                {
                    MessageBox.Show(_excep.Message);
                }
            }

            try
            {
                dateTimePicker_AssayDate.Value = DateTime.ParseExact(_parameters[ExperimentInfo._DateInx],
                    _CustomDateFormat, null);
                if (Convert.ToByte(comboBox_Project.SelectedValue) == 11)
                {
                    textBox_subject.Text = _parameters[2];
                }
            }
            catch (Exception _exep)
            {
                MessageBox.Show(_exep.Message);
            }


            ExperimentInfo._Date = dateTimePicker_AssayDate.Value.ToShortDateString();
            ExperimentInfo._Subject = textBox_subject.Text;
            ExperimentInfo._matrix = comboBox_Fluid.Text;
            ExperimentInfo._matrixID = Convert.ToByte(comboBox_Fluid.SelectedValue);
            ExperimentInfo._Instument = comboBox_Instrument.Text;
            ExperimentInfo._IP = comboBox_Abody.Text;
            ExperimentInfo._Enzyme = comboBox_Enzyme.Text;
            ExperimentInfo._EnzymeID = Convert.ToByte(comboBox_Enzyme.SelectedValue);
            ExperimentInfo._quanType = comboBox_QuantType.Text;

            ExperimentInfo._CommonChartNameStr = String.Format
                    ("{0} {1} {2} {3} {4} ({5})",
                    ExperimentInfo._Subject,
                    ExperimentInfo._matrix,
                    ExperimentInfo._IP,
                    ExperimentInfo._Enzyme,
                    ExperimentInfo._Instument,
                    ExperimentInfo._Date);

            ExperimentInfo._SampleProcessBy = Convert.ToByte(comboBox_SampleProcessBy.SelectedValue);
            ExperimentInfo._DoneBy = Convert.ToByte(comboBox_DoneBy.SelectedValue);
            ExperimentInfo._QuantitatedBy = Convert.ToByte(comboBox_QuantitatedBy.SelectedValue);
        }

        private void getExperimentInfoType(string _inputFileName)
        {
            if (_inputFileName == "") { return; }

            string[] _parameters = _inputFileName.Split('_');

            ExperimentInfo = new ExperimentInfoType();
            ExperimentInfo._Date = _parameters[ExperimentInfo._DateInx];
            ExperimentInfo._Instument = _parameters[ExperimentInfo._InstumentInx];
            ExperimentInfo._matrix = _parameters[ExperimentInfo._matrixInx];
            ExperimentInfo._IP = _parameters[ExperimentInfo._IPInx];
            ExperimentInfo._Enzyme = _parameters[ExperimentInfo._EnzymeInx];
            ExperimentInfo._Subject = _parameters[ExperimentInfo._SubjectInx];
            ExperimentInfo._quanType = _parameters[ExperimentInfo._quanTypeInx];

            ExperimentInfo._CommonChartNameStr = String.Format
                    ("{0} {1} {2} {3} {4} ({5})",
                    ExperimentInfo._Subject,
                    ExperimentInfo._matrix,
                    ExperimentInfo._IP,
                    ExperimentInfo._Enzyme,
                    ExperimentInfo._Instument,
                    ExperimentInfo._Date);
            ExperimentInfo._fileName = _inputFileName;
            ExperimentInfo._expName = _inputFileName.Replace("_REL", "").Replace("_ABS", "");

            ExperimentInfo._matrixID = Convert.ToByte(comboBox_Fluid.SelectedValue);
            ExperimentInfo._SampleProcessBy = Convert.ToByte(comboBox_SampleProcessBy.SelectedValue);
            ExperimentInfo._DoneBy = Convert.ToByte(comboBox_DoneBy.SelectedValue);
            ExperimentInfo._QuantitatedBy = Convert.ToByte(comboBox_QuantitatedBy.SelectedValue);

        }

        private void ToolStripMenuItem_About_Click(object sender, EventArgs e)
        {
            fmAbout _fmAbout = new fmAbout();
            _fmAbout.ShowDialog();
        }

        private void ToolStripMenuItem_Exit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button_SelectQLD_Click(object sender, EventArgs e)
        {
            string _WorkFolder = @"c:\";

            if (_workFolder == null) { _workFolder = _WorkFolder; }

            OpenFileDialog _openDlg = new OpenFileDialog();
            _openDlg.Filter = "XML|*.XML";
            _openDlg.Multiselect = true;
            _openDlg.InitialDirectory = _workFolder;

            if (_openDlg.ShowDialog() == DialogResult.OK)
            {
                _workFolder = Path.GetDirectoryName(_openDlg.FileName);

                if (checkBox_fileName.Checked == true)
                {
                    foreach (ComboBox _com in this.groupBox_ExperInfo.Controls.OfType<ComboBox>())
                    {
                        if (_com.DisplayMember != "LAB_MEMBERS_LOGIN")
                        {
                            _com.Text = "";
                        }
                    }
                    textBox_subject.Text = "";
                }

                _FileList = _openDlg.FileNames;

                textBox_QLDfile.Text = _openDlg.FileName;
                textBox_Excelfile.Text = Path.ChangeExtension(textBox_QLDfile.Text, "xlsx");

                if (_openDlg.FileNames.Count() > 1)
                {
                    this.button_StartExport_Batch.Enabled = true;
                    _FileList = _openDlg.FileNames;
                }
                else
                {
                    this.button_StartExport_Batch.Enabled = false;
                }
            }

            getExperimentInfoType(Path.GetFileNameWithoutExtension(textBox_QLDfile.Text), checkBox_fileName.Checked);

            groupBox_ExperInfo.Visible = true;



        }

        private void checkBox_CustomDate_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_CustomDate.Checked)
            {
                textBox_date_format.Visible = true;
            }
            else
            {
                textBox_date_format.Visible = false;
            }
        }

        private void textBox_date_format_TextChanged(object sender, EventArgs e)
        {
            this._CustomDateFormat = textBox_date_format.Text;
        }

        private void fmMain_Shown(object sender, EventArgs e)
        {
            RegistryKey _key;
            _key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\QLDReader");

            if (_key == null) { return; };

            try
            {
                foreach (CheckBox _chbox in this.Controls.OfType<CheckBox>())
                {
                    _chbox.Checked = Convert.ToBoolean(_key.GetValue(_chbox.Name).ToString());
                }

                textBox_date_format.Text = _key.GetValue(textBox_date_format.Name).ToString();

                foreach (ComboBox _combo in this.groupBox_ExperInfo.Controls.OfType<ComboBox>())
                {
                    _combo.Text = _key.GetValue(_combo.Name).ToString();
                }

                foreach (NumericUpDown _numer in this.groupBox_ExperInfo.Controls.OfType<NumericUpDown>())
                {
                    _numer.Value = Convert.ToInt32(_key.GetValue(_numer.Name));
                }

                _workFolder = _key.GetValue("last_folder").ToString();


                foreach (CheckBox _chk in this.Controls.OfType<CheckBox>())
                {
                    _chk.Checked = Convert.ToBoolean(_key.GetValue(_chk.Name));
                }

            }
            catch (Exception _ex) { MessageBox.Show(_ex.TargetSite.ToString()); }

            //****************
            _CustomDateFormat = textBox_date_format.Text;
        }


        private double AnalConc_EmptyCheck(string _analconc)
        {
            double _dblanalConc = 0.0;

            try
            {
                _dblanalConc = Convert.ToDouble(_analconc);
            }
            catch { }

            return _dblanalConc;
        }

        private int QuntTypeToDBInt(string _Analyte)
        {
            int _intQuanType = 0;
            if (_Analyte.Contains("C13N14")) { _intQuanType = 1; }
            if (_Analyte.Contains("C12N14")) { _intQuanType = 2; }
            return _intQuanType;
        }

        private int FluidTypeToDBInt(string _FluidType)
        {
            int _intFluidType = 0;
            if (_FluidType.Contains("hCSF")) { _intFluidType = 1; }
            if (_FluidType.Contains("Plas")) { _intFluidType = 2; }
            if (_FluidType.Contains("hPlasma")) { _intFluidType = 2; }

            return _intFluidType;
        }

        private int SampleTypeToDBInt(string _SampleType)
        {
            int _intSampleType = 0;
            if (_SampleType.Contains("Standard")) { _intSampleType = 5; }
            if (_SampleType.Contains("Analyte")) { _intSampleType = 0; }

            return _intSampleType;
        }

        private int EnzymeToDBInt(string _Enzyme)
        {
            int _intEnzyme = 0;
            if (_Enzyme.Contains("Tryp")) { _intEnzyme = 1; }
            if (_Enzyme.Contains("Lys-N")) { _intEnzyme = 2; }
            if (_Enzyme.Contains("LysN")) { _intEnzyme = 2; }
            if (_Enzyme.Contains("LysC") || _Enzyme.Contains("Lys-C")) { _intEnzyme = 3; }

            return _intEnzyme;
        }

        private int IPToDBInt(string _IP)
        {
            int _intIP = 0;
            if (_IP.Contains("21F12")) { _intIP = 1; }
            if (_IP.Contains("2G3")) { _intIP = 2; }
            if (_IP.Contains("HJ5")) { _intIP = 3; }
            if (_IP.Contains("M266")) { _intIP = 4; }
            return _intIP;
        }

        private int InjectionToDBInt(string _Injection)
        {
            int _intInjection = 0;
            if (_Injection.Contains("INJ. B")) { _intInjection = 1; }
            if (_Injection.Contains("INJ. C")) { _intInjection = 2; }

            return _intInjection;
        }

        private double GetSlopeOfEquation(string _equation)
        {

            return Convert.ToDouble(_equation.Substring(0, _equation.IndexOf('*')));
        }

        private double GetInterOfEquation(string _equation)
        {

            return Convert.ToDouble(_equation.Substring(_equation.LastIndexOf('+') + 1));
        }

        private double GetLevel_TP(string _sampleID, string _sampleType)
        {
            double _doubleTP = 0;

            try
            {
                if (_sampleType == "Standard")
                {
                    _doubleTP = Convert.ToDouble(_sampleID.Substring(_sampleID.LastIndexOf('#') + 1));
                }
                else
                {
                    if (_sampleType == "Analyte")
                    {
                        _doubleTP = Convert.ToDouble(_sampleID.Replace('^', '.'));


                        var query_TP_ID = tIME_POINTTableAdapter.GetData().Select("TIME_POINT_VALUE = "
                                                                      + _doubleTP.ToString()
                                                                      + "AND STUDY_ID = "
                                                                      + comboBox_Study.SelectedValue);

                        // MessageBox.Show(query_TP_ID[0]["TIME_POINT_ID"].ToString());

                        _doubleTP = Convert.ToDouble(query_TP_ID[0]["TIME_POINT_ID"]);
                    }


                }

            }
            catch (Exception _err)
            {
                MessageBox.Show(_err.Message + ": " + _sampleID, "Can't convert to Integer");
            }

            return _doubleTP;
        }

        private void ExportLOAD_plasma()
        {
            
            List<string> _strExportSQL = new List<string>();

            XDocument _QLDDoc;
            try
            {
                _QLDDoc = XDocument.Load(textBox_QLDfile.Text);
            }
            catch (Exception _XMLExeption)
            {
                MessageBox.Show(_XMLExeption.Message);
                return;
            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = true;
            toolStripStatusLabelMain.Visible = true;

            _xlApp = new Excel.Application();
            _xlApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;

            Excel.Workbook xlWB;

            try
            {
                xlWB =
                    _xlApp.Workbooks.Add(
                                         Application.StartupPath + Path.DirectorySeparatorChar +
                                         @"\Templates\FACS_Xevo_Template_Summary.xlsx"
                );
            }

            catch (Exception _exception)
            {
                MessageBox.Show("File: " + Application.StartupPath + Path.DirectorySeparatorChar +
                @"\Templates\FACS_Xevo_Template_Summary.xlsx", _exception.Message, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                
                return;
            }

            finally
            { }

            //MessageBox.Show(ExperimentInfo._quanType);

           

            string _quanTypeKey = "";

            if (ExperimentInfo._quanType == "RELATIVE")
            {
                _quanTypeKey = "C13N14";
            }
            if (ExperimentInfo._quanType == "ABSOLUTE")
            {
                _quanTypeKey = "C12N14";
            }
            if (ExperimentInfo._quanType == "REL+ABS")
            {
                _quanTypeKey = "N14";
            }

            

            Excel.Worksheet _wsSummary = (Excel.Worksheet)xlWB.ActiveSheet;
            

            Excel.Worksheet _wsCurrent = null;
                
            //new Microsoft.Office.Interop.Excel.Worksheet();


            var queryCOMPOUND_LIST = from compound in _QLDDoc.Descendants("CALIBRATIONDATA").Descendants("COMPOUND")
                                     where compound.Attribute("name").Value.Contains(_quanTypeKey)
                                           &&
                                           !compound.Attribute("name").Value.Contains("Ch")

                                           // select only group 1   //2013-05-03 VO
                                           && compound.Parent.Parent.Attribute("id").Value == "1"

                                     select compound;

            foreach (XElement XEcompound in queryCOMPOUND_LIST)
            {

                //MessageBox.Show(XEcompound.Attribute("name").Value, "Compound");

                _wsCurrent = (Excel.Worksheet)xlWB.Sheets.Add
                          (_wsSummary, Type.Missing, Type.Missing,
                          Application.StartupPath + Path.DirectorySeparatorChar +
                            @"\Templates\FACS_Xevo_Template.xlsx");
                string _sheetName = XEcompound.Attribute("name").Value;

                _wsCurrent.Cells[2, 1] = _sheetName;

                if (_sheetName.Length > 31)
                {
                    _sheetName = "alt_" + _sheetName.Remove(1, _sheetName.Length - 31 + 4);
                }


                _wsCurrent.Name = _sheetName.Replace("+", "_");

                //_wsCurrent.Cells[2, 1] = XEcompound.Attribute("name").Value;






                var XECurve = XEcompound.Element("CURVE");
                _wsCurrent.Cells[2, 3] = XECurve.Attribute("type").Value;
                _wsCurrent.Cells[2, 6] = XECurve.Attribute("origin").Value;
                _wsCurrent.Cells[2, 4] = XECurve.Attribute("weighting").Value;
                _wsCurrent.Cells[2, 8] = XECurve.Attribute("axistrans").Value;


                try
                {
                    _wsCurrent.Cells[2, 8] = XECurve.Element("CALIBRATIONCURVE").Attribute("curve").Value;
                    _wsCurrent.Cells[2, 11] = XECurve.Element("CORRELATION").Attribute("r").Value;
                    _wsCurrent.Cells[2, 13] = XECurve.Element("DETERMINATION").Attribute("rsquared").Value;
                }
                catch (Exception)
                {
                    if (checkBox_STD.Checked)
                    {
                        MessageBox.Show(String.Format("Calibration curve parameters for {0} unavailable", XEcompound.Attribute("name").Value));
                    }

                }

                                

                var querySAMPLE_LIST = from sample in _QLDDoc.Descendants("SAMPLE")
                                       // orderby sample.Attribute("type").Value descending, sample.Attribute("id").Value
                                       where (sample.Attribute("type").Value == "Standard" ||
                                              sample.Attribute("type").Value == "Analyte" ||
                                              sample.Attribute("type").Value == "QC")
                                       && sample.Attribute("name").Value != ""
                                       orderby sample.Attribute("type").Value.Length descending
                                       select sample;


                int _sampleCount = querySAMPLE_LIST.Count();
                int _standardCount = querySAMPLE_LIST.Count(p => p.Attribute("type").Value == "Standard");
                int _analyteCount = querySAMPLE_LIST.Count(p => p.Attribute("type").Value == "Analyte");
                int _qcCount = querySAMPLE_LIST.Count(p => p.Attribute("type").Value == "QC");


                double _maxTP = (from _sample in querySAMPLE_LIST select _sample).Count();

                try
                {
                    _maxTP = (from _sample in querySAMPLE_LIST
                              where _sample.Attribute("type").Value == "Analyte"
                                               && double.TryParse(_sample.Attribute("sampleid").Value, out _maxTP)
                              select Convert.ToDouble(_sample.Attribute("sampleid").Value)).Max();
                }
                catch
                {
                }



                //MessageBox.Show(test.ToString());

                int index = _OutputFormat._titleShift + 1;
                foreach (XElement XEsample in querySAMPLE_LIST)
                {
                    _wsCurrent.Cells[index, 1] = XEsample.Attribute("name").Value;
                    _wsCurrent.Cells[index, 2] = XEsample.Attribute("type").Value;
                    _wsCurrent.Cells[index, 3] = XEsample.Attribute("sampleid").Value;
                    _wsCurrent.Cells[index, 14] = XEsample.Attribute("subjecttext").Value;

                    _wsCurrent.Cells[index, 15] = XEsample.Parent.Parent.Attribute("name").Value;

                    // _wsCurrent.Cells[index, 20] = XEsample.Parent.Parent.Attribute("name").Value;


                    var queryCOMPOUND = from compound in XEsample.Descendants("COMPOUND")
                                        where compound.Attribute("name").Value == XEcompound.Attribute("name").Value
                                        select compound;

                    _wsCurrent.Cells[index, 4] = queryCOMPOUND.Single().Attribute("stdconc").Value;

                    var XEpeak = queryCOMPOUND.First().Element("PEAK");
                    var XEpeakis = XEpeak.Element("ISPEAK");


                    /*
                    if (XEsample.Attribute("id").Value == "34")
                    {
                        XEpeak = XEsample.Element("COMPOUND").Element("PEAK");
                        foreach (XAttribute _at in XEpeak.Attributes() )
                        {
                          MessageBox.Show
                              ( XEcompound.Attribute("name").Value.ToString()+ "  " 
                                 + _at.Name+ "  -   "+ _at.Value.ToString());
                        }
                        //MessageBox.Show(XEpeak. .Element .Attribute("area").Value.ToString());
                    }
                    */

                    if (!(XEpeak.Attribute("pkflags").Value.ToString() == ""))
                    {
                        _wsCurrent.Cells[index, 5] = XEpeak.Attribute("area").Value;
                        _wsCurrent.Cells[index, 6] = XEpeakis.Attribute("area").Value;
                        _wsCurrent.Cells[index, 7] = XEpeak.Attribute("foundrt").Value;
                        _wsCurrent.Cells[index, 8] = XEpeakis.Attribute("foundrt").Value;
                        _wsCurrent.Cells[index, 9] = XEpeak.Attribute("response").Value;

                        // if (!XEpeak.Attribute("analconc").Value.StartsWith("0.00000"))
                        // {
                        _wsCurrent.Cells[index, _OutputFormat._colNum_analconc] =
                            XEpeak.Attribute("analconc").Value;
                        // }

                        _wsCurrent.Cells[index, 11] = XEpeak.Attribute("concdevperc").Value;
                        _wsCurrent.Cells[index, 12] = XEpeak.Attribute("signoise").Value;
                        _wsCurrent.Cells[index, 13] = XEpeak.Attribute("chromnoise").Value;


                    }



                    //Export into DB

                    if (checkBox_ExportIntoDB.Checked)
                    {
                        if (XEpeak.Attribute("pkflags").Value.ToString() == "" || XEsample.Attribute("type").Value == "QC")
                        {
                            // MessageBox.Show(String.Format("Sample {1} on analyte {0} show no peak",
                            //                    XEcompound.Attribute("name").Value,
                            //                    XEsample.Attribute("name").Value));
                        }
                        else
                        {
                            _strExportSQL.Add("INSERT INTO TMP_XEVO_IMPORT (" +
                                                  " EXPER_FILE, ASSAY_DATE, ASSAY_TYPE_ID," +
                                                  " ANTIBODY_ID, ENZYME_ID, QUAN_TYPE_ID," +
                                                  " SAMPLE_PROCESS_BY, ASSAY_DONE_BY, DATA_PROCESS_BY," +
                                                  " FILENAME_DESC, SAMPLE_TYPE_ID, LEVEL_TP_ID," +
                                                  " NUM_SPEC_CONC, SUBJECT_NUM, FLUID_TYPE_ID," +
                                                  " ANYLYTE_NAME, AREA, HEIGHT," +
                                                  " RT, ISTD_AREA, ISTD_HEIGHT," +
                                                  " ISTD_RT, RESPONSE, CURVE," +
                                                  " WEIGHTING, ORIGIN, EQUATION," +
                                                  " R, R_SQR, NUM_SLOPE," +
                                                  " NUM_INTERSECT, NUM_ANAL_CONC, CONC_DEV_PERC," +
                                                  " SIGNOISE, CHROMNOISE, INJECTION_ID )");


                            _strExportSQL.Add(
                                    String.Format("VALUES " +
                                                     "( '{0}', '{1}', {2}, " +
                                                     "{3}, {4}, {5}, " +
                                                     "{6}, {7}, {8}, " +
                                                     "'{9}', {10}, {11}, " +
                                                     "{12}, '{13}', '{14}', " +
                                                     "'{15}', {16}, {17}, " +
                                                     "{18}, {19}, {20}, " +
                                                     "{21}, {22}, '{23}', " +
                                                     "'{24}', '{25}', '{26}', " +
                                                     "{27}, {28}, {29}, " +
                                                     "{30}, {31}, {32}, {33}, " +
                                                     "{34}, '{35}' );"
                                    ,
                                    ExperimentInfo._expName,                          // 0*EXPER_FILE
                                    ExperimentInfo._Date,                             // 1*ASSAY_DATE
                                    7,                                                // 2*ASSAY_TYPE_ID

                                    Convert.ToByte(comboBox_Abody.SelectedValue),
                                //IPToDBInt(ExperimentInfo._IP),                    // 3*ANTIBODY_ID  
                                    Convert.ToByte(comboBox_Enzyme.SelectedValue),
                                //EnzymeToDBInt(ExperimentInfo._Enzyme),            // 4*ENZYME_ID
                                    Convert.ToByte(comboBox_QuantType.SelectedValue),
                                //QuntTypeToDBInt(ExperimentInfo._quanType),        // 5*INTEG_TYPE_ID

                                    ExperimentInfo._SampleProcessBy,                  // 6*SAMPLE_PROCESS_BY
                                    ExperimentInfo._DoneBy,                           // 7*ASSAY_DONE_BY
                                    ExperimentInfo._QuantitatedBy,                    // 8*DATA_PROCESS_BY

                                    XEsample.Attribute("name").Value,                 // 9*FILENAME_DESC
                                    SampleTypeToDBInt
                                    (XEsample.Attribute("type").Value),               //10*SAMPLE_TYPE_ID
                                    GetLevel_TP(
                                    XEsample.Attribute("sampleid").Value,
                                        XEsample.Attribute("type").Value),             //11*LEVEL_TP_ID

                                    queryCOMPOUND.Single().Attribute("stdconc").Value, //12*NUM_SPEC_CONC
                                    XEsample.Attribute("subjecttext").Value,           //13*SUBJECT_NUM     
                                    ExperimentInfo._matrixID,          //14*FLUID_TYPE_ID

                                    XEcompound.Attribute("name").Value,                //15*ANALYTE_NAME
                                    XEpeak.Attribute("area").Value,                    //16*AREA
                                    XEpeak.Attribute("height").Value,                  //17*HEIGHT

                                    XEpeak.Attribute("foundrt").Value,                 //18*RT
                                    XEpeakis.Attribute("area").Value,                  //19*ISTD AREA
                                    XEpeakis.Attribute("height").Value,                //20*ISTD HEIGHT

                                    XEpeakis.Attribute("foundrt").Value,               //21*ISTD RT
                                    XEpeak.Attribute("response").Value,                //22*RESPONSE

                                    XECurve.Attribute("type").Value,                   //23*CURVE
                                    XECurve.Attribute("weighting").Value,              //24*WEIGHTING
                                    XECurve.Attribute("origin").Value,                 //25*ORIGIN
                                    XECurve.Element("CALIBRATIONCURVE")
                                        .Attribute("curve").Value,                     //26*EQUATION
                                    XECurve.Element("CORRELATION")
                                        .Attribute("r").Value,                         //27*R
                                    XECurve.Element("DETERMINATION")
                                        .Attribute("rsquared").Value,                  //28*R_SQR
                                    GetSlopeOfEquation(
                                    XECurve.Element("CALIBRATIONCURVE")
                                        .Attribute("curve").Value),                    //29*NUM_SLOPE
                                    GetInterOfEquation(
                                    XECurve.Element("CALIBRATIONCURVE")
                                        .Attribute("curve").Value),                    //30*NUM_INTERSECT
                                    AnalConc_EmptyCheck(
                                             XEpeak.Attribute("analconc").Value),      //31*NUM_ANAL_CONC
                                    XEpeak.Attribute("concdevperc").Value,             //32*CONC_DEV_PERC
                                    (XEpeak.Attribute("signoise").Value == "") ? "NULL" : XEpeak.Attribute("signoise").Value,    //33*SIGNOISE
                                    XEpeak.Attribute("chromnoise").Value,       //34*CHROMNOISE
                                    InjectionToDBInt(XEsample.Parent.Parent.Attribute("name").Value))
                                    );
                        }
                    }

                    // ProgressBar Begin
                    toolStripProgressBarMain.PerformStep();
                    toolStripStatusLabelMain.Text = "Proccesing ... " +
                        XEcompound.Attribute("name").Value + " - " + XEsample.Attribute("sampleid").Value;
                    toolStripStatusLabelMain.PerformClick();
                    // ProgressBar End

                    index++;

                }


                Excel.Range _xlRange_Std_Conc, _xlRange_Std_Response, _xlRange_Std_Dev, _xlRange_Std_AnalConc,
                            _xlRange_Analyte_SampleID, _xlRange_Analyte_AnalConc, _xlRange_Analyte_ISArea,
                            _xlRange_Std_Box, _xlRange_Analyte_Box, _xlRange_Sample_List_Box,
                            _xlRange_Tableau;

                //Standard Ranges****
                _xlRange_Std_Conc = _wsCurrent.get_Range
                (_OutputFormat._colLet_stdconc + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_stdconc + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Conc.Name = _wsCurrent.Name + "_Std_Conc";

                _xlRange_Std_Response = _wsCurrent.get_Range
                (_OutputFormat._colLet_response + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_response + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Response.Name = _wsCurrent.Name + "_Std_Response";

                _xlRange_Std_AnalConc = _wsCurrent.get_Range
                   (_OutputFormat._colLet_analconc + (_OutputFormat._titleShift + 1).ToString(),
                    _OutputFormat._colLet_analconc + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_AnalConc.Name = _wsCurrent.Name + "_Std_AnalConc";

                _xlRange_Std_Dev = _wsCurrent.get_Range
                (_OutputFormat._colLet_stddev + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_stddev + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Dev.Name = _wsCurrent.Name + "_Std_Deviation";

                _xlRange_Std_Box = _wsCurrent.get_Range
                (_OutputFormat._colLet_First + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Box.Name = _wsCurrent.Name + "_Std_Box";



                //Analyte Ranges****
                _xlRange_Analyte_SampleID = _wsCurrent.get_Range
             (_OutputFormat._colLet_sampleid + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
              _OutputFormat._colLet_sampleid + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_SampleID.Name = _wsCurrent.Name + "_Analyte_SampleID";

                _xlRange_Analyte_AnalConc = _wsCurrent.get_Range
             (_OutputFormat._colLet_analconc + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
              _OutputFormat._colLet_analconc + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_AnalConc.Name = _wsCurrent.Name + "_Analyte_AnalConc";

                _xlRange_Analyte_ISArea = _wsCurrent.get_Range
              (_OutputFormat._colLet_analISarea + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
               _OutputFormat._colLet_analISarea + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_ISArea.Name = _wsCurrent.Name + "_Analyte_ISArea";

                _xlRange_Analyte_Box = _wsCurrent.get_Range
             (_OutputFormat._colLet_First + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
              _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_Box.Name = _wsCurrent.Name + "_Analyte_Box";


                //SAMPLE_LIST Range
                _xlRange_Sample_List_Box = _wsCurrent.get_Range
              (_OutputFormat._colLet_First + (_OutputFormat._titleShift + 1).ToString(),
            _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Sample_List_Box.Name = _wsCurrent.Name + "_Sample_Box";

                //Tableau Range
                _xlRange_Tableau = _wsCurrent.get_Range
              (_OutputFormat._colLet_First + (_OutputFormat._titleShift).ToString(),
            _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Tableau.Name = _wsCurrent.Name + "_Tableau";



                //Format Ranges**** 
                _xlRange_Sample_List_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                _xlRange_Sample_List_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).Weight = 3.75;
                _xlRange_Sample_List_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).Weight = 3.75;
                _xlRange_Std_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                _xlRange_Std_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).Weight = 1.75;
                _xlRange_Analyte_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).Weight = 1.75;

                _xlRange_Std_Box.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);

                //////////////////////////////////////////////////////
                Excel.ColorScale cfColorScale = (Excel.ColorScale)(_xlRange_Analyte_ISArea.FormatConditions.AddColorScale(3));
                // Set the minimum threshold to red (0x000000FF) and maximum threshold
                // to blue (0x00FF0000).
                cfColorScale.ColorScaleCriteria[1].FormatColor.Color = 7039480;  //0x000000FF;  
                cfColorScale.ColorScaleCriteria[2].FormatColor.Color = 8711167;  // 0x00FF0000;
                cfColorScale.ColorScaleCriteria[3].FormatColor.Color = 13011546; // 0x00FF0000;
                /////////////////////////////////////////////////////



                //Plots****
                //Time Serie****

                double _rowHeight = (double)_wsCurrent.get_Range("A1", "A1").RowHeight;

                double _graphOffset = (double)_xlRange_Sample_List_Box.Top + (double)_xlRange_Sample_List_Box.Height +
                       (_rowHeight + _rowHeight * _qcCount);

                Excel.ChartObjects _xlCharts = (Excel.ChartObjects)_wsCurrent.ChartObjects(Type.Missing);
                ///*********************************************************************************************************************
                Excel.ChartObject _xlChartCalcConc = (Excel.ChartObject)_xlCharts.Add(
                                  0, _graphOffset, 950, 250);
                Excel.Chart _xlChartCalcConcPage = _xlChartCalcConc.Chart;

                _xlChartCalcConcPage.ChartType = Excel.XlChartType.xlColumnClustered;
                _xlChartCalcConcPage.HasTitle = true;
                _xlChartCalcConcPage.ChartTitle.Font.Size = 11;

                _xlChartCalcConcPage.ChartTitle.Caption = XEcompound.Attribute("name").Value + " " + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_TP = (Excel.Axes)_xlChartCalcConcPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_TP.HasMajorGridlines = true;
                _xlAxisX_TP.HasTitle = true;
                _xlAxisY_TP.HasTitle = true;

                //_xlAxisX_TP.MinimumScale = 0;
                //_xlAxisX_TP.MaximumScale = _maxTP;

                _xlAxisY_TP.MinimumScale = 0;

                _xlChartCalcConcPage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartCalcConcPage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_TimeCourse = (Excel.SeriesCollection)_xlChartCalcConcPage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_TTR = (Excel.Series)_xlSeriesColl_TimeCourse.NewSeries();
                _xlSeries_TTR.XValues = _xlRange_Analyte_SampleID.get_Offset(0, 11).Cells;
                _xlSeries_TTR.Values = _xlRange_Analyte_AnalConc;


                string _serieCurveName = "H:L Abeta";

                string _axeCurveXNumFormat = "0",
                       _axeCurveYNumFormat = "0.0%";

                string _axeCurveXCaption = "Subject #",
                       _axeCurveYCaption = "H:L ratio";

                if (ExperimentInfo._quanType == "ABS")
                {
                    _axeCurveYNumFormat = "0.0";
                    _axeCurveYCaption = "ng/mL, Concentration";
                    _serieCurveName = "Abeta level";
                }

                if (_standardCount == 0)
                {
                    _axeCurveYCaption = "Area Ratio";
                }

                //_xlAxisX_TP.TickLabels.NumberFormat = _axeCurveXNumFormat;
                _xlAxisY_TP.TickLabels.NumberFormat = _axeCurveYNumFormat;
                _xlAxisX_TP.AxisTitle.Caption = _axeCurveXCaption;
                _xlAxisY_TP.AxisTitle.Caption = _axeCurveYCaption;

                _xlSeries_TTR.Name = _serieCurveName;
                _xlSeries_TTR.Name = XEcompound.Attribute("name").Value;



                ///*********************************************************************************************************************

                Excel.ChartObject _xlChartISTD_Area = (Excel.ChartObject)_xlCharts.Add(
                                  0, _graphOffset + _xlChartCalcConc.Height, 950, 250);
                Excel.Chart _xlChartISTD_AreaPage = _xlChartISTD_Area.Chart;

                _xlChartISTD_AreaPage.ChartType = Excel.XlChartType.xlColumnClustered;
                _xlChartISTD_AreaPage.HasTitle = true;
                _xlChartISTD_AreaPage.ChartTitle.Font.Size = 11;

                _xlChartISTD_AreaPage.ChartTitle.Caption = XEcompound.Attribute("name").Value + " " + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_ISTD_Area = (Excel.Axes)_xlChartISTD_AreaPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_ISTD_Area = _xlAxes_ISTD_Area.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_ISTD_Area = _xlAxes_ISTD_Area.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_ISTD_Area.HasMajorGridlines = true;
                _xlAxisX_ISTD_Area.HasTitle = true;
                _xlAxisY_ISTD_Area.HasTitle = true;

                //_xlAxisX_TP.MinimumScale = 0;
                //_xlAxisX_TP.MaximumScale = _maxTP;

                _xlAxisY_ISTD_Area.MinimumScale = 0;

                _xlChartISTD_AreaPage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartISTD_AreaPage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_ISTD_Area = (Excel.SeriesCollection)_xlChartISTD_AreaPage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_ISTD_Area = (Excel.Series)_xlSeriesColl_ISTD_Area.NewSeries();
                _xlSeries_ISTD_Area.XValues = _xlRange_Analyte_SampleID.get_Offset(0, 11).Cells;
                _xlSeries_ISTD_Area.Values = _xlRange_Analyte_AnalConc.get_Offset(0, -4).Cells;


                //string _serieCurveName = "H:L Abeta";

                //string _axeCurveXNumFormat = "0",
                _axeCurveYNumFormat = "0.0E+0";

                //string _axeCurveXCaption = "Subject #",
                _axeCurveYCaption = "ISTD Area";

                if (ExperimentInfo._quanType == "ABS")
                {
                    _axeCurveYNumFormat = "0.0";
                    _axeCurveYCaption = "ng/mL, Concentration";
                    _serieCurveName = "Abeta level";
                }

                if (_standardCount == 0)
                {
                    _axeCurveYCaption = "Area Ratio";
                }

                //_xlAxisX_TP.TickLabels.NumberFormat = _axeCurveXNumFormat;
                _xlAxisY_ISTD_Area.TickLabels.NumberFormat = _axeCurveYNumFormat;
                _xlAxisX_ISTD_Area.AxisTitle.Caption = _axeCurveXCaption;
                _xlAxisY_ISTD_Area.AxisTitle.Caption = _axeCurveYCaption;

                _xlSeries_ISTD_Area.Name = _serieCurveName;
                _xlSeries_ISTD_Area.Name = XEcompound.Attribute("name").Value;


                ///*********************************************************************************************************************

                ///*********************************************************************************************************************
                //RT 
                Excel.ChartObject _xlChartRT = (Excel.ChartObject)_xlCharts.Add(
                                  0, _graphOffset + _xlChartCalcConc.Height + _xlChartISTD_Area.Height, 950, 250);
                Excel.Chart _xlChartRTPage = _xlChartRT.Chart;

                _xlChartRTPage.ChartType = Excel.XlChartType.xlLineMarkers;
                _xlChartRTPage.HasTitle = true;
                _xlChartRTPage.ChartTitle.Font.Size = 11;

                _xlChartRTPage.ChartTitle.Caption = XEcompound.Attribute("name").Value + " " + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_RT = (Excel.Axes)_xlChartRTPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_RT = _xlAxes_RT.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_RT = _xlAxes_RT.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_RT.HasMajorGridlines = true;
                _xlAxisX_RT.HasTitle = true;
                _xlAxisY_RT.HasTitle = true;

                //_xlAxisX_TP.MinimumScale = 0;
                //_xlAxisX_TP.MaximumScale = _maxTP;

                //_xlAxisY_RT.MinimumScale = 0;

                _xlChartRTPage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartRTPage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_RT = (Excel.SeriesCollection)_xlChartRTPage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_RT = (Excel.Series)_xlSeriesColl_RT.NewSeries();
                _xlSeries_RT.XValues = _xlRange_Analyte_SampleID.get_Offset(0, 11).Cells;
                _xlSeries_RT.Values = _xlRange_Analyte_AnalConc.get_Offset(0, -2).Cells; ;


                //string _serieCurveName = "H:L Abeta";

                //string _axeCurveXNumFormat = "0",
                _axeCurveYNumFormat = "0.00";

                //string _axeCurveXCaption = "Subject #",
                _axeCurveYCaption = "RT";

                if (ExperimentInfo._quanType == "ABS")
                {
                    _axeCurveYNumFormat = "0.0";
                    _axeCurveYCaption = "ng/mL, Concentration";
                    _serieCurveName = "Abeta level";
                }

                if (_standardCount == 0)
                {
                    _axeCurveYCaption = "Area Ratio";
                }

                //_xlAxisX_TP.TickLabels.NumberFormat = _axeCurveXNumFormat;
                _xlAxisY_RT.TickLabels.NumberFormat = _axeCurveYNumFormat;
                _xlAxisX_RT.AxisTitle.Caption = _axeCurveXCaption;
                _xlAxisY_RT.AxisTitle.Caption = _axeCurveYCaption;

                _xlSeries_RT.Name = _serieCurveName;
                _xlSeries_RT.Name = XEcompound.Attribute("name").Value;


                ///*********************************************************************************************************************


                //Std Curves***
                Excel.ChartObject _xlChartStdCurve = (Excel.ChartObject)_xlCharts.Add(
                                   490, _graphOffset +
                                        _xlChartCalcConc.Height +
                                        _xlChartISTD_Area.Height +
                                        _xlChartRT.Height, 470, 250);

                Excel.Chart _xlChartStdCurvePage = _xlChartStdCurve.Chart;

                _xlChartStdCurvePage.ChartType = Excel.XlChartType.xlXYScatter;

                _xlChartStdCurvePage.HasTitle = true;
                _xlChartStdCurvePage.ChartTitle.Font.Size = 11;

                string _chartTitleStdCurve = String.Format
                    ("Calibration curve {0} {1} {2} {3} ({4})",
                    XEcompound.Attribute("name").Value,
                    ExperimentInfo._IP,
                    ExperimentInfo._Enzyme,
                    ExperimentInfo._Instument,
                    ExperimentInfo._Date);


                _xlChartStdCurvePage.ChartTitle.Caption = _chartTitleStdCurve;



                Excel.Axes _xlAxes_Std = (Excel.Axes)_xlChartStdCurvePage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_Std = _xlAxes_Std.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_Std = _xlAxes_Std.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_Std.HasTitle = true; _xlAxisY_Std.HasTitle = true;
                _xlAxisX_Std.HasMajorGridlines = true;
                _xlAxisY_Std.HasMajorGridlines = true;
                _xlAxisY_Std.MinimumScale = 0;

                _xlChartStdCurvePage.Legend.IncludeInLayout = true;


                Excel.SeriesCollection _xlSerColl_StdCurve = (Excel.SeriesCollection)_xlChartStdCurvePage.SeriesCollection(Type.Missing);

                Excel.Series _xlSerie_Stds = (Excel.Series)_xlSerColl_StdCurve.NewSeries();
                _xlSerie_Stds.XValues = _xlRange_Std_Conc.Cells;
                _xlSerie_Stds.Values = _xlRange_Std_Response.Cells;


                string _serieStdName = "H:L Abeta";

                string _axeStdXNumFormat = "0.00",
                       _axeStdYNumFormat = "0.00";

                string _axeStdXCaption = "%, Predicted labeling Abeta",
                       _axeStdYCaption = "%, Measured labeling Abeta";

                if (ExperimentInfo._quanType == "ABS")
                {
                    _axeStdXNumFormat = "0";
                    _axeStdYNumFormat = "0";

                    _axeStdXCaption = "ng/mL, Specified amount";
                    _axeStdYCaption = "Area Ratio";

                    _serieStdName = "Analyte:ISTD Ratio";
                    _xlRange_Analyte_AnalConc.NumberFormat = "0.00";
                    _xlRange_Analyte_AnalConc.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
                    _xlRange_Analyte_AnalConc.Font.Bold = true;
                    _xlRange_Analyte_AnalConc.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    _xlRange_Analyte_AnalConc.IndentLevel = 2;

                    _xlRange_Std_AnalConc.NumberFormat = "0.00";
                    _xlRange_Std_AnalConc.Font.Bold = true;
                    _xlRange_Std_AnalConc.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    _xlRange_Std_AnalConc.IndentLevel = 2;

                }

                _xlAxisX_Std.TickLabels.NumberFormat = _axeStdXNumFormat;
                _xlAxisY_Std.TickLabels.NumberFormat = _axeStdYNumFormat;
                _xlAxisX_Std.AxisTitle.Caption = _axeStdXCaption;
                _xlAxisY_Std.AxisTitle.Caption = _axeStdYCaption;

                _xlSerie_Stds.Name = _serieStdName;

                Excel.Trendlines _xlTrendlines_Stds = (Excel.Trendlines)_xlSerie_Stds.Trendlines(Type.Missing);

                Excel.XlTrendlineType _XlTrendlineType =
                    Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear;
                string _CurveIndex = XECurve.Attribute("type").Value;
                //_wsComponent.get_Range("C2", "C2").Value2.ToString();

                if (_CurveIndex == "Quadratic")
                { _XlTrendlineType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlExponential; }


                _xlTrendlines_Stds.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, true, true, Type.Missing);

                //End Plots***

                //****************************
                Excel.ChartObjects _xlChartsSummary = (Excel.ChartObjects)_wsSummary.ChartObjects(Type.Missing);
                Excel.ChartObject _xlChartSummary1;
                Excel.Chart _xlChartSummary1Page;

                if (_xlChartsSummary.Count < 1)
                {
                    _xlChartSummary1 = (Excel.ChartObject)_xlChartsSummary.Add(
                                   10, 10, 850, 500);
                    _xlChartSummary1Page = _xlChartSummary1.Chart;
                    _xlChartSummary1Page.ChartType = Excel.XlChartType.xlColumnClustered;
                    _xlChartSummary1Page.HasTitle = true;
                    _xlChartSummary1Page.ChartTitle.Font.Size = 11;



                    Excel.Axes _xlAxes_Summary1 = (Excel.Axes)_xlChartSummary1Page.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisX_Summary1 = _xlAxes_Summary1.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisY_Summary1 = _xlAxes_Summary1.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                    _xlAxisX_Summary1.HasMajorGridlines = true;
                    _xlAxisY_Summary1.HasMajorGridlines = true;

                    //_xlAxisX_CurveSummary.Format.Line.Weight = 3.0F;
                    //_xlAxisY_CurveSummary.Format.Line.Weight = 3.0F;

                    _xlAxisX_Summary1.HasTitle = true;
                    _xlAxisY_Summary1.HasTitle = true;

                    //_xlAxisX_Summary1.MaximumScale = _maxTP;

                    _xlAxisX_Summary1.AxisTitle.Caption = _axeCurveXCaption;
                    _xlAxisY_Summary1.AxisTitle.Caption = "TTR";
                    //_axeCurveYCaption;

                    _xlAxisX_Summary1.TickLabels.NumberFormat = _axeCurveXNumFormat;
                    _xlAxisY_Summary1.TickLabels.NumberFormat = "0.0%";         // _axeCurveYNumFormat;


                    _xlChartSummary1Page.ChartTitle.Caption = "H:L Abeta ";

                    if (ExperimentInfo._quanType == "ABS")
                    {
                        _xlChartSummary1Page.ChartTitle.Caption = "Abeta levels in ";
                        try
                        {
                            _xlAxisY_Summary1.ScaleType = Excel.XlScaleType.xlScaleLogarithmic;
                            _xlAxisY_Summary1.LogBase = 10;
                        }
                        catch (Exception _exception)
                        { MessageBox.Show(_exception.Message); }
                        finally
                        { }
                    }

                    _xlChartSummary1Page.ChartTitle.Caption += ExperimentInfo._CommonChartNameStr;

                    _xlChartSummary1Page.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;
                    _xlChartSummary1Page.Legend.IncludeInLayout = true;

                }
                else
                {
                    _xlChartSummary1 = (Excel.ChartObject)_xlChartsSummary.Item(1);
                    _xlChartSummary1Page = _xlChartSummary1.Chart;
                }


                Excel.SeriesCollection _xlSeriesColl_Summary1 = (Excel.SeriesCollection)_xlChartSummary1Page.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_Summary1 = (Excel.Series)_xlSeriesColl_Summary1.NewSeries();
                _xlSeries_Summary1.XValues = _xlRange_Analyte_SampleID.get_Offset(0, 11).Cells;
                _xlSeries_Summary1.Values = _xlRange_Analyte_AnalConc;
                _xlSeries_Summary1.Name = XEcompound.Attribute("name").Value;

                //***************************

            }








            toolStripProgressBarMain.Visible = false;
            toolStripProgressBarMain.Value = 0;
            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();


            string _AppPath = Application.StartupPath + Path.DirectorySeparatorChar;
            toolStripStatusLabelMain.Text = "Exporting data into DB ...";
            toolStripStatusLabelMain.PerformClick();
            string[] _sql_Part1 = File.ReadAllLines(_AppPath + "sql_XevoImport_P1.sql");
            string[] _sql_Part2 = File.ReadAllLines(_AppPath + "sql_XevoImport_P2.sql");
            string[] _sql_Part3 = File.ReadAllLines(_AppPath + "sql_XevoImport_P3.sql");

            _strExportSQL.Insert(0, " ");

            _strExportSQL.InsertRange(0, _sql_Part1.ToList<string>());

            _strExportSQL.Insert(0, String.Format("/* Generated by QLDReader on {0} @ {1} */", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString()));

            _strExportSQL.AddRange(_sql_Part3.ToList<string>());

            File.WriteAllLines(Path.ChangeExtension(textBox_Excelfile.Text, "sql"), _strExportSQL.ToArray());

            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();



            object misValue = System.Reflection.Missing.Value;

            try
            {
                xlWB.SaveAs(textBox_Excelfile.Text,
                   misValue,
                   misValue, misValue, misValue, misValue,
                   Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlUserResolution,
                   misValue, misValue, misValue, misValue);
            }

            catch (COMException _exception)
            {
                if (_exception.ErrorCode == unchecked((int)0x800A03EC))
                {
                    MessageBox.Show("Time stamp has been added to file name you specified", this.Text,
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    string _newExcelFileName = textBox_Excelfile.Text.Replace(
                        Path.GetFileNameWithoutExtension(textBox_Excelfile.Text),
                        Path.GetFileNameWithoutExtension(textBox_Excelfile.Text)
                        + "_created_"
                        + System.DateTime.Now.ToString("s")
                        .Replace("/", "")
                        .Replace(":", "")
                        .Replace(" ", "_")
                        );

                    xlWB.SaveAs(_newExcelFileName,
                       misValue,
                       misValue, misValue, misValue, misValue,
                       Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlUserResolution,
                       misValue, misValue, misValue, misValue);
                }
            }

            finally
            {
                if (this.checkBox_OpenExcel.Checked)
                {
                    _xlApp.Visible = true;
                    // _wsSummary.Activate();

                }
                else
                {
                    _xlApp.Quit();
                }
            }

        }

        private void ExportFACS_csf()
        {
            List<string> _strExportSQL = new List<string>();

            XDocument _QLDDoc;
            try
            {
                _QLDDoc = XDocument.Load(textBox_QLDfile.Text);
            }
            catch (Exception _XMLExeption)
            {
                MessageBox.Show(_XMLExeption.Message);
                return;
            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = true;
            toolStripStatusLabelMain.Visible = true;

            _xlApp = new Excel.Application();
            _xlApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;

            Excel.Workbook xlWB;

            try
            {
                xlWB =
                    _xlApp.Workbooks.Add(
                                         Application.StartupPath + Path.DirectorySeparatorChar +
                                         @"\Templates\FACS_Xevo_Template_Summary.xlsx"
                );
            }

            catch (Exception)
            {
                MessageBox.Show("File: " + Application.StartupPath + Path.DirectorySeparatorChar +
                @"\Templates\FACS_Xevo_Template_Summary.xlsx", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            finally
            { }

            //MessageBox.Show(ExperimentInfo._quanType);

            string _quanTypeKey = "";

            if (ExperimentInfo._quanType == "RELATIVE")
            {
                _quanTypeKey = "C13N14";
            }
            if (ExperimentInfo._quanType == "ABSOLUTE")
            {
                _quanTypeKey = "C12N14";
            }
            if (ExperimentInfo._quanType == "REL+ABS")
            {
                _quanTypeKey = "N14";
            }

            if (ExperimentInfo._quanType == "TP")
            {
                _quanTypeKey = "Tra";
            }

            Excel.Worksheet _wsSummary = (Excel.Worksheet)xlWB.ActiveSheet;

            Excel.Worksheet _wsCurrent = new Microsoft.Office.Interop.Excel.Worksheet();

            var queryCOMPOUND_LIST = from compound in _QLDDoc.Descendants("CALIBRATIONDATA").Descendants("COMPOUND")
                                     where

                                     compound.Attribute("name").Value.Contains(_quanTypeKey)
                                           &&
                                           !compound.Attribute("name").Value.Contains("Ch")

                                           // select only group 1   //2013-05-03 VO
                                           && compound.Parent.Parent.Attribute("id").Value == "1"

                                     select compound;


            foreach (XElement XEcompound in queryCOMPOUND_LIST)
            {
                //MessageBox.Show(XEcompound.Attribute("name").Value, "Compound");

                _wsCurrent = (Excel.Worksheet)xlWB.Sheets.Add
                          (_wsSummary, Type.Missing, Type.Missing,
                          Application.StartupPath + Path.DirectorySeparatorChar +
                            @"\Templates\FACS_Xevo_Template.xlsx");
                string _sheetName = XEcompound.Attribute("name").Value;

                _wsCurrent.Cells[2, 1] = _sheetName;

                if (_sheetName.Length > 31)
                {
                    _sheetName = "alt_" + _sheetName.Remove(1, _sheetName.Length - 31 + 4);
                }


                _wsCurrent.Name = _sheetName.Replace("+", "_").Replace(" ", "_");

                //_wsCurrent.Cells[2, 1] = XEcompound.Attribute("name").Value;






                var XECurve = XEcompound.Element("CURVE");
                _wsCurrent.Cells[2, 3] = XECurve.Attribute("type").Value;
                _wsCurrent.Cells[2, 6] = XECurve.Attribute("origin").Value;
                _wsCurrent.Cells[2, 4] = XECurve.Attribute("weighting").Value;
                _wsCurrent.Cells[2, 8] = XECurve.Attribute("axistrans").Value;


                try
                {
                    _wsCurrent.Cells[2, 8] = XECurve.Element("CALIBRATIONCURVE").Attribute("curve").Value;
                    _wsCurrent.Cells[2, 11] = XECurve.Element("CORRELATION").Attribute("r").Value;
                    _wsCurrent.Cells[2, 13] = XECurve.Element("DETERMINATION").Attribute("rsquared").Value;
                }
                catch (Exception)
                {
                    if (checkBox_STD.Checked)
                    {
                        MessageBox.Show(String.Format("Calibration curve parameters for {0} unavailable", XEcompound.Attribute("name").Value));
                    }
                }



                var querySAMPLE_LIST = from sample in _QLDDoc.Descendants("SAMPLE")
                                       // orderby sample.Attribute("type").Value descending, sample.Attribute("id").Value
                                       where (sample.Attribute("type").Value == "Standard" ||
                                              sample.Attribute("type").Value == "Analyte" ||
                                              sample.Attribute("type").Value == "QC")
                                       && sample.Attribute("name").Value != ""
                                       orderby sample.Attribute("type").Value.Length descending
                                       select sample;


                int _sampleCount = querySAMPLE_LIST.Count();
                int _standardCount = querySAMPLE_LIST.Count(p => p.Attribute("type").Value == "Standard");
                int _analyteCount = querySAMPLE_LIST.Count(p => p.Attribute("type").Value == "Analyte");
                int _qcCount = querySAMPLE_LIST.Count(p => p.Attribute("type").Value == "QC");


                double _maxTP = (from _sample in querySAMPLE_LIST select _sample).Count();

                try
                {
                    _maxTP = (from _sample in querySAMPLE_LIST
                              where _sample.Attribute("type").Value == "Analyte"
                                               && double.TryParse(_sample.Attribute("sampleid").Value, out _maxTP)
                              select Convert.ToDouble(_sample.Attribute("sampleid").Value)).Max();
                }
                catch
                {
                }



                //MessageBox.Show(test.ToString());

                int index = _OutputFormat._titleShift + 1;
                foreach (XElement XEsample in querySAMPLE_LIST)
                {
                    _wsCurrent.Cells[index, 1] = XEsample.Attribute("name").Value;
                    _wsCurrent.Cells[index, 2] = XEsample.Attribute("type").Value;
                    _wsCurrent.Cells[index, 3] = XEsample.Attribute("sampleid").Value;
                    _wsCurrent.Cells[index, 14] = XEsample.Attribute("subjecttext").Value;

                    // _wsCurrent.Cells[index, 20] = XEsample.Parent.Parent.Attribute("name").Value;


                    var queryCOMPOUND = from compound in XEsample.Descendants("COMPOUND")
                                        where compound.Attribute("name").Value == XEcompound.Attribute("name").Value
                                        select compound;

                    _wsCurrent.Cells[index, 4] = queryCOMPOUND.Single().Attribute("stdconc").Value;

                    var XEpeak = queryCOMPOUND.First().Element("PEAK");
                    var XEpeakis = XEpeak.Element("ISPEAK");


                    /*
                    if (XEsample.Attribute("id").Value == "34")
                    {
                        XEpeak = XEsample.Element("COMPOUND").Element("PEAK");
                        foreach (XAttribute _at in XEpeak.Attributes() )
                        {
                          MessageBox.Show
                              ( XEcompound.Attribute("name").Value.ToString()+ "  " 
                                 + _at.Name+ "  -   "+ _at.Value.ToString());
                        }
                        //MessageBox.Show(XEpeak. .Element .Attribute("area").Value.ToString());
                    }
                    */

                    if (!(XEpeak.Attribute("pkflags").Value.ToString() == ""))
                    {
                        _wsCurrent.Cells[index, 5] = XEpeak.Attribute("area").Value;
                        _wsCurrent.Cells[index, 6] = XEpeakis.Attribute("area").Value;
                        _wsCurrent.Cells[index, 7] = XEpeak.Attribute("foundrt").Value;
                        _wsCurrent.Cells[index, 8] = XEpeakis.Attribute("foundrt").Value;
                        _wsCurrent.Cells[index, 9] = XEpeak.Attribute("response").Value;

                        // if (!XEpeak.Attribute("analconc").Value.StartsWith("0.00000"))
                        // {
                        _wsCurrent.Cells[index, _OutputFormat._colNum_analconc] =
                            XEpeak.Attribute("analconc").Value;
                        // }

                        _wsCurrent.Cells[index, 11] = XEpeak.Attribute("concdevperc").Value;
                        _wsCurrent.Cells[index, 12] = XEpeak.Attribute("signoise").Value;
                        _wsCurrent.Cells[index, 13] = XEpeak.Attribute("chromnoise").Value;


                    }



                    //Export into DB

                    if (checkBox_ExportIntoDB.Checked)
                    {
                        if (XEpeak.Attribute("pkflags").Value.ToString() == "" || XEsample.Attribute("type").Value == "QC")
                        {
                            // MessageBox.Show(String.Format("Sample {1} on analyte {0} show no peak",
                            //                    XEcompound.Attribute("name").Value,
                            //                    XEsample.Attribute("name").Value));
                        }
                        else
                        {
                            _strExportSQL.Add("INSERT INTO TMP_XEVO_IMPORT (" +
                                                  " EXPER_FILE, ASSAY_DATE, ASSAY_TYPE_ID," +
                                                  " ANTIBODY_ID, ENZYME_ID, QUAN_TYPE_ID," +
                                                  " SAMPLE_PROCESS_BY, ASSAY_DONE_BY, DATA_PROCESS_BY," +
                                                  " FILENAME_DESC, SAMPLE_TYPE_ID, LEVEL_TP_ID," +
                                                  " NUM_SPEC_CONC, SUBJECT_NUM, FLUID_TYPE_ID," +
                                                  " ANYLYTE_NAME, AREA, HEIGHT," +
                                                  " RT, ISTD_AREA, ISTD_HEIGHT," +
                                                  " ISTD_RT, RESPONSE, CURVE," +
                                                  " WEIGHTING, ORIGIN, EQUATION," +
                                                  " R, R_SQR, NUM_SLOPE," +
                                                  " NUM_INTERSECT, NUM_ANAL_CONC, CONC_DEV_PERC," +
                                                  " SIGNOISE, CHROMNOISE, SAMPLE_GROUP )");

                            _strExportSQL.Add(
                                    String.Format("VALUES " +
                                                     "( '{0}', '{1}', {2}, " +
                                                     "{3}, {4}, {5}, " +
                                                     "{6}, {7}, {8}, " +
                                                     "'{9}', {10}, {11}, " +
                                                     "{12}, '{13}', '{14}', " +
                                                     "'{15}', {16}, {17}, " +
                                                     "{18}, {19}, {20}, " +
                                                     "{21}, {22}, '{23}', " +
                                                     "'{24}', '{25}', '{26}', " +
                                                     "{27}, {28}, {29}, " +
                                                     "{30}, {31}, {32}, {33}, " +
                                                     "{34}, '{35}' );"
                                    ,
                                    ExperimentInfo._expName,                          // 0*EXPER_FILE
                                    ExperimentInfo._Date,                             // 1*ASSAY_DATE
                                    7,                                                // 2*ASSAY_TYPE_ID

                                    IPToDBInt(ExperimentInfo._IP),                    // 3*ANTIBODY_ID  
                                    EnzymeToDBInt(ExperimentInfo._Enzyme),            // 4*ENZYME_ID
                                    QuntTypeToDBInt(ExperimentInfo._quanType),        // 5*INTEG_TYPE_ID

                                    ExperimentInfo._SampleProcessBy,                  // 6*SAMPLE_PROCESS_BY
                                    ExperimentInfo._DoneBy,                           // 7*ASSAY_DONE_BY
                                    ExperimentInfo._QuantitatedBy,                    // 8*DATA_PROCESS_BY

                                    XEsample.Attribute("name").Value,                 // 9*FILENAME_DESC
                                    SampleTypeToDBInt
                                    (XEsample.Attribute("type").Value),               //10*SAMPLE_TYPE_ID
                                    GetLevel_TP(
                                    XEsample.Attribute("sampleid").Value,
                                        XEsample.Attribute("type").Value),             //11*LEVEL_TP_ID

                                    queryCOMPOUND.Single().Attribute("stdconc").Value, //12*NUM_SPEC_CONC
                                    XEsample.Attribute("subjecttext").Value,           //13*SUBJECT_NUM     
                                    FluidTypeToDBInt(ExperimentInfo._matrix),          //14*FLUID_TYPE_ID

                                    XEcompound.Attribute("name").Value,                //15*ANALYTE_NAME
                                    XEpeak.Attribute("area").Value,                    //16*AREA
                                    XEpeak.Attribute("height").Value,                  //17*HEIGHT

                                    XEpeak.Attribute("foundrt").Value,                 //18*RT
                                    XEpeakis.Attribute("area").Value,                  //19*ISTD AREA
                                    XEpeakis.Attribute("height").Value,                //20*ISTD HEIGHT

                                    XEpeakis.Attribute("foundrt").Value,               //21*ISTD RT
                                    XEpeak.Attribute("response").Value,                //22*RESPONSE

                                    XECurve.Attribute("type").Value,                   //23*CURVE
                                    XECurve.Attribute("weighting").Value,              //24*WEIGHTING
                                    XECurve.Attribute("origin").Value,                 //25*ORIGIN
                                    XECurve.Element("CALIBRATIONCURVE")
                                        .Attribute("curve").Value,                     //26*EQUATION
                                    XECurve.Element("CORRELATION")
                                        .Attribute("r").Value,                         //27*R
                                    XECurve.Element("DETERMINATION")
                                        .Attribute("rsquared").Value,                  //28*R_SQR
                                    GetSlopeOfEquation(
                                    XECurve.Element("CALIBRATIONCURVE")
                                        .Attribute("curve").Value),                    //29*NUM_SLOPE
                                    GetInterOfEquation(
                                    XECurve.Element("CALIBRATIONCURVE")
                                        .Attribute("curve").Value),                    //30*NUM_INTERSECT
                                    AnalConc_EmptyCheck(
                                             XEpeak.Attribute("analconc").Value),      //31*NUM_ANAL_CONC
                                    XEpeak.Attribute("concdevperc").Value,             //32*CONC_DEV_PERC
                                    (XEpeak.Attribute("signoise").Value == "") ? "NULL" : XEpeak.Attribute("signoise").Value,    //33*SIGNOISE
                                    XEpeak.Attribute("chromnoise").Value,       //34*CHROMNOISE
                                    XEsample.Parent.Parent.Attribute("name").Value)
                                    );
                        }
                    }

                    // ProgressBar Begin
                    toolStripProgressBarMain.PerformStep();
                    toolStripStatusLabelMain.Text = "Proccesing ... " +
                        XEcompound.Attribute("name").Value + " - " + XEsample.Attribute("sampleid").Value;
                    toolStripStatusLabelMain.PerformClick();
                    // ProgressBar End

                    index++;

                }


                Excel.Range _xlRange_Std_Conc, _xlRange_Std_Response, _xlRange_Std_Dev, _xlRange_Std_AnalConc,
                            _xlRange_Analyte_SampleID, _xlRange_Analyte_AnalConc, _xlRange_Analyte_ISArea,
                            _xlRange_Std_Box, _xlRange_Analyte_Box, _xlRange_Sample_List_Box,
                            _xlRange_Tableau;

                //Standard Ranges****
                _xlRange_Std_Conc = _wsCurrent.get_Range
                (_OutputFormat._colLet_stdconc + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_stdconc + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Conc.Name = _wsCurrent.Name + "_Std_Conc";

                _xlRange_Std_Response = _wsCurrent.get_Range
                (_OutputFormat._colLet_response + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_response + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Response.Name = _wsCurrent.Name + "_Std_Response";

                _xlRange_Std_AnalConc = _wsCurrent.get_Range
                   (_OutputFormat._colLet_analconc + (_OutputFormat._titleShift + 1).ToString(),
                    _OutputFormat._colLet_analconc + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_AnalConc.Name = _wsCurrent.Name + "_Std_AnalConc";

                _xlRange_Std_Dev = _wsCurrent.get_Range
                (_OutputFormat._colLet_stddev + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_stddev + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Dev.Name = _wsCurrent.Name + "_Std_Deviation";

                _xlRange_Std_Box = _wsCurrent.get_Range
                (_OutputFormat._colLet_First + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Box.Name = _wsCurrent.Name + "_Std_Box";



                //Analyte Ranges****
                _xlRange_Analyte_SampleID = _wsCurrent.get_Range
             (_OutputFormat._colLet_sampleid + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
              _OutputFormat._colLet_sampleid + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_SampleID.Name = _wsCurrent.Name + "_Analyte_SampleID";

                _xlRange_Analyte_AnalConc = _wsCurrent.get_Range
             (_OutputFormat._colLet_analconc + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
              _OutputFormat._colLet_analconc + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_AnalConc.Name = _wsCurrent.Name + "_Analyte_AnalConc";

                _xlRange_Analyte_ISArea = _wsCurrent.get_Range
              (_OutputFormat._colLet_analISarea + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
               _OutputFormat._colLet_analISarea + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_ISArea.Name = _wsCurrent.Name + "_Analyte_ISArea";

                _xlRange_Analyte_Box = _wsCurrent.get_Range
             (_OutputFormat._colLet_First + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
              _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_Box.Name = _wsCurrent.Name + "_Analyte_Box";


                //SAMPLE_LIST Range
                _xlRange_Sample_List_Box = _wsCurrent.get_Range
              (_OutputFormat._colLet_First + (_OutputFormat._titleShift + 1).ToString(),
            _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Sample_List_Box.Name = _wsCurrent.Name + "_Sample_Box";

                //Tableau Range
                _xlRange_Tableau = _wsCurrent.get_Range
              (_OutputFormat._colLet_First + (_OutputFormat._titleShift).ToString(),
            _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Tableau.Name = _wsCurrent.Name + "_Tableau";



                //Format Ranges**** 
                _xlRange_Sample_List_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                _xlRange_Sample_List_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).Weight = 3.75;
                _xlRange_Sample_List_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).Weight = 3.75;
                _xlRange_Std_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                _xlRange_Std_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).Weight = 1.75;
                _xlRange_Analyte_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).Weight = 1.75;

                _xlRange_Std_Box.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);

                //////////////////////////////////////////////////////
                Excel.ColorScale cfColorScale = (Excel.ColorScale)(_xlRange_Analyte_ISArea.FormatConditions.AddColorScale(3));
                // Set the minimum threshold to red (0x000000FF) and maximum threshold
                // to blue (0x00FF0000).
                cfColorScale.ColorScaleCriteria[1].FormatColor.Color = 7039480;  //0x000000FF;  
                cfColorScale.ColorScaleCriteria[2].FormatColor.Color = 8711167;  // 0x00FF0000;
                cfColorScale.ColorScaleCriteria[3].FormatColor.Color = 13011546; // 0x00FF0000;
                /////////////////////////////////////////////////////



                //Plots****
                //Time Serie****

                double _rowHeight = (double)_wsCurrent.get_Range("A1", "A1").RowHeight;

                double _graphOffset = (double)_xlRange_Sample_List_Box.Top + (double)_xlRange_Sample_List_Box.Height +
                       (_rowHeight + _rowHeight * _qcCount);

                Excel.ChartObjects _xlCharts = (Excel.ChartObjects)_wsCurrent.ChartObjects(Type.Missing);

                Excel.ChartObject _xlChartTimeCorse = (Excel.ChartObject)_xlCharts.Add(
                                  0, _graphOffset, 470, 250);
                Excel.Chart _xlChartTimeCorsePage = _xlChartTimeCorse.Chart;
                _xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlXYScatter;
                _xlChartTimeCorsePage.HasTitle = true;
                _xlChartTimeCorsePage.ChartTitle.Font.Size = 11;



                _xlChartTimeCorsePage.ChartTitle.Caption = XEcompound.Attribute("name").Value + " " + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_TP = (Excel.Axes)_xlChartTimeCorsePage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_TP.HasMajorGridlines = true;
                _xlAxisX_TP.HasTitle = true;
                _xlAxisY_TP.HasTitle = true;

                _xlAxisX_TP.MinimumScale = 0;
                _xlAxisX_TP.MaximumScale = _maxTP;

                _xlAxisY_TP.MinimumScale = 0;

                _xlChartTimeCorsePage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartTimeCorsePage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_TimeCourse = (Excel.SeriesCollection)_xlChartTimeCorsePage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_TTR = (Excel.Series)_xlSeriesColl_TimeCourse.NewSeries();
                _xlSeries_TTR.XValues = _xlRange_Analyte_SampleID.Cells;
                _xlSeries_TTR.Values = _xlRange_Analyte_AnalConc;

                //*****************
                _xlSeries_TTR.ChartType = Excel.XlChartType.xlXYScatterSmooth;
                //****************


                string _serieCurveName = "H:L Abeta";

                string _axeCurveXNumFormat = "0",
                       _axeCurveYNumFormat = "0.0%";

                string _axeCurveXCaption = "h, Time",
                       _axeCurveYCaption = "H:L ratio";

                if (ExperimentInfo._quanType == "ABS")
                {
                    _axeCurveYNumFormat = "0.0";
                    _axeCurveYCaption = "ng/mL, Concentration";
                    _serieCurveName = "Abeta level";
                }

                if (_standardCount == 0)
                {
                    _axeCurveYCaption = "Area Ratio";
                }

                _xlAxisX_TP.TickLabels.NumberFormat = _axeCurveXNumFormat;
                _xlAxisY_TP.TickLabels.NumberFormat = _axeCurveYNumFormat;
                _xlAxisX_TP.AxisTitle.Caption = _axeCurveXCaption;
                _xlAxisY_TP.AxisTitle.Caption = _axeCurveYCaption;

                _xlSeries_TTR.Name = _serieCurveName;
                _xlSeries_TTR.Name = XEcompound.Attribute("name").Value;

                //Std Curves***
                Excel.ChartObject _xlChartStdCurve = (Excel.ChartObject)_xlCharts.Add(
                                   490, _graphOffset, 470, 250);
                Excel.Chart _xlChartStdCurvePage = _xlChartStdCurve.Chart;
                _xlChartStdCurvePage.ChartType = Excel.XlChartType.xlXYScatter;

                _xlChartStdCurvePage.HasTitle = true;
                _xlChartStdCurvePage.ChartTitle.Font.Size = 11;

                string _chartTitleStdCurve = String.Format
                    ("Calibration curve {0} {1} {2} {3} ({4})",
                    XEcompound.Attribute("name").Value,
                    ExperimentInfo._IP,
                    ExperimentInfo._Enzyme,
                    ExperimentInfo._Instument,
                    ExperimentInfo._Date);


                _xlChartStdCurvePage.ChartTitle.Caption = _chartTitleStdCurve;



                Excel.Axes _xlAxes_Std = (Excel.Axes)_xlChartStdCurvePage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_Std = _xlAxes_Std.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_Std = _xlAxes_Std.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_Std.HasTitle = true; _xlAxisY_Std.HasTitle = true;
                _xlAxisX_Std.HasMajorGridlines = true;
                _xlAxisY_Std.HasMajorGridlines = true;
                _xlAxisY_Std.MinimumScale = 0;

                _xlChartStdCurvePage.Legend.IncludeInLayout = true;


                Excel.SeriesCollection _xlSerColl_StdCurve = (Excel.SeriesCollection)_xlChartStdCurvePage.SeriesCollection(Type.Missing);

                Excel.Series _xlSerie_Stds = (Excel.Series)_xlSerColl_StdCurve.NewSeries();
                _xlSerie_Stds.XValues = _xlRange_Std_Conc.Cells;
                _xlSerie_Stds.Values = _xlRange_Std_Response.Cells;


                string _serieStdName = "H:L Abeta";

                string _axeStdXNumFormat = "0.00",
                       _axeStdYNumFormat = "0.00";

                string _axeStdXCaption = "%, Predicted labeling Abeta",
                       _axeStdYCaption = "%, Measured labeling Abeta";

                if (ExperimentInfo._quanType == "ABS")
                {
                    _axeStdXNumFormat = "0";
                    _axeStdYNumFormat = "0";

                    _axeStdXCaption = "ng/mL, Specified amount";
                    _axeStdYCaption = "Area Ratio";

                    _serieStdName = "Analyte:ISTD Ratio";
                    _xlRange_Analyte_AnalConc.NumberFormat = "0.00";
                    _xlRange_Analyte_AnalConc.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
                    _xlRange_Analyte_AnalConc.Font.Bold = true;
                    _xlRange_Analyte_AnalConc.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    _xlRange_Analyte_AnalConc.IndentLevel = 2;

                    _xlRange_Std_AnalConc.NumberFormat = "0.00";
                    _xlRange_Std_AnalConc.Font.Bold = true;
                    _xlRange_Std_AnalConc.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    _xlRange_Std_AnalConc.IndentLevel = 2;

                }

                _xlAxisX_Std.TickLabels.NumberFormat = _axeStdXNumFormat;
                _xlAxisY_Std.TickLabels.NumberFormat = _axeStdYNumFormat;
                _xlAxisX_Std.AxisTitle.Caption = _axeStdXCaption;
                _xlAxisY_Std.AxisTitle.Caption = _axeStdYCaption;

                _xlSerie_Stds.Name = _serieStdName;

                Excel.Trendlines _xlTrendlines_Stds = (Excel.Trendlines)_xlSerie_Stds.Trendlines(Type.Missing);

                Excel.XlTrendlineType _XlTrendlineType =
                    Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear;
                string _CurveIndex = XECurve.Attribute("type").Value;
                //_wsComponent.get_Range("C2", "C2").Value2.ToString();

                if (_CurveIndex == "Quadratic")
                { _XlTrendlineType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlExponential; }


                _xlTrendlines_Stds.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, true, true, Type.Missing);

                //End Plots***

                //****************************
                Excel.ChartObjects _xlChartsSummary = (Excel.ChartObjects)_wsSummary.ChartObjects(Type.Missing);
                Excel.ChartObject _xlChartCurveSummary;
                Excel.Chart _xlChartCurveSummaryPage;

                if (_xlChartsSummary.Count < 1)
                {
                    _xlChartCurveSummary = (Excel.ChartObject)_xlChartsSummary.Add(
                                   10, 10, 850, 500);
                    _xlChartCurveSummaryPage = _xlChartCurveSummary.Chart;
                    _xlChartCurveSummaryPage.ChartType = Excel.XlChartType.xlXYScatterSmooth;
                    _xlChartCurveSummaryPage.HasTitle = true;
                    _xlChartCurveSummaryPage.ChartTitle.Font.Size = 11;



                    Excel.Axes _xlAxes_CurveSummary = (Excel.Axes)_xlChartCurveSummaryPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisX_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisY_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                    _xlAxisX_CurveSummary.HasMajorGridlines = true;
                    _xlAxisY_CurveSummary.HasMajorGridlines = true;

                    //_xlAxisX_CurveSummary.Format.Line.Weight = 3.0F;
                    //_xlAxisY_CurveSummary.Format.Line.Weight = 3.0F;

                    _xlAxisX_CurveSummary.HasTitle = true;
                    _xlAxisY_CurveSummary.HasTitle = true;

                    _xlAxisX_CurveSummary.MaximumScale = _maxTP;

                    _xlAxisX_CurveSummary.AxisTitle.Caption = _axeCurveXCaption;
                    _xlAxisY_CurveSummary.AxisTitle.Caption = _axeCurveYCaption;

                    _xlAxisX_CurveSummary.TickLabels.NumberFormat = _axeCurveXNumFormat;
                    _xlAxisY_CurveSummary.TickLabels.NumberFormat = _axeCurveYNumFormat;


                    _xlChartCurveSummaryPage.ChartTitle.Caption = "H:L Abeta ";

                    if (ExperimentInfo._quanType == "ABS")
                    {
                        _xlChartCurveSummaryPage.ChartTitle.Caption = "Abeta levels in ";
                        try
                        {
                            _xlAxisY_CurveSummary.ScaleType = Excel.XlScaleType.xlScaleLogarithmic;
                            _xlAxisY_CurveSummary.LogBase = 10;
                        }
                        catch (Exception _exception)
                        { MessageBox.Show(_exception.Message); }
                        finally
                        { }
                    }

                    _xlChartCurveSummaryPage.ChartTitle.Caption += ExperimentInfo._CommonChartNameStr;

                    _xlChartCurveSummaryPage.Legend.Position = Excel.XlLegendPosition.xlLegendPositionRight;
                    _xlChartCurveSummaryPage.Legend.IncludeInLayout = true;

                }
                else
                {
                    _xlChartCurveSummary = (Excel.ChartObject)_xlChartsSummary.Item(1);
                    _xlChartCurveSummaryPage = _xlChartCurveSummary.Chart;
                }


                Excel.SeriesCollection _xlSeriesColl_CurveSummary = (Excel.SeriesCollection)_xlChartCurveSummaryPage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_CurveTTR = (Excel.Series)_xlSeriesColl_CurveSummary.NewSeries();
                _xlSeries_CurveTTR.XValues = _xlRange_Analyte_SampleID.Cells;
                _xlSeries_CurveTTR.Values = _xlRange_Analyte_AnalConc;
                _xlSeries_CurveTTR.Name = XEcompound.Attribute("name").Value;

                //***************************

            }








            toolStripProgressBarMain.Visible = false;
            toolStripProgressBarMain.Value = 0;
            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();


            string _AppPath = Application.StartupPath + Path.DirectorySeparatorChar;
            toolStripStatusLabelMain.Text = "Exporting data into DB ...";
            toolStripStatusLabelMain.PerformClick();
            string[] _sql_Part1 = File.ReadAllLines(_AppPath + "sql_XevoImport_P1.sql");
            string[] _sql_Part2 = File.ReadAllLines(_AppPath + "sql_XevoImport_P2.sql");
            string[] _sql_Part3 = File.ReadAllLines(_AppPath + "sql_XevoImport_P3.sql");

            _strExportSQL.Insert(0, " ");
            _strExportSQL.InsertRange(0, _sql_Part1.ToList<string>());
            _strExportSQL.AddRange(_sql_Part3.ToList<string>());

            File.WriteAllLines(Path.ChangeExtension(textBox_Excelfile.Text, "sql"), _strExportSQL.ToArray());

            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();



            object misValue = System.Reflection.Missing.Value;

            try
            {
                xlWB.SaveAs(textBox_Excelfile.Text,
                   misValue,
                   misValue, misValue, misValue, misValue,
                   Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlUserResolution,
                   misValue, misValue, misValue, misValue);
            }

            catch (COMException _exception)
            {
                if (_exception.ErrorCode == unchecked((int)0x800A03EC))
                {
                    MessageBox.Show("Time stamp has been added to file name you specified", this.Text,
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    string _newExcelFileName = textBox_Excelfile.Text.Replace(
                        Path.GetFileNameWithoutExtension(textBox_Excelfile.Text),
                        Path.GetFileNameWithoutExtension(textBox_Excelfile.Text)
                        + "_created_"
                        + System.DateTime.Now.ToString("s")
                        .Replace("/", "")
                        .Replace(":", "")
                        .Replace(" ", "_")
                        );

                    xlWB.SaveAs(_newExcelFileName,
                       misValue,
                       misValue, misValue, misValue, misValue,
                       Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlUserResolution,
                       misValue, misValue, misValue, misValue);
                }
            }

            finally
            {
                if (this.checkBox_OpenExcel.Checked)
                {
                    _xlApp.Visible = true;
                    // _wsSummary.Activate();

                }
                else
                {
                    _xlApp.Quit();
                }
            }

        }

        private void ExportIVBolus_csf()
        {
            List<string> _strExportSQL = new List<string>();

            XDocument _QLDDoc;
            try
            {
                _QLDDoc = XDocument.Load(textBox_QLDfile.Text);
            }
            catch (Exception _XMLExeption)
            {
                MessageBox.Show(_XMLExeption.Message);
                return;
            }

            toolStripProgressBarMain.Value = 0;


            toolStripProgressBarMain.Visible = true;
            toolStripStatusLabelMain.Visible = true;

            _xlApp = new Excel.Application();
            _xlApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;

            Excel.Workbook xlWB;

            try
            {
                xlWB =
                    _xlApp.Workbooks.Add(
                                         Application.StartupPath + Path.DirectorySeparatorChar +
                                         @"\Templates\IVBOLUS_CSF_Xevo_Template_Summary.xlsx"
                );
            }

            catch (Exception)
            {
                MessageBox.Show("File: " + Application.StartupPath + Path.DirectorySeparatorChar +
                @"\Templates\IVBOLUS_CSF_Xevo_Template_Summary.xlsx", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //MessageBox.Show(_exception.Message, _exception.GetBaseException().Message);
                return;
            }

            finally
            { }

            //MessageBox.Show(ExperimentInfo._quanType);

            string _quanTypeKey = "";

            if (ExperimentInfo._quanType == "RELATIVE")
            {
                _quanTypeKey = "C13N14";
            }
            if (ExperimentInfo._quanType == "ABSOLUTE")
            {
                _quanTypeKey = "C12N14";
            }
            if (ExperimentInfo._quanType == "REL+ABS")
            {
                _quanTypeKey = "N14";
            }

            Excel.Worksheet _wsSummary = (Excel.Worksheet)xlWB.ActiveSheet;

            Excel.Worksheet _wsCurrent = new Microsoft.Office.Interop.Excel.Worksheet();

            var queryCOMPOUND_LIST = from compound in _QLDDoc.Descendants("CALIBRATIONDATA").Descendants("COMPOUND")
                                     where compound.Attribute("name").Value.Contains(_quanTypeKey)
                                           &&
                                           !compound.Attribute("name").Value.Contains("Ch")

                                           // select only group 1   //2013-05-03 VO
                                           && compound.Parent.Parent.Attribute("id").Value == "1"

                                     select compound;

            toolStripProgressBarMain.Maximum =
               Convert.ToInt32(_QLDDoc.Descendants("SAMPLELISTDATA").First().Attribute("count").Value) *
                               queryCOMPOUND_LIST.Count();

            //MessageBox.Show(toolStripProgressBarMain.Maximum.ToString());

            foreach (XElement XEcompound in queryCOMPOUND_LIST)
            {
                //MessageBox.Show(XEcompound.Attribute("name").Value, "Compound");

                _wsCurrent = (Excel.Worksheet)xlWB.Sheets.Add
                          (_wsSummary, Type.Missing, Type.Missing,
                          Application.StartupPath + Path.DirectorySeparatorChar +
                            @"\Templates\IVBOLUS_CSF_Xevo_Template.xlsx");
                string _sheetName = XEcompound.Attribute("name").Value;

                int _compoundQuanType = QuntTypeToDBInt(XEcompound.Attribute("name").Value);

                _wsCurrent.Cells[2, 1] = _sheetName;

                if (_sheetName.Length > 31)
                {
                    _sheetName = "alt_" + _sheetName.Remove(1, _sheetName.Length - 31 + 4);
                }


                _wsCurrent.Name = _sheetName.Replace("+", "_");

                //_wsCurrent.Cells[2, 1] = XEcompound.Attribute("name").Value;






                var XECurve = XEcompound.Element("CURVE");
                _wsCurrent.Cells[2, 3] = XECurve.Attribute("type").Value;
                _wsCurrent.Cells[2, 6] = XECurve.Attribute("origin").Value;
                _wsCurrent.Cells[2, 4] = XECurve.Attribute("weighting").Value;
                _wsCurrent.Cells[2, 8] = XECurve.Attribute("axistrans").Value;


                try
                {
                    _wsCurrent.Cells[2, 8] = XECurve.Element("CALIBRATIONCURVE").Attribute("curve").Value;
                    _wsCurrent.Cells[2, 11] = XECurve.Element("CORRELATION").Attribute("r").Value;
                    _wsCurrent.Cells[2, 13] = XECurve.Element("DETERMINATION").Attribute("rsquared").Value;
                }
                catch (Exception)
                {
                    if (checkBox_STD.Checked)
                    {
                        if (_compoundQuanType == 1)
                        {
                            MessageBox.Show(String.Format("Calibration curve parameters for {0} unavailable", XEcompound.Attribute("name").Value));
                        }
                    }
                }



                var querySAMPLE_LIST = from sample in _QLDDoc.Descendants("SAMPLE")
                                       // orderby sample.Attribute("type").Value descending, sample.Attribute("id").Value
                                       where (sample.Attribute("type").Value == "Standard" ||
                                              sample.Attribute("type").Value == "Analyte" ||
                                              sample.Attribute("type").Value == "QC")
                                       && sample.Attribute("name").Value != ""
                                       orderby sample.Attribute("type").Value.Length descending
                                       select sample;


                int _sampleCount = querySAMPLE_LIST.Count();
                int _standardCount = querySAMPLE_LIST.Count(p => p.Attribute("type").Value == "Standard");
                int _analyteCount = querySAMPLE_LIST.Count(p => p.Attribute("type").Value == "Analyte");
                int _qcCount = querySAMPLE_LIST.Count(p => p.Attribute("type").Value == "QC");




                double _maxTP = (from _sample in querySAMPLE_LIST select _sample).Count();

                try
                {
                    _maxTP = (from _sample in querySAMPLE_LIST
                              where _sample.Attribute("type").Value == "Analyte"
                                               && double.TryParse(_sample.Attribute("sampleid").Value, out _maxTP)
                              select Convert.ToDouble(_sample.Attribute("sampleid").Value)).Max();
                }
                catch
                {
                }



                //MessageBox.Show(test.ToString());

                int index = _OutputFormat._titleShift + 1;
                foreach (XElement XEsample in querySAMPLE_LIST)
                {
                    int _sample_type_id = SampleTypeToDBInt(XEsample.Attribute("type").Value);

                    _wsCurrent.Cells[index, 1] = XEsample.Attribute("name").Value;
                    _wsCurrent.Cells[index, 2] = XEsample.Attribute("type").Value;
                    _wsCurrent.Cells[index, 3] = XEsample.Attribute("sampleid").Value;
                    _wsCurrent.Cells[index, 14] = XEsample.Attribute("subjecttext").Value;

                    // _wsCurrent.Cells[index, 20] = XEsample.Parent.Parent.Attribute("name").Value;


                    var queryCOMPOUND = from compound in XEsample.Descendants("COMPOUND")
                                        where compound.Attribute("name").Value == XEcompound.Attribute("name").Value
                                        select compound;

                    _wsCurrent.Cells[index, 4] = queryCOMPOUND.Single().Attribute("stdconc").Value;

                    var XEpeak = queryCOMPOUND.First().Element("PEAK");
                    var XEpeakis = XEpeak.Element("ISPEAK");


                    _wsCurrent.Cells[index, 6] = XEpeakis.Attribute("area").Value;
                    _wsCurrent.Cells[index, 8] = XEpeakis.Attribute("foundrt").Value;

                    if (!(XEpeak.Attribute("pkflags").Value.ToString() == "") && !(XEpeak.Attribute("pkflags").Value.ToString().Contains("-")))
                    {


                        _wsCurrent.Cells[index, 5] = XEpeak.Attribute("area").Value;
                        _wsCurrent.Cells[index, 7] = XEpeak.Attribute("foundrt").Value;
                        _wsCurrent.Cells[index, 9] = XEpeak.Attribute("response").Value;


                        _wsCurrent.Cells[index, _OutputFormat._colNum_analconc] =
                        _compoundQuanType == 1 ? XEpeak.Attribute("analconc").Value : XEpeak.Attribute("response").Value;


                        Excel.Range _xlRangeConc = _wsCurrent.Cells[index, _OutputFormat._colNum_analconc] as Excel.Range;

                        _xlRangeConc.NumberFormat =
                            _compoundQuanType == 1 ? "0.00%" : "0.00";


                        _wsCurrent.Cells[index, 11] =

                            _sample_type_id == 5 ? XEpeak.Attribute("concdevperc").Value : "";


                        _wsCurrent.Cells[index, 12] = XEpeak.Attribute("signoise").Value;
                        _wsCurrent.Cells[index, 13] = XEpeak.Attribute("chromnoise").Value;


                    }



                    //Export into DB

                    if (checkBox_ExportIntoDB.Checked)
                    {
                        string fEQUATION = "", fR = "NULL", fR_SQR = "NULL", fNUM_SLOPE = "NULL", fNUM_INTERSECT = "NULL";

                        if (XECurve.Attribute("type").Value != "RF")
                        {
                            fEQUATION = XECurve.Element("CALIBRATIONCURVE")
                                                 .Attribute("curve").Value;
                            fR = XECurve.Element("CORRELATION")
                                          .Attribute("r").Value;
                            fR_SQR = XECurve.Element("DETERMINATION")
                                              .Attribute("rsquared").Value;
                            fNUM_SLOPE = GetSlopeOfEquation(
                                         XECurve.Element("CALIBRATIONCURVE")
                                                  .Attribute("curve").Value).ToString();
                            fNUM_INTERSECT = GetInterOfEquation(
                                             XECurve.Element("CALIBRATIONCURVE")
                                                      .Attribute("curve").Value).ToString();
                        }

                        if (XEpeak.Attribute("pkflags").Value.ToString() == "" ||
                             XEsample.Attribute("type").Value == "QC" ||
                             XEpeak.Attribute("pkflags").Value.ToString().Contains("-"))
                        {
                            // MessageBox.Show(String.Format("Sample {1} on analyte {0} show no peak",
                            //                    XEcompound.Attribute("name").Value,
                            //                    XEsample.Attribute("name").Value));
                        }
                        else
                        {
                            _strExportSQL.Add("INSERT INTO TMP_XEVO_IMPORT (" +
                                                  " EXPER_FILE, ASSAY_DATE, ASSAY_TYPE_ID," +
                                                  " ANTIBODY_ID, ENZYME_ID, QUAN_TYPE_ID," +
                                                  " SAMPLE_PROCESS_BY, ASSAY_DONE_BY, DATA_PROCESS_BY," +
                                                  " FILENAME_DESC, SAMPLE_TYPE_ID, LEVEL_TP_ID," +
                                                  " NUM_SPEC_CONC, SUBJECT_NUM, FLUID_TYPE_ID," +
                                                  " ANYLYTE_NAME, AREA, HEIGHT," +
                                                  " RT, ISTD_AREA, ISTD_HEIGHT," +
                                                  " ISTD_RT, RESPONSE, CURVE," +
                                                  " WEIGHTING, ORIGIN, EQUATION," +
                                                  " R, R_SQR, NUM_SLOPE," +
                                                  " NUM_INTERSECT, NUM_ANAL_CONC, CONC_DEV_PERC," +
                                                  " SIGNOISE, CHROMNOISE, INJECTION_ID )");

                            _strExportSQL.Add(
                                    String.Format("VALUES " +
                                                     "( '{0}', '{1}', {2}, " +
                                                     "{3}, {4}, {5}, " +
                                                     "{6}, {7}, {8}, " +
                                                     "'{9}', {10}, {11}, " +
                                                     "{12}, '{13}', '{14}', " +
                                                     "'{15}', {16}, {17}, " +
                                                     "{18}, {19}, {20}, " +
                                                     "{21}, {22}, '{23}', " +
                                                     "'{24}', '{25}', '{26}', " +
                                                     "{27}, {28}, {29}, " +
                                                     "{30}, {31}, {32}, {33}, " +
                                                     "{34}, {35} );"
                                    ,
                                    ExperimentInfo._expName,                          // 0*EXPER_FILE
                                    ExperimentInfo._Date,                             // 1*ASSAY_DATE
                                    7,                                                // 2*ASSAY_TYPE_ID

                                    IPToDBInt(ExperimentInfo._IP),                    // 3*ANTIBODY_ID  
                                    EnzymeToDBInt(ExperimentInfo._Enzyme),            // 4*ENZYME_ID
                                    _compoundQuanType,// 5*INTEG_TYPE_ID

                                    ExperimentInfo._SampleProcessBy,                  // 6*SAMPLE_PROCESS_BY
                                    ExperimentInfo._DoneBy,                           // 7*ASSAY_DONE_BY
                                    ExperimentInfo._QuantitatedBy,                    // 8*DATA_PROCESS_BY

                                    XEsample.Attribute("name").Value,                 // 9*FILENAME_DESC
                                    _sample_type_id,                                  //10*SAMPLE_TYPE_ID
                                    GetLevel_TP(
                                    XEsample.Attribute("sampleid").Value,
                                        XEsample.Attribute("type").Value),             //11*LEVEL_TP_ID

                                    queryCOMPOUND.Single().Attribute("stdconc").Value, //12*NUM_SPEC_CONC
                                    XEsample.Attribute("subjecttext").Value,           //13*SUBJECT_NUM     
                                    FluidTypeToDBInt(ExperimentInfo._matrix),          //14*FLUID_TYPE_ID

                                    XEcompound.Attribute("name").Value,                //15*ANALYTE_NAME
                                    XEpeak.Attribute("area").Value,                    //16*AREA
                                    XEpeak.Attribute("height").Value,                  //17*HEIGHT

                                    XEpeak.Attribute("foundrt").Value,                 //18*RT
                                    XEpeakis.Attribute("area").Value,                  //19*ISTD AREA
                                    XEpeakis.Attribute("height").Value,                //20*ISTD HEIGHT

                                    XEpeakis.Attribute("foundrt").Value,               //21*ISTD RT
                                    XEpeak.Attribute("response").Value,                //22*RESPONSE

                                    XECurve.Attribute("type").Value,                        //23*CURVE
                                    XECurve.Attribute("weighting").Value.ToString(),        //24*WEIGHTING
                                    XECurve.Attribute("origin").Value.ToString(),           //25*ORIGIN

                                    fEQUATION,                //26*EQUATION
                                    fR,                       //27*R
                                    fR_SQR,                   //28*R_SQR
                                    fNUM_SLOPE,               //29*NUM_SLOPE
                                    fNUM_INTERSECT,           //30*NUM_INTERSECT

                                    AnalConc_EmptyCheck(
                                    _compoundQuanType == 1 ?
                                             XEpeak.Attribute("analconc").Value
                                             :
                                             XEpeak.Attribute("response").Value),      //31*NUM_ANAL_CONC

                                    _sample_type_id == 5 ?
                                             XEpeak.Attribute("concdevperc").Value
                                             :
                                             "NULL",                                   //32*CONC_DEV_PERC

                                    (XEpeak.Attribute("signoise").Value == "") ? "NULL" : XEpeak.Attribute("signoise").Value,    //33*SIGNOISE
                                    XEpeak.Attribute("chromnoise").Value,       //34*CHROMNOISE
                                    1 /*XEsample.Parent.Parent.Attribute("name").Value*/)
                                    );
                        }
                    }

                    // ProgressBar Begin
                    toolStripProgressBarMain.PerformStep();
                    //MessageBox.Show(toolStripProgressBarMain.Value.ToString());
                    toolStripStatusLabelMain.Text = "Proccesing ... " +
                        XEcompound.Attribute("name").Value + " - " + XEsample.Attribute("sampleid").Value;
                    toolStripStatusLabelMain.PerformClick();
                    // ProgressBar End

                    index++;

                }


                Excel.Range _xlRange_Std_Conc, _xlRange_Std_Response, _xlRange_Std_Dev, _xlRange_Std_AnalConc,
                            _xlRange_Analyte_SampleID, _xlRange_Analyte_AnalConc, _xlRange_Analyte_Area, _xlRange_Analyte_ISArea,
                            _xlRange_Std_Box, _xlRange_Analyte_Box, _xlRange_Sample_List_Box,
                            _xlRange_Tableau;

                //Standard Ranges****
                _xlRange_Std_Conc = _wsCurrent.get_Range
                (_OutputFormat._colLet_stdconc + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_stdconc + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Conc.Name = _wsCurrent.Name + "_Std_Conc";

                _xlRange_Std_Response = _wsCurrent.get_Range
                (_OutputFormat._colLet_response + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_response + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Response.Name = _wsCurrent.Name + "_Std_Response";

                _xlRange_Std_AnalConc = _wsCurrent.get_Range
                   (_OutputFormat._colLet_analconc + (_OutputFormat._titleShift + 1).ToString(),
                    _OutputFormat._colLet_analconc + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_AnalConc.Name = _wsCurrent.Name + "_Std_AnalConc";

                _xlRange_Std_Dev = _wsCurrent.get_Range
                (_OutputFormat._colLet_stddev + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_stddev + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Dev.Name = _wsCurrent.Name + "_Std_Deviation";

                _xlRange_Std_Box = _wsCurrent.get_Range
                (_OutputFormat._colLet_First + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Box.Name = _wsCurrent.Name + "_Std_Box";



                //Analyte Ranges****
                _xlRange_Analyte_SampleID = _wsCurrent.get_Range
             (_OutputFormat._colLet_sampleid + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
              _OutputFormat._colLet_sampleid + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_SampleID.Name = _wsCurrent.Name + "_Analyte_SampleID";

                _xlRange_Analyte_AnalConc = _wsCurrent.get_Range
             (_OutputFormat._colLet_analconc + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
              _OutputFormat._colLet_analconc + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_AnalConc.Name = _wsCurrent.Name + "_Analyte_AnalConc";

                _xlRange_Analyte_Area = _wsCurrent.get_Range
              (_OutputFormat._colLet_analarea + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
               _OutputFormat._colLet_analarea + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_Area.Name = _wsCurrent.Name + "_Analyte_Area";

                _xlRange_Analyte_ISArea = _wsCurrent.get_Range
              (_OutputFormat._colLet_analISarea + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
               _OutputFormat._colLet_analISarea + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_ISArea.Name = _wsCurrent.Name + "_Analyte_ISArea";

                _xlRange_Analyte_Box = _wsCurrent.get_Range
             (_OutputFormat._colLet_First + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
              _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_Box.Name = _wsCurrent.Name + "_Analyte_Box";


                //SAMPLE_LIST Range
                _xlRange_Sample_List_Box = _wsCurrent.get_Range
              (_OutputFormat._colLet_First + (_OutputFormat._titleShift + 1).ToString(),
            _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Sample_List_Box.Name = _wsCurrent.Name + "_Sample_Box";

                //Tableau Range
                _xlRange_Tableau = _wsCurrent.get_Range
              (_OutputFormat._colLet_First + (_OutputFormat._titleShift).ToString(),
            _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Tableau.Name = _wsCurrent.Name + "_Tableau";



                //Format Ranges**** 
                _xlRange_Sample_List_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                _xlRange_Sample_List_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).Weight = 3.75;
                _xlRange_Sample_List_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).Weight = 3.75;
                _xlRange_Std_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                _xlRange_Std_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).Weight = 1.75;
                _xlRange_Analyte_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).Weight = 1.75;

                _xlRange_Std_Box.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);

                //////////////////////////////////////////////////////
                Excel.ColorScale cfColorScale = (Excel.ColorScale)(_xlRange_Analyte_ISArea.FormatConditions.AddColorScale(3));
                // Set the minimum threshold to red (0x000000FF) and maximum threshold
                // to blue (0x00FF0000).
                cfColorScale.ColorScaleCriteria[1].FormatColor.Color = 7039480;  //0x000000FF;  
                cfColorScale.ColorScaleCriteria[2].FormatColor.Color = 8711167;  // 0x00FF0000;
                cfColorScale.ColorScaleCriteria[3].FormatColor.Color = 13011546; // 0x00FF0000;
                /////////////////////////////////////////////////////



                //Plots****
                //Time Serie****



                double _rowHeight = (double)_wsCurrent.get_Range("A1", "A1").RowHeight;

                double _graphOffset = (double)_xlRange_Sample_List_Box.Top + (double)_xlRange_Sample_List_Box.Height +
                       (_rowHeight + _rowHeight * _qcCount);




                Excel.ChartObjects _xlCharts = (Excel.ChartObjects)_wsCurrent.ChartObjects(Type.Missing);

                Excel.ChartObject _xlChartTimeCorse;

                string _serieCurveName = "default curve name";

                string _axeCurveXNumFormat = "0", _axeCurveYNumFormat = "0.0%";

                string _axeCurveXCaption = "default XCaption", _axeCurveYCaption = "default YCaption";

                //_xlChartTimeCorse.Width = (double)_xlRange_Tableau.Width;    

                _xlChartTimeCorse = (Excel.ChartObject)_xlCharts.Add(
                              0, _graphOffset, (double)_xlRange_Tableau.Width, 250);

                Excel.Chart _xlChartTimeCorsePage = _xlChartTimeCorse.Chart;
                _xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlXYScatter;
                _xlChartTimeCorsePage.HasTitle = true; _xlChartTimeCorsePage.ChartTitle.Font.Size = 11;


                _xlChartTimeCorsePage.ChartTitle.Caption = XEcompound.Attribute("name").Value + " " + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_TP = (Excel.Axes)_xlChartTimeCorsePage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_TP.HasMajorGridlines = true;
                _xlAxisX_TP.HasTitle = true; _xlAxisY_TP.HasTitle = true;

                _xlAxisX_TP.MinimumScale = 0; _xlAxisX_TP.MaximumScale = _maxTP;
                _xlAxisY_TP.MinimumScale = 0;

                _xlChartTimeCorsePage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartTimeCorsePage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_TimeCourse = (Excel.SeriesCollection)_xlChartTimeCorsePage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_TTR = (Excel.Series)_xlSeriesColl_TimeCourse.NewSeries();
                _xlSeries_TTR.XValues = _xlRange_Analyte_SampleID.Cells;
                _xlSeries_TTR.Values = _xlRange_Analyte_AnalConc;

                _xlSeries_TTR.ChartType = Excel.XlChartType.xlXYScatterSmooth;

                if (_compoundQuanType == 1)
                {
                    _serieCurveName = "H:L Abeta";

                    _axeCurveXNumFormat = "0"; _axeCurveYNumFormat = "0.0%";

                    _axeCurveXCaption = "h, Time"; _axeCurveYCaption = "TTR";

                    _xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlXYScatter;

                    _xlAxisX_TP.TickLabels.NumberFormat = _axeCurveXNumFormat;

                    _xlAxisX_TP.AxisTitle.Caption = _axeCurveXCaption;

                }
                else
                {
                    _axeCurveXCaption = "h, Time";
                    _axeCurveYNumFormat = "0.0";
                    _axeCurveYCaption = "Response";
                    _serieCurveName = "Abeta level";

                    _xlAxisX_TP.AxisTitle.Caption = _axeCurveXCaption;

                    _xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlColumnStacked;

                }



                _xlAxisY_TP.TickLabels.NumberFormat = _axeCurveYNumFormat;
                _xlAxisY_TP.AxisTitle.Caption = _axeCurveYCaption;

                _xlSeries_TTR.Name = _serieCurveName;
                _xlSeries_TTR.Name = XEcompound.Attribute("name").Value;

                if (_standardCount == 0)
                {
                    _axeCurveYCaption = "Area Ratio";

                }

                //-----------------------------------------------------------------------------------------------------------------------------------

                ///*********************************************************************************************************************

                _graphOffset += _xlChartTimeCorse.Height;

                Excel.ChartObject _xlChart_Area = (Excel.ChartObject)_xlCharts.Add(
                                  0, _graphOffset, (double)_xlRange_Tableau.Width, 250);
                Excel.Chart _xlChart_AreaPage = _xlChart_Area.Chart;

                _xlChart_AreaPage.ChartType = Excel.XlChartType.xlColumnClustered;
                _xlChart_AreaPage.HasTitle = true;
                _xlChart_AreaPage.ChartTitle.Font.Size = 11;

                _xlChart_AreaPage.ChartTitle.Caption = XEcompound.Attribute("name").Value + " " + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_Area = (Excel.Axes)_xlChart_AreaPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_Area = _xlAxes_Area.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_Area = _xlAxes_Area.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_Area.HasMajorGridlines = true;
                _xlAxisX_Area.HasTitle = true;
                _xlAxisY_Area.HasTitle = true;

                _xlAxisY_Area.MinimumScale = 0;

                _xlChart_AreaPage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChart_AreaPage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_Area = (Excel.SeriesCollection)_xlChart_AreaPage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_Area, _xlSeries_ISTD_Area;

                _xlSeries_Area = (Excel.Series)_xlSeriesColl_Area.NewSeries();
                _xlSeries_ISTD_Area = (Excel.Series)_xlSeriesColl_Area.NewSeries();

                if (_compoundQuanType != 1)
                {

                    _xlSeries_Area.XValues = _xlRange_Analyte_SampleID.Cells;
                    _xlSeries_Area.Values = _xlRange_Analyte_Area.Cells;
                }

                _xlSeries_ISTD_Area.XValues = _xlRange_Analyte_SampleID.Cells;
                _xlSeries_ISTD_Area.Values = _xlRange_Analyte_ISArea.Cells;

                _axeCurveYNumFormat = "0.0E+0"; _axeCurveYCaption = "Area";

                _xlAxisY_Area.TickLabels.NumberFormat = _axeCurveYNumFormat;
                _xlAxisX_Area.AxisTitle.Caption = _axeCurveXCaption;
                _xlAxisY_Area.AxisTitle.Caption = _axeCurveYCaption;

                _xlSeries_Area.Name = "Area";
                _xlSeries_Area.Name = "ISTD Area";


                ///*********************************************************************************************************************
                ///*********************************************************************************************************************
                //RT
                _graphOffset += _xlChart_Area.Height;

                Excel.ChartObject _xlChartRT = (Excel.ChartObject)_xlCharts.Add(
                                  0, _graphOffset, (double)_xlRange_Tableau.Width, 250);
                Excel.Chart _xlChartRTPage = _xlChartRT.Chart;

                _xlChartRTPage.ChartType = Excel.XlChartType.xlLineMarkers;
                _xlChartRTPage.HasTitle = true;
                _xlChartRTPage.ChartTitle.Font.Size = 11;

                _xlChartRTPage.ChartTitle.Caption = XEcompound.Attribute("name").Value + " " + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_RT = (Excel.Axes)_xlChartRTPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_RT = _xlAxes_RT.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_RT = _xlAxes_RT.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_RT.HasMajorGridlines = true;
                _xlAxisX_RT.HasTitle = true; _xlAxisY_RT.HasTitle = true;

                _xlChartRTPage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartRTPage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_RT = (Excel.SeriesCollection)_xlChartRTPage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_RT, _xlSeries_ISTD_RT;
                _xlSeries_RT = (Excel.Series)_xlSeriesColl_RT.NewSeries();
                _xlSeries_RT.XValues = _xlRange_Analyte_SampleID.Cells;
                _xlSeries_RT.Values = _xlRange_Analyte_AnalConc.get_Offset(0, -3).Cells; ;

                _xlSeries_ISTD_RT = (Excel.Series)_xlSeriesColl_RT.NewSeries();
                _xlSeries_ISTD_RT.XValues = _xlRange_Analyte_SampleID.Cells;
                _xlSeries_ISTD_RT.Values = _xlRange_Analyte_AnalConc.get_Offset(0, -2).Cells; ;

                _axeCurveYNumFormat = "0.00";

                _axeCurveYCaption = "RT";


                _xlAxisY_RT.TickLabels.NumberFormat = _axeCurveYNumFormat;
                _xlAxisX_RT.AxisTitle.Caption = _axeCurveXCaption;
                _xlAxisY_RT.AxisTitle.Caption = _axeCurveYCaption;

                _xlSeries_RT.Name = "RT";
                _xlSeries_ISTD_RT.Name = "ISTD RT";


                ///*********************************************************************************************************************   
                //-----------------------------------------------------------------------------------------------------------------------------------
                //Std Curves***
                Excel.ChartObject _xlChartStdCurve;

                _graphOffset += +(double)_xlChartRT.Height;

                if (_compoundQuanType == 1)
                {
                    _xlChartStdCurve = (Excel.ChartObject)_xlCharts.Add(
                                       (double)_xlRange_Tableau.Width - 470, _graphOffset, 470, 250);
                    Excel.Chart _xlChartStdCurvePage = _xlChartStdCurve.Chart;
                    _xlChartStdCurvePage.ChartType = Excel.XlChartType.xlXYScatter;

                    _xlChartStdCurvePage.HasTitle = true;
                    _xlChartStdCurvePage.ChartTitle.Font.Size = 11;

                    string _chartTitleStdCurve = String.Format
                        ("Calibration curve {0} {1} {2} {3} ({4})",
                        XEcompound.Attribute("name").Value,
                        ExperimentInfo._IP,
                        ExperimentInfo._Enzyme,
                        ExperimentInfo._Instument,
                        ExperimentInfo._Date);

                    _xlChartStdCurvePage.ChartTitle.Caption = _chartTitleStdCurve;

                    Excel.Axes _xlAxes_Std = (Excel.Axes)_xlChartStdCurvePage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisX_Std = _xlAxes_Std.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisY_Std = _xlAxes_Std.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                    _xlAxisX_Std.HasTitle = true; _xlAxisY_Std.HasTitle = true;
                    _xlAxisX_Std.HasMajorGridlines = true;
                    _xlAxisY_Std.HasMajorGridlines = true;
                    _xlAxisY_Std.MinimumScale = 0;

                    _xlChartStdCurvePage.Legend.IncludeInLayout = true;
                    _xlChartStdCurvePage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionTop;

                    Excel.SeriesCollection _xlSerColl_StdCurve = (Excel.SeriesCollection)_xlChartStdCurvePage.SeriesCollection(Type.Missing);

                    Excel.Series _xlSerie_Stds = (Excel.Series)_xlSerColl_StdCurve.NewSeries();
                    _xlSerie_Stds.XValues = _xlRange_Std_Conc.Cells;
                    _xlSerie_Stds.Values = _xlRange_Std_Response.Cells;

                    string _serieStdName = "H:L Abeta";

                    string _axeStdXNumFormat = "0.00",
                           _axeStdYNumFormat = "0.00";

                    string _axeStdXCaption = "%, Specified amount",
                           _axeStdYCaption = "%, Response";


                    _xlAxisX_Std.TickLabels.NumberFormat = _axeStdXNumFormat;
                    _xlAxisY_Std.TickLabels.NumberFormat = _axeStdYNumFormat;
                    _xlAxisX_Std.AxisTitle.Caption = _axeStdXCaption;
                    _xlAxisY_Std.AxisTitle.Caption = _axeStdYCaption;

                    _xlSerie_Stds.Name = _serieStdName;

                    Excel.Trendlines _xlTrendlines_Stds = (Excel.Trendlines)_xlSerie_Stds.Trendlines(Type.Missing);

                    Excel.XlTrendlineType _XlTrendlineType =
                        Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear;
                    string _CurveIndex = XECurve.Attribute("type").Value;
                    //_wsComponent.get_Range("C2", "C2").Value2.ToString();

                    if (_CurveIndex == "Quadratic")
                    { _XlTrendlineType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlExponential; }


                    _xlTrendlines_Stds.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, true, true, Type.Missing);


                }

                //End Plots***

                //****************************
                Excel.ChartObjects _xlChartsSummary = (Excel.ChartObjects)_wsSummary.ChartObjects(Type.Missing);
                Excel.ChartObject _xlChartCurveSummary, _xlChartLevelSummary;
                Excel.Chart _xlChartCurveSummaryPage, _xlChartLevelSummaryPage;

                if (_xlChartsSummary.Count < 1)
                {
                    _axeCurveYCaption = "TTR, %";
                    _axeCurveYNumFormat = "0.00%";

                    _xlChartCurveSummary = (Excel.ChartObject)_xlChartsSummary.Add(
                                   10, 10, 850, 500);
                    _xlChartCurveSummaryPage = _xlChartCurveSummary.Chart;
                    _xlChartCurveSummaryPage.ChartType = Excel.XlChartType.xlXYScatterSmooth;
                    _xlChartCurveSummaryPage.HasTitle = true;
                    _xlChartCurveSummaryPage.ChartTitle.Font.Size = 11;

                    Excel.Axes _xlAxes_CurveSummary = (Excel.Axes)_xlChartCurveSummaryPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisX_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisY_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                    _xlAxisX_CurveSummary.HasMajorGridlines = true; _xlAxisY_CurveSummary.HasMajorGridlines = true;

                    _xlAxisX_CurveSummary.HasTitle = true; _xlAxisY_CurveSummary.HasTitle = true;

                    _xlAxisX_CurveSummary.MaximumScale = _maxTP;

                    _xlAxisX_CurveSummary.AxisTitle.Caption = _axeCurveXCaption; _xlAxisY_CurveSummary.AxisTitle.Caption = _axeCurveYCaption;

                    _xlAxisX_CurveSummary.TickLabels.NumberFormat = _axeCurveXNumFormat; _xlAxisY_CurveSummary.TickLabels.NumberFormat = _axeCurveYNumFormat;

                    _xlChartCurveSummaryPage.ChartTitle.Caption = "TTR Abeta in ";

                    _xlChartCurveSummaryPage.ChartTitle.Caption += ExperimentInfo._CommonChartNameStr;

                    _xlChartCurveSummaryPage.Legend.Position = Excel.XlLegendPosition.xlLegendPositionRight;
                    _xlChartCurveSummaryPage.Legend.IncludeInLayout = true;

                    //*************************************************************************************************
                    _axeCurveYCaption = "Response";
                    _axeCurveYNumFormat = "0.0";

                    _xlChartLevelSummary = (Excel.ChartObject)_xlChartsSummary.Add(
                                   10, 510, 850, 500);
                    _xlChartLevelSummaryPage = _xlChartLevelSummary.Chart;
                    _xlChartLevelSummaryPage.ChartType = Excel.XlChartType.xlXYScatter;
                    _xlChartLevelSummaryPage.HasTitle = true;
                    _xlChartLevelSummaryPage.ChartTitle.Font.Size = 11;

                    Excel.Axes _xlAxes_LevelSummary = (Excel.Axes)_xlChartLevelSummaryPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisX_LevelSummary = _xlAxes_LevelSummary.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisY_LevelSummary = _xlAxes_LevelSummary.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                    _xlAxisX_LevelSummary.HasMajorGridlines = true; _xlAxisY_LevelSummary.HasMajorGridlines = true;

                    _xlAxisX_LevelSummary.HasTitle = true; _xlAxisY_LevelSummary.HasTitle = true;

                    _xlAxisX_LevelSummary.MaximumScale = _maxTP;

                    _xlAxisX_LevelSummary.AxisTitle.Caption = _axeCurveXCaption; _xlAxisY_LevelSummary.AxisTitle.Caption = _axeCurveYCaption;

                    _xlAxisX_LevelSummary.TickLabels.NumberFormat = _axeCurveXNumFormat; _xlAxisY_LevelSummary.TickLabels.NumberFormat = _axeCurveYNumFormat;

                    _xlChartLevelSummaryPage.ChartTitle.Caption = "Abeta levels in ";

                    _xlChartLevelSummaryPage.ChartTitle.Caption += ExperimentInfo._CommonChartNameStr;

                    _xlChartLevelSummaryPage.Legend.Position = Excel.XlLegendPosition.xlLegendPositionRight;
                    _xlChartLevelSummaryPage.Legend.IncludeInLayout = true;

                }
                else
                {
                    _xlChartCurveSummary = (Excel.ChartObject)_xlChartsSummary.Item(1);
                    _xlChartCurveSummaryPage = _xlChartCurveSummary.Chart;
                    _xlChartLevelSummary = (Excel.ChartObject)_xlChartsSummary.Item(2);
                    _xlChartLevelSummaryPage = _xlChartLevelSummary.Chart;
                }


                Excel.SeriesCollection _xlSeriesColl_Summary =
                    _compoundQuanType == 1 ?
                                            (Excel.SeriesCollection)_xlChartCurveSummaryPage.SeriesCollection(Type.Missing)
                                            :
                                            (Excel.SeriesCollection)_xlChartLevelSummaryPage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_timeCourse = (Excel.Series)_xlSeriesColl_Summary.NewSeries();
                _xlSeries_timeCourse.XValues = _xlRange_Analyte_SampleID.Cells;
                _xlSeries_timeCourse.Values = _xlRange_Analyte_AnalConc;
                _xlSeries_timeCourse.Name = XEcompound.Attribute("name").Value;

            }

            toolStripProgressBarMain.Visible = false; toolStripProgressBarMain.Value = 0;

            toolStripStatusLabelMain.Text = "Done"; toolStripStatusLabelMain.PerformClick();


            string _AppPath = Application.StartupPath + Path.DirectorySeparatorChar;
            toolStripStatusLabelMain.Text = "Exporting data into DB ...";
            toolStripStatusLabelMain.PerformClick();
            string[] _sql_Part1 = File.ReadAllLines(_AppPath + "sql_XevoImport_P1.sql");
            string[] _sql_Part2 = File.ReadAllLines(_AppPath + "sql_XevoImport_P2.sql");
            string[] _sql_Part3 = File.ReadAllLines(_AppPath + "sql_XevoImport_P3.sql");

            _strExportSQL.Insert(0, " ");
            _strExportSQL.InsertRange(0, _sql_Part1.ToList<string>());

            _strExportSQL.Insert(0, "/*");

            _strExportSQL.Insert(1, String.Format("Generated by QLD Reader {0} from {1}", Assembly.GetExecutingAssembly().GetName().Version.ToString(), Convert.ToDateTime(Properties.Resources.BuildDate).ToShortDateString()));
            _strExportSQL.Insert(2, String.Format("       Machine: {0} ran by {1}@{2}", Environment.MachineName, Environment.UserName, Environment.UserDomainName));
            _strExportSQL.Insert(3, String.Format("     Timestamp: {0} @ {1}", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString()));
            _strExportSQL.Insert(4, String.Format("        Source: {0} ", textBox_QLDfile.Text));

            _strExportSQL.Insert(5, "*/");



            _strExportSQL.AddRange(_sql_Part3.ToList<string>());

            File.WriteAllLines(Path.ChangeExtension(textBox_Excelfile.Text, "sql"), _strExportSQL.ToArray());

            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();



            object misValue = System.Reflection.Missing.Value;

            try
            {
                xlWB.SaveAs(textBox_Excelfile.Text,
                   misValue,
                   misValue, misValue, misValue, misValue,
                   Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlUserResolution,
                   misValue, misValue, misValue, misValue);
            }

            catch (COMException _exception)
            {
                if (_exception.ErrorCode == unchecked((int)0x800A03EC))
                {
                    MessageBox.Show("Time stamp has been added to file name you specified", this.Text,
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    string _newExcelFileName = textBox_Excelfile.Text.Replace(
                        Path.GetFileNameWithoutExtension(textBox_Excelfile.Text),
                        Path.GetFileNameWithoutExtension(textBox_Excelfile.Text)
                        + "_created_"
                        + System.DateTime.Now.ToString("s")
                        .Replace("/", "")
                        .Replace(":", "")
                        .Replace(" ", "_")
                        );

                    xlWB.SaveAs(_newExcelFileName,
                       misValue,
                       misValue, misValue, misValue, misValue,
                       Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlUserResolution,
                       misValue, misValue, misValue, misValue);
                }
            }

            finally
            {
                if (this.checkBox_OpenExcel.Checked)
                {
                    _xlApp.Visible = true;
                    // _wsSummary.Activate();

                }
                else
                {
                    _xlApp.Quit();
                }
            }

            if (checkBox_ShowDebug.Checked)
            {
                MessageBox.Show("ExportIVBolus_csf()");
            }

        }

        private void ExportSleepStudy_csf()
        {
            List<string> _strExportSQL = new List<string>();

            XDocument _QLDDoc;
            try
            {
                _QLDDoc = XDocument.Load(textBox_QLDfile.Text);
            }
            catch (Exception _XMLExeption)
            {
                MessageBox.Show(_XMLExeption.Message);
                return;
            }

            toolStripProgressBarMain.Value = 0;

            
            toolStripProgressBarMain.Visible = true;
            toolStripStatusLabelMain.Visible = true;

            

            _xlApp = new Excel.Application();
            _xlApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
            _xlApp.ScreenUpdating = false;

            Excel.Workbook xlWB;
                                  

            try
            {
                xlWB =
                    _xlApp.Workbooks.Add(
                                         Application.StartupPath + Path.DirectorySeparatorChar +
                                         @"\Templates\SleepStudy_CSF_Xevo_Template_Summary.xlsx"
                );
            }

            catch (Exception)
            {
                MessageBox.Show("File: " + Application.StartupPath + Path.DirectorySeparatorChar +
                @"\Templates\SleepStudy_CSF_Xevo_Template_Summary.xlsx", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //MessageBox.Show(_exception.Message, _exception.GetBaseException().Message);
                return;
            }

            finally
            { }

            //MessageBox.Show(ExperimentInfo._quanType);

           

            string _quanTypeKey = "";

            if (ExperimentInfo._quanType == "RELATIVE")
            {
                _quanTypeKey = "C13N14";
            }
            if (ExperimentInfo._quanType == "ABSOLUTE")
            {
                _quanTypeKey = "C12N14";
            }
            if (ExperimentInfo._quanType == "REL+ABS")
            {
                _quanTypeKey = "N14";
            }
                       
            Excel.Worksheet _wsSummary = (Excel.Worksheet)xlWB.ActiveSheet;

            Excel.Worksheet _wsCurrent;
            
            var queryCOMPOUND_LIST = from compound in _QLDDoc.Descendants("CALIBRATIONDATA").Descendants("COMPOUND")
                                     where compound.Attribute("name").Value.Contains(_quanTypeKey)

                                     //  &&
                                     //  !compound.Attribute("name").Value.Contains("Ch")

                                     // select only group 1   //2013-05-03 VO
                                     //&& compound.Parent.Parent.Attribute("id").Value == "1"

                                     select compound;

            toolStripProgressBarMain.Maximum =
               Convert.ToInt32(_QLDDoc.Descendants("SAMPLELISTDATA").First().Attribute("count").Value) *
                               queryCOMPOUND_LIST.Count();

           

            foreach (XElement XEcompound in queryCOMPOUND_LIST)
            {

                int _compoundQuanType = QuntTypeToDBInt(XEcompound.Attribute("name").Value);

                if (XEcompound.Parent.Parent.Attribute("name").Value == "ABS" && _compoundQuanType == 1) { continue; }
                if (XEcompound.Parent.Parent.Attribute("name").Value == "REL" && _compoundQuanType == 2) { continue; }


                _wsCurrent = (Excel.Worksheet)xlWB.Sheets.Add
                          (_wsSummary, Type.Missing, Type.Missing,
                          Application.StartupPath + Path.DirectorySeparatorChar +
                            @"\Templates\SleepStudy_CSF_Xevo_Template.xlsx");
                string _sheetName = XEcompound.Attribute("name").Value;


                _wsCurrent.Cells[2, 1] = _sheetName;

                if (_sheetName.Length > 31)
                {
                    _sheetName = "alt_" + _sheetName.Remove(1, _sheetName.Length - 31 + 4);
                }


                _wsCurrent.Name = _sheetName.Replace("+", "_");


                var XECurve = XEcompound.Element("CURVE");
                _wsCurrent.Cells[2, 3] = XECurve.Attribute("type").Value;
                _wsCurrent.Cells[2, 6] = XECurve.Attribute("origin").Value;
                _wsCurrent.Cells[2, 4] = XECurve.Attribute("weighting").Value;
                _wsCurrent.Cells[2, 8] = XECurve.Attribute("axistrans").Value;

                try
                {
                    _wsCurrent.Cells[2, 8] = XECurve.Element("CALIBRATIONCURVE").Attribute("curve").Value;
                    _wsCurrent.Cells[2, 11] = XECurve.Element("CORRELATION").Attribute("r").Value;
                    _wsCurrent.Cells[2, 13] = XECurve.Element("DETERMINATION").Attribute("rsquared").Value;
                }
                catch (Exception)
                {
                    if (checkBox_STD.Checked)
                    {
                        if (_compoundQuanType == 1)
                        {
                            MessageBox.Show(String.Format("Calibration curve parameters for {0} unavailable", XEcompound.Attribute("name").Value));
                        }
                    }
                }

                var querySAMPLE_LIST_R = from sample in _QLDDoc.Descendants("SAMPLE")
                                         where (sample.Attribute("type").Value == "Standard" ||
                                                sample.Attribute("type").Value == "Analyte" ||
                                                sample.Attribute("type").Value == "QC") &&
                                                sample.Parent.Parent.Attribute("name").Value == "REL" &&
                                                sample.Attribute("name").Value != ""
                                         orderby sample.Attribute("type").Value.Length descending
                                         select sample;

                var querySAMPLE_LIST_A = from sample in _QLDDoc.Descendants("SAMPLE")
                                         where (sample.Attribute("type").Value == "Standard" ||
                                                sample.Attribute("type").Value == "Analyte" ||
                                                sample.Attribute("type").Value == "QC") &&
                                                sample.Parent.Parent.Attribute("name").Value == "ABS" &&
                                                sample.Attribute("name").Value != ""
                                         orderby sample.Attribute("type").Value.Length descending
                                         select sample;


                int _sampleCount_R = querySAMPLE_LIST_R.Count(), _sampleCount_A = querySAMPLE_LIST_A.Count();
                
                int _standardCount_R = querySAMPLE_LIST_R.Count(p => p.Attribute("type").Value == "Standard"),
                    _standardCount_A = querySAMPLE_LIST_A.Count(p => p.Attribute("type").Value == "Standard");

                int _analyteCount_R = querySAMPLE_LIST_R.Count(p => p.Attribute("type").Value == "Analyte"),
                    _analyteCount_A = querySAMPLE_LIST_A.Count(p => p.Attribute("type").Value == "Analyte");

                int _qcCount_R = querySAMPLE_LIST_R.Count(p => p.Attribute("type").Value == "QC"),
                    _qcCount_A = querySAMPLE_LIST_A.Count(p => p.Attribute("type").Value == "QC");


                double _maxTP = (from _sample in querySAMPLE_LIST_R select _sample).Count();

                try
                {
                    _maxTP = (from _sample in querySAMPLE_LIST_R
                              where _sample.Attribute("type").Value == "Analyte"
                                               && double.TryParse(_sample.Attribute("sampleid").Value, out _maxTP)
                              select Convert.ToDouble(_sample.Attribute("sampleid").Value)).Max();
                }
                catch
                {
                }


                int index = _OutputFormat._titleShift + 1,
                    index_R = index, index_A = index;

                foreach (XElement XEsample in querySAMPLE_LIST_R.Concat(querySAMPLE_LIST_A))
                {
                    if (XEsample.Parent.Parent.Attribute("name").Value == "ABS" && _compoundQuanType == 1) { continue; }
                    if (XEsample.Parent.Parent.Attribute("name").Value == "REL" && _compoundQuanType == 2) { continue; }

                    int _sample_type_id = SampleTypeToDBInt(XEsample.Attribute("type").Value);

                    _wsCurrent.Cells[index, 1] = XEsample.Attribute("name").Value;
                    _wsCurrent.Cells[index, 2] = XEsample.Attribute("type").Value;
                    _wsCurrent.Cells[index, 3] = XEsample.Attribute("sampleid").Value;
                    _wsCurrent.Cells[index, 14] = XEsample.Attribute("subjecttext").Value;

                    var queryCOMPOUND = from compound in XEsample.Descendants("COMPOUND")
                                        where compound.Attribute("name").Value == XEcompound.Attribute("name").Value
                                        select compound;

                    _wsCurrent.Cells[index, 4] = queryCOMPOUND.Single().Attribute("stdconc").Value;

                    var XEpeak = queryCOMPOUND.First().Element("PEAK");
                    var XEpeakis = XEpeak.Element("ISPEAK");


                    _wsCurrent.Cells[index, 6] = XEpeakis.Attribute("area").Value;
                    _wsCurrent.Cells[index, 8] = XEpeakis.Attribute("foundrt").Value;

                    if (!(XEpeak.Attribute("pkflags").Value.ToString() == "") && !(XEpeak.Attribute("pkflags").Value.ToString().Contains("-")))
                    {
                        _wsCurrent.Cells[index, 5] = XEpeak.Attribute("area").Value;
                        _wsCurrent.Cells[index, 7] = XEpeak.Attribute("foundrt").Value;
                        _wsCurrent.Cells[index, 9] = XEpeak.Attribute("response").Value;

                        _wsCurrent.Cells[index, _OutputFormat._colNum_analconc] =
                        _compoundQuanType == 1 ? XEpeak.Attribute("analconc").Value : XEpeak.Attribute("analconc").Value;

                        Excel.Range _xlRangeConc = _wsCurrent.Cells[index, _OutputFormat._colNum_analconc] as Excel.Range;

                        _xlRangeConc.NumberFormat =
                            _compoundQuanType == 1 ? "0.00%" : "0.00";

                        _wsCurrent.Cells[index, 11] =

                            _sample_type_id == 5 ? XEpeak.Attribute("concdevperc").Value : "";

                        _wsCurrent.Cells[index, 12] = XEpeak.Attribute("signoise").Value;
                        _wsCurrent.Cells[index, 13] = XEpeak.Attribute("chromnoise").Value;
                    }
                   
                    //Export into DB

                    if (checkBox_ExportIntoDB.Checked)
                    {
                        string fEQUATION = "", fR = "NULL", fR_SQR = "NULL", fNUM_SLOPE = "NULL", fNUM_INTERSECT = "NULL";

                        if (XECurve.Attribute("type").Value != "RF")
                        {
                            fEQUATION = XECurve.Element("CALIBRATIONCURVE")
                                                 .Attribute("curve").Value;
                            fR = XECurve.Element("CORRELATION")
                                          .Attribute("r").Value;
                            fR_SQR = XECurve.Element("DETERMINATION")
                                              .Attribute("rsquared").Value;
                            fNUM_SLOPE = GetSlopeOfEquation(
                                         XECurve.Element("CALIBRATIONCURVE")
                                                  .Attribute("curve").Value).ToString();
                            fNUM_INTERSECT = GetInterOfEquation(
                                             XECurve.Element("CALIBRATIONCURVE")
                                                      .Attribute("curve").Value).ToString();
                        }

                        if (XEpeak.Attribute("pkflags").Value.ToString() == "" ||
                             XEsample.Attribute("type").Value == "QC" ||
                             XEpeak.Attribute("pkflags").Value.ToString().Contains("-"))
                        {
                            // MessageBox.Show(String.Format("Sample {1} on analyte {0} show no peak",
                            //                    XEcompound.Attribute("name").Value,
                            //                    XEsample.Attribute("name").Value));
                        }
                        else
                        {
                            _strExportSQL.Add("INSERT INTO TMP_XEVO_IMPORT (" +
                                                  " EXPER_FILE, ASSAY_DATE, ASSAY_TYPE_ID," +
                                                  " ANTIBODY_ID, ENZYME_ID, QUAN_TYPE_ID," +
                                                  " SAMPLE_PROCESS_BY, ASSAY_DONE_BY, DATA_PROCESS_BY," +
                                                  " FILENAME_DESC, SAMPLE_TYPE_ID, LEVEL_TP_ID," +
                                                  " NUM_SPEC_CONC, SUBJECT_NUM, FLUID_TYPE_ID," +
                                                  " ANYLYTE_NAME, AREA, HEIGHT," +
                                                  " RT, ISTD_AREA, ISTD_HEIGHT," +
                                                  " ISTD_RT, RESPONSE, CURVE," +
                                                  " WEIGHTING, ORIGIN, EQUATION," +
                                                  " R, R_SQR, NUM_SLOPE," +
                                                  " NUM_INTERSECT, NUM_ANAL_CONC, CONC_DEV_PERC," +
                                                  " SIGNOISE, CHROMNOISE, INJECTION_ID )");

                            _strExportSQL.Add(
                                    String.Format("VALUES " +
                                                     "( '{0}', '{1}', {2}, " +
                                                     "{3}, {4}, {5}, " +
                                                     "{6}, {7}, {8}, " +
                                                     "'{9}', {10}, {11}, " +
                                                     "{12}, '{13}', '{14}', " +
                                                     "'{15}', {16}, {17}, " +
                                                     "{18}, {19}, {20}, " +
                                                     "{21}, {22}, '{23}', " +
                                                     "'{24}', '{25}', '{26}', " +
                                                     "{27}, {28}, {29}, " +
                                                     "{30}, {31}, {32}, {33}, " +
                                                     "{34}, {35} );"
                                    ,
                                    ExperimentInfo._expName,                          // 0*EXPER_FILE
                                    ExperimentInfo._Date,                             // 1*ASSAY_DATE
                                    7,                                                // 2*ASSAY_TYPE_ID
                                    ExperimentInfo._AbodyID,                          // 3*ANTIBODY_ID  
                                    ExperimentInfo._EnzymeID,                         // 4*ENZYME_ID
                                    _compoundQuanType,                                // 5*INTEG_TYPE_ID

                                    ExperimentInfo._SampleProcessBy,                  // 6*SAMPLE_PROCESS_BY
                                    ExperimentInfo._DoneBy,                           // 7*ASSAY_DONE_BY
                                    ExperimentInfo._QuantitatedBy,                    // 8*DATA_PROCESS_BY

                                    XEsample.Attribute("name").Value,                 // 9*FILENAME_DESC
                                    _sample_type_id,                                  //10*SAMPLE_TYPE_ID
                                    GetLevel_TP(
                                    XEsample.Attribute("sampleid").Value,
                                        XEsample.Attribute("type").Value),             //11*LEVEL_TP_ID

                                    queryCOMPOUND.Single().Attribute("stdconc").Value, //12*NUM_SPEC_CONC
                                    XEsample.Attribute("subjecttext").Value,           //13*SUBJECT_NUM     
                                    ExperimentInfo._matrixID,                         //14*FLUID_TYPE_ID

                                    XEcompound.Attribute("name").Value,                //15*ANALYTE_NAME
                                    XEpeak.Attribute("area").Value,                    //16*AREA
                                    XEpeak.Attribute("height").Value,                  //17*HEIGHT

                                    XEpeak.Attribute("foundrt").Value,                 //18*RT
                                    XEpeakis.Attribute("area").Value,                  //19*ISTD AREA
                                    XEpeakis.Attribute("height").Value,                //20*ISTD HEIGHT

                                    XEpeakis.Attribute("foundrt").Value,               //21*ISTD RT
                                    XEpeak.Attribute("response").Value,                //22*RESPONSE

                                    XECurve.Attribute("type").Value,                        //23*CURVE
                                    XECurve.Attribute("weighting").Value.ToString(),        //24*WEIGHTING
                                    XECurve.Attribute("origin").Value.ToString(),           //25*ORIGIN

                                    fEQUATION,                //26*EQUATION
                                    fR,                       //27*R
                                    fR_SQR,                   //28*R_SQR
                                    fNUM_SLOPE,               //29*NUM_SLOPE
                                    fNUM_INTERSECT,           //30*NUM_INTERSECT

                                    AnalConc_EmptyCheck(
                                    _compoundQuanType == 1 ?
                                             XEpeak.Attribute("analconc").Value
                                             :
                                             XEpeak.Attribute("analconc").Value),      //31*NUM_ANAL_CONC

                                    _sample_type_id == 5 ?
                                             XEpeak.Attribute("concdevperc").Value
                                             :
                                             "NULL",                                   //32*CONC_DEV_PERC

                                    (XEpeak.Attribute("signoise").Value == "") ? "NULL" : XEpeak.Attribute("signoise").Value,    //33*SIGNOISE
                                    XEpeak.Attribute("chromnoise").Value,       //34*CHROMNOISE
                                    1 /*XEsample.Parent.Parent.Attribute("name").Value*/)
                                    );
                        }
                    }


                    // ProgressBar Begin
                    toolStripProgressBarMain.PerformStep();
                    //MessageBox.Show(toolStripProgressBarMain.Value.ToString());
                    toolStripStatusLabelMain.Text = "Proccesing ... " +
                        XEcompound.Attribute("name").Value + " - " + XEsample.Attribute("sampleid").Value;
                    toolStripStatusLabelMain.PerformClick();
                    // ProgressBar End

                    index++;
                }

                Excel.Range _xlRange_Std_Conc, _xlRange_Std_Response, _xlRange_Std_Dev, _xlRange_Std_AnalConc,
                            _xlRange_Analyte_SampleID, _xlRange_Analyte_AnalConc, _xlRange_Analyte_Area, _xlRange_Analyte_ISArea,
                            _xlRange_Std_Box, _xlRange_Analyte_Box, _xlRange_Sample_List_Box,
                            _xlRange_Tableau;

                //Standard Ranges****

                int _standardCount = _compoundQuanType == 1 ? _standardCount_R : _standardCount_A;
                int _analyteCount = _compoundQuanType == 1 ? _analyteCount_R : _analyteCount_A;
                int _qcCount = _compoundQuanType == 1 ? _qcCount_R : _qcCount_A;
                

                _xlRange_Std_Conc = _wsCurrent.get_Range
                (_OutputFormat._colLet_stdconc + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_stdconc + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Conc.Name = _wsCurrent.Name + "_Std_Conc";

                _xlRange_Std_Response = _wsCurrent.get_Range
                (_OutputFormat._colLet_response + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_response + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Response.Name = _wsCurrent.Name + "_Std_Response";

                _xlRange_Std_AnalConc = _wsCurrent.get_Range
                   (_OutputFormat._colLet_analconc + (_OutputFormat._titleShift + 1).ToString(),
                    _OutputFormat._colLet_analconc + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_AnalConc.Name = _wsCurrent.Name + "_Std_AnalConc";

                _xlRange_Std_Dev = _wsCurrent.get_Range
                (_OutputFormat._colLet_stddev + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_stddev + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Dev.Name = _wsCurrent.Name + "_Std_Deviation";

                _xlRange_Std_Box = _wsCurrent.get_Range
                (_OutputFormat._colLet_First + (_OutputFormat._titleShift + 1).ToString(),
                 _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount).ToString());
                _xlRange_Std_Box.Name = _wsCurrent.Name + "_Std_Box";


                //Analyte Ranges****
                _xlRange_Analyte_SampleID = _wsCurrent.get_Range
             (_OutputFormat._colLet_sampleid + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
              _OutputFormat._colLet_sampleid + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_SampleID.Name = _wsCurrent.Name + "_Analyte_SampleID";

                _xlRange_Analyte_AnalConc = _wsCurrent.get_Range
             (_OutputFormat._colLet_analconc + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
              _OutputFormat._colLet_analconc + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_AnalConc.Name = _wsCurrent.Name + "_Analyte_AnalConc";

                _xlRange_Analyte_Area = _wsCurrent.get_Range
              (_OutputFormat._colLet_analarea + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
               _OutputFormat._colLet_analarea + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_Area.Name = _wsCurrent.Name + "_Analyte_Area";

                _xlRange_Analyte_ISArea = _wsCurrent.get_Range
              (_OutputFormat._colLet_analISarea + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
               _OutputFormat._colLet_analISarea + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_ISArea.Name = _wsCurrent.Name + "_Analyte_ISArea";

                _xlRange_Analyte_Box = _wsCurrent.get_Range
             (_OutputFormat._colLet_First + (_OutputFormat._titleShift + _standardCount + 1).ToString(),
              _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Analyte_Box.Name = _wsCurrent.Name + "_Analyte_Box";


                //SAMPLE_LIST Range
                _xlRange_Sample_List_Box = _wsCurrent.get_Range
              (_OutputFormat._colLet_First + (_OutputFormat._titleShift + 1).ToString(),
            _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Sample_List_Box.Name = _wsCurrent.Name + "_Sample_Box";

                //Tableau Range
                _xlRange_Tableau = _wsCurrent.get_Range
              (_OutputFormat._colLet_First + (_OutputFormat._titleShift).ToString(),
            _OutputFormat._colLet_Last + (_OutputFormat._titleShift + _standardCount + _analyteCount).ToString());
                _xlRange_Tableau.Name = _wsCurrent.Name + "_Tableau";



                //Format Ranges**** 
                _xlRange_Sample_List_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                _xlRange_Sample_List_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).Weight = 3.75;
                _xlRange_Sample_List_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).Weight = 3.75;
                _xlRange_Std_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                _xlRange_Std_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).Weight = 1.75;
                _xlRange_Analyte_Box.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).Weight = 1.75;

                _xlRange_Std_Box.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);

                //////////////////////////////////////////////////////
                Excel.ColorScale cfColorScale = (Excel.ColorScale)(_xlRange_Analyte_ISArea.FormatConditions.AddColorScale(3));
                // Set the minimum threshold to red (0x000000FF) and maximum threshold
                // to blue (0x00FF0000).
                cfColorScale.ColorScaleCriteria[1].FormatColor.Color = 7039480;  //0x000000FF;  
                cfColorScale.ColorScaleCriteria[2].FormatColor.Color = 8711167;  // 0x00FF0000;
                cfColorScale.ColorScaleCriteria[3].FormatColor.Color = 13011546; // 0x00FF0000;
                /////////////////////////////////////////////////////


                //Plots****
                //Time Series****
                double _rowHeight = (double)_wsCurrent.get_Range("A1", "A1").RowHeight;

                
                double _graphOffset;

                _graphOffset = _compoundQuanType == 1 ?
                               (double)_xlRange_Sample_List_Box.Top + (double)_xlRange_Sample_List_Box.Height +
                               (_rowHeight + _rowHeight * _qcCount_R) 
                               :
                               (double)_xlRange_Sample_List_Box.Top + (double)_xlRange_Sample_List_Box.Height +
                               (_rowHeight + _rowHeight * _qcCount_A);

                Excel.ChartObjects _xlCharts = (Excel.ChartObjects)_wsCurrent.ChartObjects(Type.Missing);

                Excel.ChartObject _xlChartTimeCorse;

                string _serieCurveName = "default curve name";

                string _axeCurveXNumFormat = "0", _axeCurveYNumFormat = "0.0%";

                string _axeCurveXCaption = "default XCaption", _axeCurveYCaption = "default YCaption";

                
                _xlChartTimeCorse = (Excel.ChartObject)_xlCharts.Add(
                              0, _graphOffset, (double)_xlRange_Tableau.Width, 250);

                Excel.Chart _xlChartTimeCorsePage = _xlChartTimeCorse.Chart;
                _xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlXYScatter;
                _xlChartTimeCorsePage.HasTitle = true; _xlChartTimeCorsePage.ChartTitle.Font.Size = 11;

                _xlChartTimeCorsePage.ChartTitle.Caption = XEcompound.Attribute("name").Value + " " + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_TP = (Excel.Axes)_xlChartTimeCorsePage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_TP.HasMajorGridlines = true;
                _xlAxisX_TP.HasTitle = true; _xlAxisY_TP.HasTitle = true;

                _xlAxisX_TP.MinimumScale = 0; _xlAxisX_TP.MaximumScale = _maxTP;
                _xlAxisY_TP.MinimumScale = 0;

                _xlChartTimeCorsePage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartTimeCorsePage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_TimeCourse = (Excel.SeriesCollection)_xlChartTimeCorsePage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_TTR = (Excel.Series)_xlSeriesColl_TimeCourse.NewSeries();
                _xlSeries_TTR.XValues = _xlRange_Analyte_SampleID.Cells;
                _xlSeries_TTR.Values = _xlRange_Analyte_AnalConc;

                _xlSeries_TTR.ChartType = Excel.XlChartType.xlXYScatterSmooth;

                if (_compoundQuanType == 1)
                {
                    _serieCurveName = "H:L Abeta";

                    _axeCurveXNumFormat = "0"; _axeCurveYNumFormat = "0.0%";

                    _axeCurveXCaption = "h, Time"; _axeCurveYCaption = "TTR";

                    _xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlXYScatter;

                    _xlAxisX_TP.TickLabels.NumberFormat = _axeCurveXNumFormat;

                    _xlAxisX_TP.AxisTitle.Caption = _axeCurveXCaption;

                }
                else
                {
                    _axeCurveXCaption = "h, Time";
                    _axeCurveYNumFormat = "0.0";
                    _axeCurveYCaption = "Conc., ng/sample volume";
                    _serieCurveName = "Abeta level";

                    _xlAxisX_TP.AxisTitle.Caption = _axeCurveXCaption;

                    _xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlColumnStacked;

                }


                _xlAxisY_TP.TickLabels.NumberFormat = _axeCurveYNumFormat;
                _xlAxisY_TP.AxisTitle.Caption = _axeCurveYCaption;

                _xlSeries_TTR.Name = _serieCurveName;
                _xlSeries_TTR.Name = XEcompound.Attribute("name").Value;

                if (_standardCount_R == 0)
                {
                    _axeCurveYCaption = "Area Ratio";

                }

                
                _graphOffset += _xlChartTimeCorse.Height;

                Excel.ChartObject _xlChart_Area = (Excel.ChartObject)_xlCharts.Add(
                                  0, _graphOffset, (double)_xlRange_Tableau.Width, 250);
                Excel.Chart _xlChart_AreaPage = _xlChart_Area.Chart;

                _xlChart_AreaPage.ChartType = Excel.XlChartType.xlColumnClustered;
                _xlChart_AreaPage.HasTitle = true;
                _xlChart_AreaPage.ChartTitle.Font.Size = 11;

                _xlChart_AreaPage.ChartTitle.Caption = XEcompound.Attribute("name").Value + " " + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_Area = (Excel.Axes)_xlChart_AreaPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_Area = _xlAxes_Area.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_Area = _xlAxes_Area.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_Area.HasMajorGridlines = true;
                _xlAxisX_Area.HasTitle = true;
                _xlAxisY_Area.HasTitle = true;

                _xlAxisY_Area.MinimumScale = 0;

                _xlChart_AreaPage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChart_AreaPage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_Area = (Excel.SeriesCollection)_xlChart_AreaPage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_Area, _xlSeries_ISTD_Area;


                _xlSeries_ISTD_Area = (Excel.Series)_xlSeriesColl_Area.NewSeries();
                _xlSeries_ISTD_Area.XValues = _xlRange_Analyte_SampleID.Cells;
                _xlSeries_ISTD_Area.Values = _xlRange_Analyte_ISArea.Cells;
                _xlSeries_ISTD_Area.Name = "ISTD Area";
                                

                if (_compoundQuanType != 1)
                {
                    _xlSeries_Area = (Excel.Series)_xlSeriesColl_Area.NewSeries();
                    _xlSeries_Area.XValues = _xlRange_Analyte_SampleID.Cells;
                    _xlSeries_Area.Values = _xlRange_Analyte_Area.Cells;
                    _xlSeries_Area.Name = "Area";
                }
                
                _axeCurveYNumFormat = "0.0E+0"; _axeCurveYCaption = "Area";

                _xlAxisY_Area.TickLabels.NumberFormat = _axeCurveYNumFormat;
                _xlAxisX_Area.AxisTitle.Caption = _axeCurveXCaption;
                _xlAxisY_Area.AxisTitle.Caption = _axeCurveYCaption;

              
                

// RT graph           
                _graphOffset += _xlChart_Area.Height;

                Excel.ChartObject _xlChartRT = (Excel.ChartObject)_xlCharts.Add(
                                  0, _graphOffset, (double)_xlRange_Tableau.Width, 250);
                Excel.Chart _xlChartRTPage = _xlChartRT.Chart;

                _xlChartRTPage.ChartType = Excel.XlChartType.xlLineMarkers;
                _xlChartRTPage.HasTitle = true;
                _xlChartRTPage.ChartTitle.Font.Size = 11;

                _xlChartRTPage.ChartTitle.Caption = XEcompound.Attribute("name").Value + " " + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_RT = (Excel.Axes)_xlChartRTPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_RT = _xlAxes_RT.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_RT = _xlAxes_RT.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_RT.HasMajorGridlines = true;
                _xlAxisX_RT.HasTitle = true; _xlAxisY_RT.HasTitle = true;

                _xlChartRTPage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartRTPage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_RT = (Excel.SeriesCollection)_xlChartRTPage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_RT, _xlSeries_ISTD_RT;

                _xlSeries_ISTD_RT = (Excel.Series)_xlSeriesColl_RT.NewSeries();
                _xlSeries_ISTD_RT.XValues = _xlRange_Analyte_SampleID.Cells;
                _xlSeries_ISTD_RT.Values = _xlRange_Analyte_AnalConc.get_Offset(0, -2).Cells; ;

                _xlSeries_RT = (Excel.Series)_xlSeriesColl_RT.NewSeries();
                _xlSeries_RT.XValues = _xlRange_Analyte_SampleID.Cells;
                _xlSeries_RT.Values = _xlRange_Analyte_AnalConc.get_Offset(0, -3).Cells; ;

                

                _axeCurveYNumFormat = "0.00"; _axeCurveYCaption = "RT";
                
                _xlAxisY_RT.TickLabels.NumberFormat = _axeCurveYNumFormat;
                _xlAxisX_RT.AxisTitle.Caption = _axeCurveXCaption;
                _xlAxisY_RT.AxisTitle.Caption = _axeCurveYCaption;

                _xlSeries_RT.Name = "RT"; _xlSeries_ISTD_RT.Name = "ISTD RT";


//Std Curve Graphs
                Excel.ChartObject _xlChartStdCurve;

                _graphOffset += +(double)_xlChartRT.Height;

             
                    _xlChartStdCurve = (Excel.ChartObject)_xlCharts.Add(
                                       (double)_xlRange_Tableau.Width - 470, _graphOffset, 470, 250);
                    Excel.Chart _xlChartStdCurvePage = _xlChartStdCurve.Chart;
                    _xlChartStdCurvePage.ChartType = Excel.XlChartType.xlXYScatter;

                    _xlChartStdCurvePage.HasTitle = true;
                    _xlChartStdCurvePage.ChartTitle.Font.Size = 11;

                    string _chartTitleStdCurve = String.Format
                        ("Calibration curve {0} {1} {2} {3} ({4})",
                        XEcompound.Attribute("name").Value,
                        ExperimentInfo._IP,
                        ExperimentInfo._Enzyme,
                        ExperimentInfo._Instument,
                        ExperimentInfo._Date);

                    _xlChartStdCurvePage.ChartTitle.Caption = _chartTitleStdCurve;

                    Excel.Axes _xlAxes_Std = (Excel.Axes)_xlChartStdCurvePage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisX_Std = _xlAxes_Std.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisY_Std = _xlAxes_Std.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                    _xlAxisX_Std.HasTitle = true; _xlAxisY_Std.HasTitle = true;
                    _xlAxisX_Std.HasMajorGridlines = true;
                    _xlAxisY_Std.HasMajorGridlines = true;
                    _xlAxisY_Std.MinimumScale = 0;

                    _xlChartStdCurvePage.Legend.IncludeInLayout = true;
                    _xlChartStdCurvePage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionTop;

                    Excel.SeriesCollection _xlSerColl_StdCurve = (Excel.SeriesCollection)_xlChartStdCurvePage.SeriesCollection(Type.Missing);

                    Excel.Series _xlSerie_Stds = (Excel.Series)_xlSerColl_StdCurve.NewSeries();
                    _xlSerie_Stds.XValues = _xlRange_Std_Conc.Cells;
                    _xlSerie_Stds.Values = _xlRange_Std_Response.Cells;

                    string _serieStdName = null;

                    string _axeStdXNumFormat = null, _axeStdYNumFormat = null;

                    string _axeStdXCaption = null, _axeStdYCaption = null;

                    if (_compoundQuanType == 1)
                    {
                        _serieStdName = "H:L Abeta";
                        _axeStdXNumFormat = "0.00"; _axeStdYNumFormat = "0.00";
                        _axeStdXCaption = "%, Predicted TTR"; _axeStdYCaption = "Response";
                    }
                    else
                    {
                        _serieStdName = "Analyte:ISTD Abeta";
                        _axeStdXNumFormat = "0.0"; _axeStdYNumFormat = "0.0";
                        _axeStdXCaption = "ng, Specified amount"; _axeStdYCaption = "Response";
                    }


                    _xlAxisX_Std.TickLabels.NumberFormat = _axeStdXNumFormat;
                    _xlAxisY_Std.TickLabels.NumberFormat = _axeStdYNumFormat;
                    _xlAxisX_Std.AxisTitle.Caption = _axeStdXCaption;
                    _xlAxisY_Std.AxisTitle.Caption = _axeStdYCaption;

                    _xlSerie_Stds.Name = _serieStdName;

                    Excel.Trendlines _xlTrendlines_Stds = (Excel.Trendlines)_xlSerie_Stds.Trendlines(Type.Missing);

                    Excel.XlTrendlineType _XlTrendlineType =
                        Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear;
                    string _CurveIndex = XECurve.Attribute("type").Value;
 
                    if (_CurveIndex == "Quadratic")
                    { _XlTrendlineType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlExponential; }


                    _xlTrendlines_Stds.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, true, true, Type.Missing);


               // }

//
                Excel.ChartObjects _xlChartsSummary = (Excel.ChartObjects)_wsSummary.ChartObjects(Type.Missing);
                Excel.ChartObject _xlChartCurveSummary, _xlChartLevelSummary;
                Excel.Chart _xlChartCurveSummaryPage, _xlChartLevelSummaryPage;

                if (_xlChartsSummary.Count < 1)
                {
                    _axeCurveYCaption = "TTR, %";
                    _axeCurveYNumFormat = "0.00%";

                    _xlChartCurveSummary = (Excel.ChartObject)_xlChartsSummary.Add(
                                   10, 10, 850, 500);
                    _xlChartCurveSummaryPage = _xlChartCurveSummary.Chart;
                    _xlChartCurveSummaryPage.ChartType = Excel.XlChartType.xlXYScatterSmooth;
                    _xlChartCurveSummaryPage.HasTitle = true;
                    _xlChartCurveSummaryPage.ChartTitle.Font.Size = 11;

                    Excel.Axes _xlAxes_CurveSummary = (Excel.Axes)_xlChartCurveSummaryPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisX_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisY_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                    _xlAxisX_CurveSummary.HasMajorGridlines = true; _xlAxisY_CurveSummary.HasMajorGridlines = true;

                    _xlAxisX_CurveSummary.HasTitle = true; _xlAxisY_CurveSummary.HasTitle = true;

                    _xlAxisX_CurveSummary.MaximumScale = _maxTP;

                    _xlAxisX_CurveSummary.AxisTitle.Caption = _axeCurveXCaption; _xlAxisY_CurveSummary.AxisTitle.Caption = _axeCurveYCaption;

                    _xlAxisX_CurveSummary.TickLabels.NumberFormat = _axeCurveXNumFormat; _xlAxisY_CurveSummary.TickLabels.NumberFormat = _axeCurveYNumFormat;

                    _xlChartCurveSummaryPage.ChartTitle.Caption = "TTR Abeta in ";

                    _xlChartCurveSummaryPage.ChartTitle.Caption += ExperimentInfo._CommonChartNameStr;

                    _xlChartCurveSummaryPage.Legend.Position = Excel.XlLegendPosition.xlLegendPositionRight;
                    _xlChartCurveSummaryPage.Legend.IncludeInLayout = true;

 
                    _axeCurveYCaption = "Conc., ng/sample volume";
                    _axeCurveYNumFormat = "0.0";

                    _xlChartLevelSummary = (Excel.ChartObject)_xlChartsSummary.Add(
                                   10, 510, 850, 500);
                    _xlChartLevelSummaryPage = _xlChartLevelSummary.Chart;
                    _xlChartLevelSummaryPage.ChartType = Excel.XlChartType.xlXYScatter;
                    _xlChartLevelSummaryPage.HasTitle = true;
                    _xlChartLevelSummaryPage.ChartTitle.Font.Size = 11;

                    Excel.Axes _xlAxes_LevelSummary = (Excel.Axes)_xlChartLevelSummaryPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisX_LevelSummary = _xlAxes_LevelSummary.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisY_LevelSummary = _xlAxes_LevelSummary.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                    _xlAxisX_LevelSummary.HasMajorGridlines = true; _xlAxisY_LevelSummary.HasMajorGridlines = true;

                    _xlAxisX_LevelSummary.HasTitle = true; _xlAxisY_LevelSummary.HasTitle = true;

                    _xlAxisX_LevelSummary.MaximumScale = _maxTP;

                    _xlAxisX_LevelSummary.AxisTitle.Caption = _axeCurveXCaption; _xlAxisY_LevelSummary.AxisTitle.Caption = _axeCurveYCaption;

                    _xlAxisX_LevelSummary.TickLabels.NumberFormat = _axeCurveXNumFormat; _xlAxisY_LevelSummary.TickLabels.NumberFormat = _axeCurveYNumFormat;

                    _xlChartLevelSummaryPage.ChartTitle.Caption = "Abeta levels in ";

                    _xlChartLevelSummaryPage.ChartTitle.Caption += ExperimentInfo._CommonChartNameStr;

                    _xlChartLevelSummaryPage.Legend.Position = Excel.XlLegendPosition.xlLegendPositionRight;
                    _xlChartLevelSummaryPage.Legend.IncludeInLayout = true;

                }
                else
                {
                    _xlChartCurveSummary = (Excel.ChartObject)_xlChartsSummary.Item(1);
                    _xlChartCurveSummaryPage = _xlChartCurveSummary.Chart;
                    _xlChartLevelSummary = (Excel.ChartObject)_xlChartsSummary.Item(2);
                    _xlChartLevelSummaryPage = _xlChartLevelSummary.Chart;
                }


                Excel.SeriesCollection _xlSeriesColl_Summary =
                    _compoundQuanType == 1 ?
                                            (Excel.SeriesCollection)_xlChartCurveSummaryPage.SeriesCollection(Type.Missing)
                                            :
                                            (Excel.SeriesCollection)_xlChartLevelSummaryPage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_timeCourse = (Excel.Series)_xlSeriesColl_Summary.NewSeries();
                _xlSeries_timeCourse.XValues = _xlRange_Analyte_SampleID.Cells;
                _xlSeries_timeCourse.Values = _xlRange_Analyte_AnalConc;
                _xlSeries_timeCourse.Name = XEcompound.Attribute("name").Value;

            }

            toolStripProgressBarMain.Visible = false; toolStripProgressBarMain.Value = 0;

            toolStripStatusLabelMain.Text = "Done"; toolStripStatusLabelMain.PerformClick();


            string _AppPath = Application.StartupPath + Path.DirectorySeparatorChar;
            toolStripStatusLabelMain.Text = "Exporting data into DB ...";
            toolStripStatusLabelMain.PerformClick();
            string[] _sql_Part1 = File.ReadAllLines(_AppPath + "sql_XevoImport_P1.sql");
            string[] _sql_Part2 = File.ReadAllLines(_AppPath + "sql_XevoImport_P2.sql");
            string[] _sql_Part3 = File.ReadAllLines(_AppPath + "sql_XevoImport_P3.sql");

            _strExportSQL.Insert(0, " ");
            _strExportSQL.InsertRange(0, _sql_Part1.ToList<string>());

            _strExportSQL.Insert(0, "/*");

            _strExportSQL.Insert(1, String.Format("Generated by QLD Reader {0} from {1}", Assembly.GetExecutingAssembly().GetName().Version.ToString(), Convert.ToDateTime(Properties.Resources.BuildDate).ToShortDateString()));
            _strExportSQL.Insert(2, String.Format("       Machine: {0} ran by {1}@{2}", Environment.MachineName, Environment.UserName, Environment.UserDomainName));
            _strExportSQL.Insert(3, String.Format("     Timestamp: {0} @ {1}", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString()));
            _strExportSQL.Insert(4, String.Format("        Source: {0} ", textBox_QLDfile.Text));

            _strExportSQL.Insert(5, "*/");



            _strExportSQL.AddRange(_sql_Part3.ToList<string>());

            File.WriteAllLines(Path.ChangeExtension(textBox_Excelfile.Text, "sql"), _strExportSQL.ToArray());

            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();



            object misValue = System.Reflection.Missing.Value;

            try
            {
                xlWB.SaveAs(textBox_Excelfile.Text,
                   misValue,
                   misValue, misValue, misValue, misValue,
                   Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlUserResolution,
                   misValue, misValue, misValue, misValue);
            }

            catch (COMException _exception)
            {
                if (_exception.ErrorCode == unchecked((int)0x800A03EC))
                {
                    MessageBox.Show("Time stamp has been added to file name you specified", this.Text,
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    string _newExcelFileName = textBox_Excelfile.Text.Replace(
                        Path.GetFileNameWithoutExtension(textBox_Excelfile.Text),
                        Path.GetFileNameWithoutExtension(textBox_Excelfile.Text)
                        + "_created_"
                        + System.DateTime.Now.ToString("s")
                        .Replace("/", "")
                        .Replace(":", "")
                        .Replace(" ", "_")
                        );

                    xlWB.SaveAs(_newExcelFileName,
                       misValue,
                       misValue, misValue, misValue, misValue,
                       Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlUserResolution,
                       misValue, misValue, misValue, misValue);
                }
            }

            finally
            {
                if (this.checkBox_OpenExcel.Checked)
                {
                    _xlApp.Visible = true;

                    _xlApp.ScreenUpdating = true;
                    // _wsSummary.Activate();

                }
                else
                {
                    _xlApp.Quit();
                }
            }

            if (checkBox_ShowDebug.Checked)
            {
                MessageBox.Show("ExportIVBolus_csf()");
            }

        }

        private struct ElementCounter
        {
            public int quandataset, xmlfile, dataset, groupdata, group, sample, compound, peak, ispeak, method;
        }

        private void XMLToCSV()
        {
            List<string> _strCSV_Row = new List<string>();

            XDocument _xdQUANDATASET;
            try
            {
                _xdQUANDATASET = XDocument.Load(textBox_QLDfile.Text);
            }
            catch (Exception _XMLExeption)
            {
                MessageBox.Show(_XMLExeption.Message);
                return;
            }


            ElementCounter _elemCnt = new ElementCounter();

            string _strCSVSeparator = "|", _strSpace = " ";
            string _strCSVHeader = "";

            _elemCnt.quandataset++;

            IEnumerable<XAttribute> _QUANDATASET_attrList = from at in _xdQUANDATASET.Root.Attributes()
                                                            select at;
            string _strCSV_QUANDATASETData = "";

            foreach (XAttribute attr in _QUANDATASET_attrList)
            {
                if (_elemCnt.quandataset == 1)
                {
                    _strCSVHeader = _strCSVHeader + "QUANDATASET_" + attr.Name + _strCSVSeparator + _strSpace;
                }
                _strCSV_QUANDATASETData = _strCSV_QUANDATASETData + attr.Value + _strCSVSeparator + _strSpace;
            }



            XElement _xeXMLFILE = _xdQUANDATASET.Root.Element("XMLFILE");

            _elemCnt.xmlfile++;

            IEnumerable<XAttribute> _XMLFILE_attrList = from at in _xeXMLFILE.Attributes()
                                                        select at;
            string _strCSV_XMLFILEData = "";

            foreach (XAttribute attr in _XMLFILE_attrList)
            {
                if (_elemCnt.xmlfile == 1)
                {
                    _strCSVHeader = _strCSVHeader + "XMLFILE_" + attr.Name + _strCSVSeparator + _strSpace;
                }
                _strCSV_XMLFILEData = _strCSV_XMLFILEData + attr.Value + _strCSVSeparator + _strSpace;
            }


            XElement _xeDATASET = _xdQUANDATASET.Root.Element("DATASET");

            _elemCnt.dataset++;

            IEnumerable<XAttribute> _DATASET_attrList = from at in _xeDATASET.Attributes()
                                                        select at;
            string _strCSV_DATASETData = "";

            foreach (XAttribute attr in _DATASET_attrList)
            {
                if (_elemCnt.dataset == 1)
                {
                    _strCSVHeader = _strCSVHeader + "DATASET_" + attr.Name + _strCSVSeparator + _strSpace;
                }
                _strCSV_DATASETData = _strCSV_DATASETData + attr.Value + _strCSVSeparator + _strSpace;
            }

            XElement _xeGROUPDATA = _xdQUANDATASET.Root.Element("GROUPDATA");

            _elemCnt.groupdata++;

            IEnumerable<XAttribute> _GROUPDATA_attrList = from at in _xeGROUPDATA.Attributes()
                                                          select at;
            string _strCSV_GROUPDATAData = "";

            foreach (XAttribute attr in _GROUPDATA_attrList)
            {
                if (_elemCnt.groupdata == 1)
                {
                    _strCSVHeader = _strCSVHeader + "GROUPDATA_" + attr.Name + _strCSVSeparator + _strSpace;
                }
                _strCSV_GROUPDATAData = _strCSV_GROUPDATAData + attr.Value + _strCSVSeparator + _strSpace;
            }


            var queryGROUP_LIST = from _group in _xeGROUPDATA.Descendants("GROUP")
                                  select _group;


            foreach (XElement _xeGROUP in queryGROUP_LIST)
            {
                _elemCnt.group++;

                IEnumerable<XAttribute> _GROUP_attrList = from at in _xeGROUP.Attributes()
                                                          select at;
                string _strCSV_GROUPData = "";

                foreach (XAttribute attr in _GROUP_attrList)
                {
                    if (_elemCnt.group == 1)
                    {
                        _strCSVHeader = _strCSVHeader + "GROUP_" + attr.Name + _strCSVSeparator + _strSpace;
                    }
                    _strCSV_GROUPData = _strCSV_GROUPData + attr.Value + _strCSVSeparator + _strSpace;
                }


                XElement _xeSAMPLELISTDATA = _xeGROUP.Element("SAMPLELISTDATA");

                var querySAMPLE_LIST = from _sample in _xeSAMPLELISTDATA.Descendants("SAMPLE")
                                       select _sample;
                foreach (XElement _xeSAMPLE in querySAMPLE_LIST)
                {
                    _elemCnt.sample++;

                    IEnumerable<XAttribute> _SAMPLE_attrList = from at in _xeSAMPLE.Attributes()
                                                               select at;
                    string _strCSV_SAMPLEData = "";

                    foreach (XAttribute attr in _SAMPLE_attrList)
                    {
                        if (_elemCnt.sample == 1)
                        {
                            _strCSVHeader = _strCSVHeader + "SAMPLE_" + attr.Name + _strCSVSeparator + _strSpace;
                        }
                        _strCSV_SAMPLEData = _strCSV_SAMPLEData + attr.Value + _strCSVSeparator + _strSpace;
                    }

                    var queryCOMPOUND_LIST = from _compound in _xeSAMPLE.Descendants("COMPOUND")
                                             select _compound;
                    foreach (XElement _xeCOMPOUND in queryCOMPOUND_LIST)
                    {
                        _elemCnt.compound++;

                        IEnumerable<XAttribute> _COMPOUND_attrList = from at in _xeCOMPOUND.Attributes()
                                                                     select at;
                        string _strCSV_COMPOUNDData = "";

                        foreach (XAttribute attr in _COMPOUND_attrList)
                        {
                            if (_elemCnt.compound == 1)
                            {
                                _strCSVHeader = _strCSVHeader + "COMPOUND_" + attr.Name + _strCSVSeparator + _strSpace;
                            }
                            _strCSV_COMPOUNDData = _strCSV_COMPOUNDData + attr.Value + _strCSVSeparator + _strSpace;
                        }

                        XElement _xePEAK = _xeCOMPOUND.Element("PEAK");

                        _elemCnt.peak++;

                        IEnumerable<XAttribute> _PEAK_attrList = from at in _xePEAK.Attributes()
                                                                 select at;
                        string _strCSV_PEAKData = "";

                        foreach (XAttribute attr in _PEAK_attrList)
                        {
                            if (_elemCnt.peak == 1)
                            {
                                _strCSVHeader = _strCSVHeader + "PEAK_" + attr.Name + _strCSVSeparator + _strSpace;
                            }
                            _strCSV_PEAKData = _strCSV_PEAKData + attr.Value + _strCSVSeparator + _strSpace;
                        }

                        XElement _xeISPEAK = _xePEAK.Element("ISPEAK");

                        _elemCnt.ispeak++;

                        IEnumerable<XAttribute> _ISPEAK_attrList = from at in _xeISPEAK.Attributes()
                                                                   select at;
                        string _strCSV_ISPEAKData = "";

                        foreach (XAttribute attr in _ISPEAK_attrList)
                        {
                            if (_elemCnt.ispeak == 1)
                            {
                                _strCSVHeader = _strCSVHeader + "ISPEAK_" + attr.Name + _strCSVSeparator + _strSpace;
                            }
                            _strCSV_ISPEAKData = _strCSV_ISPEAKData + attr.Value + _strCSVSeparator + _strSpace;
                        }


                        XElement _xeMETHOD = _xeCOMPOUND.Element("METHOD");

                        _elemCnt.method++;

                        IEnumerable<XAttribute> _METHOD_attrList = from at in _xeMETHOD.Attributes()
                                                                   select at;
                        string _strCSV_METHODData = "";

                        foreach (XAttribute attr in _METHOD_attrList)
                        {
                            if (_elemCnt.method == 1)
                            {
                                _strCSVHeader = _strCSVHeader + "METHOD_" + attr.Name + _strCSVSeparator + _strSpace;
                            }
                            _strCSV_METHODData = _strCSV_METHODData + attr.Value + _strCSVSeparator + _strSpace;
                        }



                        var _queryCAL_COMPOUND = from _compound in _xeGROUP.Element("CALIBRATIONDATA").Elements("COMPOUND")
                                                 where _compound.Attribute("id").Value == _xeCOMPOUND.Attribute("id").Value
                                                 select _compound;

                        XElement _xeCAL_COMPOUND = _queryCAL_COMPOUND.Single();

                        string _strCSV_CAL_COMPOUND = "";

                        if (_elemCnt.compound == 1)
                        {
                            _strCSVHeader += String.Format("{1}{0} {2}{0} {3}{0} {4}{0} {5}{0} {6}{0} {7}{0} {8}{0} {9}{0} {10}{0}", _strCSVSeparator,
                                                            "RESPONSE_type",
                                                            "RESPONSE_ref",
                                                            "RESPONSE_rah",
                                                            "CURVE_type",
                                                            "CURVE_origin",
                                                            "CURVE_weighting",
                                                            "CURVE_axistrans",
                                                            "CALIBRATIONCURVE_curve",
                                                            "CORRELATION_r",
                                                            "DETERMINATION_rsquared");
                        }
                        _strCSV_CAL_COMPOUND = String.Format("{1}{0} {2}{0} {3}{0} {4}{0} {5}{0} {6}{0} {7}{0} {8}{0} {9}{0} {10}{0}", _strCSVSeparator,
                                                              _xeCAL_COMPOUND.Element("RESPONSE").Attribute("type").Value,
                                                              _xeCAL_COMPOUND.Element("RESPONSE").Attribute("ref").Value,
                                                              _xeCAL_COMPOUND.Element("RESPONSE").Attribute("rah").Value,
                                                              _xeCAL_COMPOUND.Element("CURVE").Attribute("type").Value,
                                                              _xeCAL_COMPOUND.Element("CURVE").Attribute("origin").Value,
                                                              _xeCAL_COMPOUND.Element("CURVE").Attribute("weighting").Value,
                                                              _xeCAL_COMPOUND.Element("CURVE").Attribute("axistrans").Value,

                                                              (_xeCAL_COMPOUND.Element("CURVE").Element("CALIBRATIONCURVE") == null) ? ""
                                                              : _xeCAL_COMPOUND.Element("CURVE").Element("CALIBRATIONCURVE").Attribute("curve").Value,

                                                              (_xeCAL_COMPOUND.Element("CURVE").Element("CORRELATION") == null) ? ""
                                                              : _xeCAL_COMPOUND.Element("CURVE").Element("CORRELATION").Attribute("r").Value,

                                                               (_xeCAL_COMPOUND.Element("CURVE").Element("DETERMINATION") == null) ? ""
                                                              : _xeCAL_COMPOUND.Element("CURVE").Element("DETERMINATION").Attribute("rsquared").Value

                                                            );

                        /*************************/
                        /**/
                        string _strCSV_DATARow = _strCSV_QUANDATASETData + _strCSV_XMLFILEData + _strCSV_DATASETData + _strCSV_GROUPDATAData + _strCSV_GROUPData + _strCSV_SAMPLEData + _strCSV_COMPOUNDData
                                                 +
                                                 _strCSV_PEAKData + _strCSV_ISPEAKData + _strCSV_METHODData
                                                 + _strCSV_CAL_COMPOUND;
                        /**/
                        _strCSV_Row.Add(_strCSV_DATARow);
                        /*************************/
                    }
                }
            }

            _strCSV_Row.Insert(0, _strCSVHeader);



            File.WriteAllLines(Path.ChangeExtension(textBox_Excelfile.Text, "csv"), _strCSV_Row.ToArray());


            if (checkBox_ShowDebug.Checked)
            {
                MessageBox.Show("XMLToCSV()");
            }
        }



        private void XMLToCSV1()
        {
            List<string> _strCSV_Row = new List<string>();

            XDocument _xdQUANDATASET;
            try
            {
                _xdQUANDATASET = XDocument.Load(textBox_QLDfile.Text);
            }
            catch (Exception _XMLExeption)
            {
                MessageBox.Show(_XMLExeption.Message);
                return;
            }


            ElementCounter _elemCnt = new ElementCounter();

            string _strCSVSeparator = "|", _strSpace = " "; 
            string _strCSVHeader = "";

            _elemCnt.quandataset++;

            IEnumerable<XAttribute> _QUANDATASET_attrList = from at in _xdQUANDATASET.Root.Attributes()
                                                            select at;
            string _strCSV_QUANDATASETData = "";

            foreach (XAttribute attr in _QUANDATASET_attrList)
            {
                if (_elemCnt.quandataset == 1)
                {
                    _strCSVHeader = _strCSVHeader + "QUANDATASET_" + attr.Name + _strCSVSeparator + _strSpace;
                }
                _strCSV_QUANDATASETData = _strCSV_QUANDATASETData + attr.Value + _strCSVSeparator + _strSpace;
            }

           

            XElement _xeXMLFILE = _xdQUANDATASET.Root.Element("XMLFILE");

            _elemCnt.xmlfile++;

            IEnumerable<XAttribute> _XMLFILE_attrList = from at in _xeXMLFILE.Attributes()
                                                        select at;
            string _strCSV_XMLFILEData = "";

            foreach (XAttribute attr in _XMLFILE_attrList)
            {
                if (_elemCnt.xmlfile == 1)
                {
                    _strCSVHeader = _strCSVHeader + "XMLFILE_" + attr.Name + _strCSVSeparator + _strSpace;
                }
                _strCSV_XMLFILEData = _strCSV_XMLFILEData + attr.Value + _strCSVSeparator + _strSpace;
            }


            XElement _xeDATASET = _xdQUANDATASET.Root.Element("DATASET");

            _elemCnt.dataset++;

            IEnumerable<XAttribute> _DATASET_attrList = from at in _xeDATASET.Attributes()
                                                        select at;
            string _strCSV_DATASETData = "";

            foreach (XAttribute attr in _DATASET_attrList)
            {
                if (_elemCnt.dataset == 1)
                {
                    _strCSVHeader = _strCSVHeader + "DATASET_" + attr.Name + _strCSVSeparator + _strSpace;
                }
                _strCSV_DATASETData = _strCSV_DATASETData + attr.Value + _strCSVSeparator + _strSpace;
            }

            XElement _xeGROUPDATA = _xdQUANDATASET.Root.Element("GROUPDATA");

            _elemCnt.groupdata++;

            IEnumerable<XAttribute> _GROUPDATA_attrList = from at in _xeGROUPDATA.Attributes()
                                                        select at;
            string _strCSV_GROUPDATAData = "";

            foreach (XAttribute attr in _GROUPDATA_attrList)
            {
                if (_elemCnt.groupdata == 1)
                {
                    _strCSVHeader = _strCSVHeader + "GROUPDATA_" + attr.Name + _strCSVSeparator + _strSpace;
                }
                _strCSV_GROUPDATAData = _strCSV_GROUPDATAData + attr.Value + _strCSVSeparator + _strSpace;
            }

           
            var queryGROUP_LIST = from _group in _xeGROUPDATA.Descendants("GROUP")
                                  select _group;


            foreach (XElement _xeGROUP in queryGROUP_LIST)
            {
                _elemCnt.group++;

                IEnumerable<XAttribute> _GROUP_attrList = from at in _xeGROUP.Attributes()
                                                          select at;
                string _strCSV_GROUPData = "";

                foreach (XAttribute attr in _GROUP_attrList)
                {
                    if (_elemCnt.group == 1)
                    {
                        _strCSVHeader = _strCSVHeader + "GROUP_" + attr.Name + _strCSVSeparator + _strSpace;
                    }
                    _strCSV_GROUPData = _strCSV_GROUPData + attr.Value + _strCSVSeparator + _strSpace;
                }

                
                XElement _xeSAMPLELISTDATA = _xeGROUP.Element("SAMPLELISTDATA");
                
                var querySAMPLE_LIST = from _sample in _xeSAMPLELISTDATA.Descendants("SAMPLE")
                                       select _sample;
                foreach (XElement _xeSAMPLE in querySAMPLE_LIST)
                {
                    _elemCnt.sample++;

                    IEnumerable<XAttribute> _SAMPLE_attrList = from at in _xeSAMPLE.Attributes()
                                                               select at;
                    string _strCSV_SAMPLEData = "";
                    
                    foreach (XAttribute attr in _SAMPLE_attrList)
                    {
                        if (_elemCnt.sample == 1)
                        {
                            _strCSVHeader = _strCSVHeader + "SAMPLE_" + attr.Name + _strCSVSeparator + _strSpace;
                        }
                        _strCSV_SAMPLEData = _strCSV_SAMPLEData + attr.Value + _strCSVSeparator + _strSpace;
                    }

                    var queryCOMPOUND_LIST = from _compound in _xeSAMPLE.Descendants("COMPOUND")
                                             select _compound;
                    foreach (XElement _xeCOMPOUND in queryCOMPOUND_LIST)
                    {
                        _elemCnt.compound++;

                        IEnumerable<XAttribute> _COMPOUND_attrList = from at in _xeCOMPOUND.Attributes()
                                                                   select at;
                        string _strCSV_COMPOUNDData = "";

                        foreach (XAttribute attr in _COMPOUND_attrList)
                        {
                            if (_elemCnt.compound == 1)
                            {
                                _strCSVHeader = _strCSVHeader + "COMPOUND_" + attr.Name + _strCSVSeparator + _strSpace;
                            }
                            _strCSV_COMPOUNDData = _strCSV_COMPOUNDData + attr.Value + _strCSVSeparator + _strSpace;
                        }

                        XElement _xePEAK = _xeCOMPOUND.Element("PEAK");

                        _elemCnt.peak++;

                        IEnumerable<XAttribute> _PEAK_attrList = from at in _xePEAK.Attributes()
                                                                 select at;
                        string _strCSV_PEAKData = "";

                        foreach (XAttribute attr in _PEAK_attrList)
                        {
                            if (_elemCnt.peak == 1)
                            {
                                _strCSVHeader = _strCSVHeader + "PEAK_" + attr.Name + _strCSVSeparator + _strSpace;
                            }
                            _strCSV_PEAKData = _strCSV_PEAKData + attr.Value + _strCSVSeparator + _strSpace;
                        }

                        XElement _xeISPEAK = _xePEAK.Element("ISPEAK");

                        _elemCnt.ispeak++;

                        IEnumerable<XAttribute> _ISPEAK_attrList = from at in _xeISPEAK.Attributes()
                                                                   select at;
                        string _strCSV_ISPEAKData = "";

                        foreach (XAttribute attr in _ISPEAK_attrList)
                        {
                            if (_elemCnt.ispeak == 1)
                            {
                                _strCSVHeader = _strCSVHeader + "ISPEAK_" + attr.Name + _strCSVSeparator + _strSpace;
                            }
                            _strCSV_ISPEAKData = _strCSV_ISPEAKData + attr.Value + _strCSVSeparator + _strSpace;
                        }

                        
                        XElement _xeMETHOD = _xeCOMPOUND.Element("METHOD");

                        _elemCnt.method++;

                        IEnumerable<XAttribute> _METHOD_attrList = from at in _xeMETHOD.Attributes()
                                                                   select at;
                        string _strCSV_METHODData = "";

                        foreach (XAttribute attr in _METHOD_attrList)
                        {
                            if (_elemCnt.method == 1)
                            {
                                _strCSVHeader = _strCSVHeader + "METHOD_" + attr.Name + _strCSVSeparator + _strSpace;
                            }
                            _strCSV_METHODData = _strCSV_METHODData + attr.Value + _strCSVSeparator + _strSpace;
                        }


                     
                        var _queryCAL_COMPOUND = from _compound in _xeGROUP.Element("CALIBRATIONDATA").Elements("COMPOUND")
                                              where _compound.Attribute("id").Value == _xeCOMPOUND.Attribute("id").Value
                                              select _compound;

                        XElement _xeCAL_COMPOUND = _queryCAL_COMPOUND.Single();

                        string _strCSV_CAL_COMPOUND = "";

                        if (_elemCnt.compound == 1)
                        {
                            _strCSVHeader += String.Format("{1}{0} {2}{0} {3}{0} {4}{0} {5}{0} {6}{0} {7}{0} {8}{0} {9}{0} {10}{0}", _strCSVSeparator,
                                                            "RESPONSE_type",
                                                            "RESPONSE_ref",
                                                            "RESPONSE_rah",
                                                            "CURVE_type",
                                                            "CURVE_origin",
                                                            "CURVE_weighting",
                                                            "CURVE_axistrans",
                                                            "CALIBRATIONCURVE_curve",
                                                            "CORRELATION_r",
                                                            "DETERMINATION_rsquared");
                                                                                                            }
                        _strCSV_CAL_COMPOUND = String.Format("{1}{0} {2}{0} {3}{0} {4}{0} {5}{0} {6}{0} {7}{0} {8}{0} {9}{0} {10}{0}", _strCSVSeparator,
                                                              _xeCAL_COMPOUND.Element("RESPONSE").Attribute("type").Value,
                                                              _xeCAL_COMPOUND.Element("RESPONSE").Attribute("ref").Value,
                                                              _xeCAL_COMPOUND.Element("RESPONSE").Attribute("rah").Value,
                                                              _xeCAL_COMPOUND.Element("CURVE").Attribute("type").Value,
                                                              _xeCAL_COMPOUND.Element("CURVE").Attribute("origin").Value,
                                                              _xeCAL_COMPOUND.Element("CURVE").Attribute("weighting").Value,
                                                              _xeCAL_COMPOUND.Element("CURVE").Attribute("axistrans").Value,

                                                              (_xeCAL_COMPOUND.Element("CURVE").Element("CALIBRATIONCURVE") == null) ? ""
                                                              : _xeCAL_COMPOUND.Element("CURVE").Element("CALIBRATIONCURVE").Attribute("curve").Value,

                                                              (_xeCAL_COMPOUND.Element("CURVE").Element("CORRELATION") == null) ? ""
                                                              : _xeCAL_COMPOUND.Element("CURVE").Element("CORRELATION").Attribute("r").Value,

                                                               (_xeCAL_COMPOUND.Element("CURVE").Element("DETERMINATION") == null) ? ""
                                                              : _xeCAL_COMPOUND.Element("CURVE").Element("DETERMINATION").Attribute("rsquared").Value

                                                            );

                      /*************************/
                      /**/  string _strCSV_DATARow = _strCSV_QUANDATASETData + _strCSV_XMLFILEData + _strCSV_DATASETData + _strCSV_GROUPDATAData + _strCSV_GROUPData + _strCSV_SAMPLEData + _strCSV_COMPOUNDData
                                                     +
                                                     _strCSV_PEAKData + _strCSV_ISPEAKData + _strCSV_METHODData
                                                     + _strCSV_CAL_COMPOUND;
                      /**/  _strCSV_Row.Add( _strCSV_DATARow );
                      /*************************/
                    }
                }
            }

            _strCSV_Row.Insert(0, _strCSVHeader);

           

            File.WriteAllLines(Path.ChangeExtension(textBox_Excelfile.Text, "csv"), _strCSV_Row.ToArray());

           
            if (checkBox_ShowDebug.Checked)
            {
                MessageBox.Show("XMLToCSV()");
            }
        }


        private void button_StartExport_Click(object sender, EventArgs e)
        {
            if (checkBox_Tableau.Checked)
            {
                //XMLToCSV();
               // XMLToCSV();
            }

            MessageBox.Show(comboBox_Project.SelectedValue.ToString());
/*
            switch (Convert.ToByte(comboBox_Project.SelectedValue))
            {

                case 4: ExportLOAD_plasma();
                    // MessageBox.Show("LOAD100 Plasma" + " Complete");
                    break;

                case 5: ExportLOAD_plasma();
                    // MessageBox.Show("LOAD100 Plasma" + " Complete");
                    break;

                case 7: ExportLOAD_plasma();
                    // MessageBox.Show("LOAD100 Plasma" + " Complete");
                    break;

                case 6: ExportFACS_csf();
                    // MessageBox.Show("ApoE" + " Complete");
                    break;

                case 8: ExportIVBolus_csf();
                    // MessageBox.Show("D-series " + " Complete");
                    break;

                case 11: ExportSleepStudy_csf();
                    // MessageBox.Show("D-series " + " Complete");
                    break;

                case 12: ExportLOAD_plasma();
                    // MessageBox.Show("D-series " + " Complete");
                    break;

                default:
                   
                    break;
            }
 */
        }


        private void fmMain_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'dsBatemanLabDB.QUANT_TYPE' table. You can move, or remove it, as needed.
            this.qUANT_TYPETableAdapter.Fill(this.dsBatemanLabDB.QUANT_TYPE);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.EQUIPMENT' table. You can move, or remove it, as needed.
            this.eQUIPMENTTableAdapter.Fill(this.dsBatemanLabDB.EQUIPMENT);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.ANTIBODY' table. You can move, or remove it, as needed.
            this.aNTIBODYTableAdapter.Fill(this.dsBatemanLabDB.ANTIBODY);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.ENZYME' table. You can move, or remove it, as needed.
            this.eNZYMETableAdapter.Fill(this.dsBatemanLabDB.ENZYME);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.FLUID_TYPE' table. You can move, or remove it, as needed.
            this.fLUID_TYPETableAdapter.Fill(this.dsBatemanLabDB.FLUID_TYPE);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.PROJECT' table. You can move, or remove it, as needed.
            this.pROJECTTableAdapter.Fill(this.dsBatemanLabDB.PROJECT);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.TIME_POINT' table. You can move, or remove it, as needed.
            this.tIME_POINTTableAdapter.Fill(this.dsBatemanLabDB.TIME_POINT);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.STUDY' table. You can move, or remove it, as needed.
            this.sTUDYTableAdapter.Fill(this.dsBatemanLabDB.STUDY);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.LAB_MEMBERS' table. You can move, or remove it, as needed.
            this.lAB_MEMBERSTableAdapter.Fill(this.dsBatemanLabDB.LAB_MEMBERS);

        }

        private void fmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            RegistryKey _key;
            _key = Registry.CurrentUser.CreateSubKey("SOFTWARE\\QLDReader");

            try
            {
                _key.SetValue(textBox_date_format.Name, textBox_date_format.Text);

                foreach (CheckBox _chbox in this.Controls.OfType<CheckBox>())
                {
                    _key.SetValue(_chbox.Name, _chbox.Checked);
                }

                foreach (ComboBox _com in this.groupBox_ExperInfo.Controls.OfType<ComboBox>())
                {
                    _key.SetValue(_com.Name, _com.Text);
                }
                foreach (NumericUpDown _numer in this.groupBox_ExperInfo.Controls.OfType<NumericUpDown>())
                {
                    _key.SetValue(_numer.Name, _numer.Value);
                }
                _key.SetValue("last_folder", _workFolder);
            }
            catch (Exception _ex)
            {
                MessageBox.Show(_ex.Message);
            }
        } //button_StartExport_Click 


        private void comboBox_SampleProcessBy_SelectedValueChanged(object sender, EventArgs e)
        {
            if (ExperimentInfo != null)
            {
                ExperimentInfo._SampleProcessBy = Convert.ToByte(((ComboBox)sender).SelectedValue);
            }
        }

        private void comboBox_DoneBy_SelectedValueChanged(object sender, EventArgs e)
        {
            if (ExperimentInfo != null)
            {
                ExperimentInfo._DoneBy = Convert.ToByte(((ComboBox)sender).SelectedValue);
            }

        }

        private void comboBox_QuantitatedBy_SelectedValueChanged(object sender, EventArgs e)
        {
            if (ExperimentInfo != null)
            {
                ExperimentInfo._QuantitatedBy = Convert.ToByte(((ComboBox)sender).SelectedValue);
            }

        }

        private void button_StartExport_Batch_Click(object sender, EventArgs e)
        {
            foreach (string _selectedFile in _FileList)
            {
                textBox_QLDfile.Text = _selectedFile;
                textBox_Excelfile.Text = Path.ChangeExtension(textBox_QLDfile.Text, "xlsx");
                getExperimentInfoType(Path.GetFileNameWithoutExtension(textBox_QLDfile.Text), checkBox_fileName.Checked);

                this.button_StartExport_Click(this, null);

            }
            MessageBox.Show("Batch completed");
        }


        private void button_EditExcelfileName_Click(object sender, EventArgs e)
        {
            SaveFileDialog _saveDlg = new SaveFileDialog();
            _saveDlg.Filter = "xlsx|*.xlsx";

            if (textBox_QLDfile.Text != "")
            {
                _saveDlg.InitialDirectory = Path.GetDirectoryName(textBox_QLDfile.Text);
            }

            if (_saveDlg.ShowDialog() == DialogResult.OK)
            {
                textBox_Excelfile.Text = _saveDlg.FileName;


            }
        }

        private void dateTimePicker_AssayDate_ValueChanged(object sender, EventArgs e)
        {
            if (ExperimentInfo != null)
            {
                ExperimentInfo._Date = ((DateTimePicker)sender).Text;
                ExperimentInfo.Update();
            }

        }

       

        private void comboBox_Fluid_SelectedValueChanged(object sender, EventArgs e)
        {
            if (ExperimentInfo != null)
            {
                ExperimentInfo._matrix = ((ComboBox)sender).Text;
                ExperimentInfo._matrixID = Convert.ToByte(((ComboBox)sender).SelectedValue);
                ExperimentInfo.Update();
            }

        }

        private void comboBox_Enzyme_SelectedValueChanged(object sender, EventArgs e)
        {
            if (ExperimentInfo != null)
            {
                ExperimentInfo._Enzyme = ((ComboBox)sender).Text;
                ExperimentInfo._EnzymeID = Convert.ToByte(((ComboBox)sender).SelectedValue);
                ExperimentInfo.Update();
            }

        }

        private void comboBox_Abody_SelectedValueChanged(object sender, EventArgs e)
        {
            if (ExperimentInfo != null)
            {
                ExperimentInfo._IP = ((ComboBox)sender).Text;
                ExperimentInfo._AbodyID = Convert.ToByte(((ComboBox)sender).SelectedValue);
                // MessageBox.Show( ExperimentInfo._IP = ((ComboBox)sender).Text );
                ExperimentInfo.Update();
            }    
            

        }

        private void comboBox_QuantType_SelectedValueChanged(object sender, EventArgs e)
        {
            if (ExperimentInfo != null)
            {
                ExperimentInfo._quanType = ((ComboBox)sender).Text;
                ExperimentInfo.Update();
            }

        }

        private void comboBox_Instrument_SelectedValueChanged(object sender, EventArgs e)
        {
            if (ExperimentInfo != null)
            {
                ExperimentInfo._Instument = ((ComboBox)sender).Text;
                ExperimentInfo._InstrumentID = Convert.ToByte(((ComboBox)sender).SelectedValue);
                ExperimentInfo.Update();
            }

        }

        
    }
}
