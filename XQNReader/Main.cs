using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;

using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

using Raw = WashU.BatemanLab.Tools.ThermoRawFileReader;

using System.Runtime.ExceptionServices;
using System.Security;


namespace XQNReader
{
    /// <summary> Main form class of XQN Reader </summary>
    public partial class fmMain : Form
    {
        /// <summary> Partial class of XQN Reader main window </summary>
        public fmMain()
        {
            InitializeComponent();
        }

        private XCALIBURFILESLib.XXQNClass _XXQN;

        private Excel.Application _xlApp;

        private string[] _FileList;

        string[] ConvertToStringArray(System.Array values)
        {

            // create a new string array
            string[] theArray = new string[values.Length];

            // loop through the 2-D System.Array and populate the 1-D String Array
            for (int i = 1; i <= values.Length; i++)
            {
                if (values.GetValue(i, 1) == null)
                    theArray[i - 1] = "";
                else
                    theArray[i - 1] = (string)values.GetValue(i, 1).ToString();
            }

            return theArray;
        }

        private string _CustomDateFormat;

        private class ExperimentInfoType
        {
            public string _Date = "noDate";
            public byte _DateInx = 0;
            public string _Subject = "noSubjectNo";
            public byte _SubjectInx = 1;
            public string _matrix = "noMatrix";
            public byte _matrixInx = 2;
            public string _IP = "noIP";
            public byte _IPInx = 3;
            public string _Enzyme = "noEnzyme";
            public byte _EnzymeInx = 4;
            public string _Instument = "noInstument";
            public byte _InstumentInx = 5;
            
            public string _quanType = "noQuanType";
            public byte _quanTypeInx = 6;
            public string _CommonChartNameStr;
            public string _ProcessedBy = "";
            public string _fileName = "";
            public string _expName = "";
            public int _totSampleCnt = 0;
            public byte TitleHeight = 5;
            public byte _UnknownCount = 0, _StdBracketCount = 0, _BlankCount = 0, _QCCount = 0, _TotalCount = 0;

            public byte _SamplePreparedBy_ID;
            public byte _SampleRanBy_ID;
            public byte _SampleQuantitatedBy_ID;

            public byte _Instrument_ID;
            public byte _Matrix_ID;

            public void reset()
            {
                _totSampleCnt = 0;
                _UnknownCount = 0;
                _StdBracketCount = 0;
                _BlankCount = 0;
                _QCCount = 0;
                _TotalCount = 0;
            }
        }

        private ExperimentInfoType ExperimentInfo;

        private void getExperimentInfoType(string _inputFileName, bool _parsefileName)
        {

            if (_inputFileName == "") { return; }

            if (ExperimentInfo == null)
            {
                ExperimentInfo = new ExperimentInfoType();

            }

            ExperimentInfo.reset();

            string[] _parameters;

            if (_inputFileName != null)
            {
                ExperimentInfo._fileName = _inputFileName;
                ExperimentInfo._expName = _inputFileName.Replace("_REL", "").Replace("_ABS", "");
                _parameters = _inputFileName.Split('_');


                if (_parsefileName)
                {
                    //_parameters = _inputFileName.Split('_');

                    try
                    {
                        dateTimePicker_AssayDate.Value = DateTime.ParseExact(_parameters[ExperimentInfo._DateInx],
                            _CustomDateFormat, null);


                        comboBox_Fluid.SelectedIndex = -1;
                        comboBox_Abody.SelectedIndex = -1;
                        comboBox_Enzyme.SelectedIndex = -1;
                        comboBox_Instrument.SelectedIndex = -1;
                        comboBox_QuantType.SelectedIndex = -1;



                        comboBox_Fluid.SelectedItem = ParseFluid(_parameters[ExperimentInfo._matrixInx]);
                        comboBox_Abody.SelectedItem = ParseIP(_parameters[ExperimentInfo._IPInx]);
                        comboBox_Enzyme.SelectedItem = ParseEnzym(_parameters[ExperimentInfo._EnzymeInx]);
                        comboBox_Instrument.SelectedItem = _parameters[ExperimentInfo._InstumentInx];
                        comboBox_QuantType.SelectedItem = _parameters[ExperimentInfo._quanTypeInx];
                        textBox_subject.Text = _parameters[ExperimentInfo._SubjectInx];
                    }
                    catch (FormatException)
                    {
                        MessageBox.Show("change custom date format");
                    }

                    catch (Exception)
                    {
                        
                    }
                }

                try
                {
                    dateTimePicker_AssayDate.Value = DateTime.ParseExact(_parameters[ExperimentInfo._DateInx],
                        _CustomDateFormat, null);
                    textBox_subject.Text = _parameters[ExperimentInfo._SubjectInx].Split('-')[1] ;
                }
                catch (Exception _exep)
                {
                    MessageBox.Show( _exep.Message, "getExperimentInfoType():", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }

            ExperimentInfo._Date = dateTimePicker_AssayDate.Value.ToShortDateString();
            ExperimentInfo._Subject = textBox_subject.Text;
            ExperimentInfo._matrix = comboBox_Fluid.Text;
            ExperimentInfo._Instument = comboBox_Instrument.Text;
            ExperimentInfo._IP = comboBox_Abody.Text;
            ExperimentInfo._Enzyme = comboBox_Enzyme.Text;
            ExperimentInfo._quanType = comboBox_QuantType.Text;


            ExperimentInfo._SamplePreparedBy_ID = Convert.ToByte(comboBox_SampleProcessBy.SelectedValue);
            ExperimentInfo._SampleRanBy_ID = Convert.ToByte(comboBox_DoneBy.SelectedValue);
            ExperimentInfo._SampleQuantitatedBy_ID = Convert.ToByte(comboBox_QuantitatedBy.SelectedValue);

            ExperimentInfo._Instrument_ID = Convert.ToByte(comboBox_Instrument.SelectedValue);
            
            ExperimentInfo._Matrix_ID = Convert.ToByte(comboBox_Fluid.SelectedValue);


            ExperimentInfo._CommonChartNameStr = String.Format
                    ("{0} {1} {2} {3} {4} ({5})",
                    ExperimentInfo._Subject,
                    ExperimentInfo._matrix,
                    ExperimentInfo._IP,
                    ExperimentInfo._Enzyme,
                    ExperimentInfo._Instument,
                    ExperimentInfo._Date);



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

        }

        private string ParseFluid(string _fluid)
        {
            string _tmp = _fluid;

            if (_fluid.Contains("Plasma"))
            {
                _tmp = "hPlasma";
            }

            if (_fluid.Contains("CSF"))
            {
                _tmp = "hCSF";
            }
            return _tmp;
        }

        private string ParseEnzym(string _enzym)
        {
            string _tmp = _enzym;

            if (_enzym.Contains("Tryp"))
            {
                _tmp = "Enzym-Tryp";
            }

            if (_enzym.Contains("LysC"))
            {
                _tmp = "Enzym-LysC";
            }

            if (_enzym.Contains("LysN"))
            {
                _tmp = "Enzym-LysN";
            }

            return _tmp;
        }

        private string ParseIP(string _ip)
        {
            string _tmp = _ip;

            if (_ip.Contains("HJ5X"))
            {
                _tmp = "IP-HJ5X";
            }
            /*
            if (_ip.Contains("CSF"))
            {
                _tmp = "hCSF";
            }*/
            return _tmp;
        }



        private string CurveTypeToStr(XCALIBURFILESLib.XCurveType _curveType)
        {
            string _curveTypeStr = "N/A";
            switch (_curveType)
            {
                case XCALIBURFILESLib.XCurveType.XAverageRF:
                    _curveTypeStr = "Average RF";
                    break;
                case XCALIBURFILESLib.XCurveType.XCubicSpline:
                    _curveTypeStr = "Cubic Spline";
                    break;
                case XCALIBURFILESLib.XCurveType.XLinear:
                    _curveTypeStr = "Linear";
                    break;
                case XCALIBURFILESLib.XCurveType.XLinearLogLog:
                    _curveTypeStr = "Linear Log-Log";
                    break;
                case XCALIBURFILESLib.XCurveType.XLocallyWeighted:
                    _curveTypeStr = "Locally Weighted";
                    break;
                case XCALIBURFILESLib.XCurveType.XPointToPoint:
                    _curveTypeStr = "Point-to-Point";
                    break;
                case XCALIBURFILESLib.XCurveType.XQuadratic:
                    _curveTypeStr = "Quadratic";
                    break;
                case XCALIBURFILESLib.XCurveType.XQuadraticLogLog:
                    _curveTypeStr = "Quadratic Log-Log";
                    break;
                default:
                    break;
            }
            return _curveTypeStr;
        }

        private string weightingTypeToStr(XCALIBURFILESLib.XWeightingType _weightingType)
        {
            string _weightingTypeStr = "N/A";
            switch (_weightingType)
            {
                case XCALIBURFILESLib.XWeightingType.XEqual:
                    _weightingTypeStr = "Equal";
                    break;
                case XCALIBURFILESLib.XWeightingType.XOneOverSSquared:
                    _weightingTypeStr = "1/s^2";
                    break;
                case XCALIBURFILESLib.XWeightingType.XOneOverX:
                    _weightingTypeStr = "1/X";
                    break;
                case XCALIBURFILESLib.XWeightingType.XOneOverXSquared:
                    _weightingTypeStr = "1/X^2";
                    break;
                case XCALIBURFILESLib.XWeightingType.XOneOverY:
                    _weightingTypeStr = "1/Y";
                    break;
                case XCALIBURFILESLib.XWeightingType.XOneOverYSquared:
                    _weightingTypeStr = "1/Y^2";
                    break;

                default:
                    break;
            }
            return _weightingTypeStr;
        }

        private string OriginTypeToStr(XCALIBURFILESLib.XOriginType _originType)
        {
            string _originTypeStr = "N/A";
            switch (_originType)
            {
                case XCALIBURFILESLib.XOriginType.XForce:
                    _originTypeStr = "Force";
                    break;
                case XCALIBURFILESLib.XOriginType.XIgnore:
                    _originTypeStr = "Ignore";
                    break;
                case XCALIBURFILESLib.XOriginType.XInclude:
                    _originTypeStr = "Include";
                    break;

                default:
                    break;
            }
            return _originTypeStr;
        }

        private string SampleTypeToStr(XCALIBURFILESLib.XSampleTypes _sampleType)
        {
            string _sampleTypeStr = "N/A";
            switch ((int)_sampleType)
            {
                case 0:
                    _sampleTypeStr = "Unknown";
                    break;
                case 1:
                    _sampleTypeStr = "Blank";
                    break;
                case 2:
                    _sampleTypeStr = "QC";
                    break;
                case 3:
                    _sampleTypeStr = "StdClear";
                    break;
                case 4:
                    _sampleTypeStr = "StdUpdate";
                    break;
                case 5:
                    _sampleTypeStr = "Std Bracket";
                    break;
                default:
                    break;
            }
            return _sampleTypeStr;
        }

        private int QuanTypeToDBInt(string _quanType)
        {
            int _intQuanType = 0;
            if (_quanType.Contains("REL")) { _intQuanType = 1; }
            if (_quanType.Contains("ABS")) { _intQuanType = 2; }
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


        private int SampleTypeToDBSampleType(XCALIBURFILESLib.XSampleTypes _sampleType)
        {
            int _sampleTypeint = -1;
            switch ((int)_sampleType)
            {
                case 0:
                    _sampleTypeint = 2; // "Unknown";
                    break;
                /*
            case 1:
                _sampleTypeint = "Blank";
                break;
            case 2:
                _sampleTypeint = "QC";
                break;
            case 3:
                _sampleTypeStr = "StdClear";
                break;
            case 4:
                _sampleTypeStr = "StdUpdate";
                break; */
                case 5:
                    _sampleTypeint = 1; // "Std Bracket";
                    break;
                default:
                    break;
            }
            return _sampleTypeint;
        }

        private string IntegTypeToStr(XCALIBURFILESLib.XIntegrationType _integType)
        {
            string _integTypeStr = "N/A";
            switch (_integType)
            {
                case XCALIBURFILESLib.XIntegrationType.XAutoMatic_MethodSettings:
                    _integTypeStr = "Method settings";
                    break;
                case XCALIBURFILESLib.XIntegrationType.XUser_Integration:
                    _integTypeStr = "User settings";
                    break;
                case XCALIBURFILESLib.XIntegrationType.XManual_Integration:
                    _integTypeStr = "Manual";
                    break;
                default:
                    break;
            }
            return _integTypeStr;
        }

        private double GetSlopeOfEquation(string _equation, string _type)
        {
            double _res = -1;
            Char SlopeSign = '+';

            if (_equation != null)
            {
                if (_equation.Contains('-') && !_equation.Contains("e-")) { SlopeSign = '-'; }
            }

            try
            {

                // MessageBox.Show(_equation.Substring(_equation.IndexOf(SlopeSign) + 1, _equation.IndexOf('*') - (_equation.IndexOf(SlopeSign) + 1)));

                if (_type == "Ignore")
                {

                    Char[] Sign = new Char[] { '+', '-' };
                    _res = Convert.ToDouble(_equation.Substring(_equation.LastIndexOfAny(Sign) + 1, _equation.IndexOf('*') - (_equation.LastIndexOfAny(Sign) + 1)));
                };

                if (_type == "Force")
                {
                    _res = Convert.ToDouble(_equation.Substring(_equation.IndexOf('=') + 1, _equation.IndexOf('*') - (_equation.IndexOf('=') + 1)));
                }

            }
            catch (Exception _ex)
            {
                if (_equation != null)
                {
                    MessageBox.Show(_equation + ": " + _ex.Message, "GetSlopeOfEquation():", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                    /*MessageBox.Show("Can't get the slope from Equation : " + _equation + "I got" +
                      _equation.Substring(_equation.IndexOf(SlopeSign) + 1, _equation.IndexOf('*') - (_equation.IndexOf(SlopeSign) + 1)));*/
                } 
                else 
                {
                    MessageBox.Show( _ex.Message, "GetSlopeOfEquation():", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            return SlopeSign == '-' ? _res * -1 : _res;
        }

        private double GetInterOfEquation(string _equation, string _type)
        {
            if (_equation != null)
            {
                double _res = -1;
                Char SlopeSign = '+';
                string _intersect = null;
                if (_equation.Contains('-') && !_equation.Contains("e-")) { SlopeSign = '-'; }
                try
                {
                    if (_type == "Ignore")
                    {
                        Char[] Sign = new Char[] { '+', '-' };
                        _intersect = _equation.Substring(_equation.IndexOf('=') + 1, _equation.LastIndexOfAny(Sign) - (_equation.IndexOf('=') + 1));

                        _res = Convert.ToDouble(_equation.Substring(_equation.IndexOf('=') + 1, _equation.LastIndexOfAny(Sign) - (_equation.IndexOf('=') + 1)));

                    }
                    if (_type == "Force")
                    {
                        _res = 0;
                    }
                }
                catch (FormatException _exFormat)
                {
                    if (checkBox_ShowError.Checked)
                    {
                        MessageBox.Show("Can't get the intersect in - " + _equation, "GetInterOfEquation():", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    if (checkBox_ShowDebug.Checked)
                    {
                        MessageBox.Show(_intersect, _exFormat.Message);
                    }
                }
                catch (Exception _ex)
                {
                    MessageBox.Show(_ex.Message, "GetInterOfEquation():", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return SlopeSign == '-' ? _res * -1 : _res;
            }
            else
            {
                return -1;
            }
        }

        private double GetROfEquation(string _equation)
        {
            double _res = -1;
            try
            {
                _res = Math.Sqrt(Convert.ToDouble(_equation.Substring(_equation.LastIndexOf('=') + 1, (_equation.IndexOf('W') - 1) - (_equation.LastIndexOf('=') + 1))));
            }
            catch (Exception _ex)
            {
                if (checkBox_ShowDebug.Checked)
                {
                    MessageBox.Show(_equation + ": " + _ex.Message, "GetROfEquation():", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                
            }
            return _res;
        }

        private double GetR_SQROfEquation(string _equation)
        {
            double _res = -1;
            try
            {
                _res = Convert.ToDouble(_equation.Substring(_equation.LastIndexOf('=') + 1, (_equation.IndexOf('W') - 1) - (_equation.LastIndexOf('=') + 1)));
            }
            catch (Exception _ex)
            {
                if (checkBox_ShowDebug.Checked)
                {
                    MessageBox.Show(_equation + ": " + _ex.Message, "GetR_SQROfEquation():", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            return _res;
        }

        private string GetEquaOfEquation(string _equation)
        {
            string _R = "";
            try
            {
                _R = _equation.Substring(0, _equation.IndexOf('R') - 3);
            }
            catch (Exception _ex)
            {
                if (checkBox_ShowDebug.Checked)
                {
                    MessageBox.Show(_equation + ": " + _ex.Message, "GetEquaOfEquation():", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return _R;
        }
        private string GetEquaOfEquation(string _equation, string _component)
        {
            string _R = "";
            try
            {
                _R = _equation.Substring(0, _equation.IndexOf('R') - 3);
            }
            catch (Exception _ex)
            {
                if (checkBox_ShowDebug.Checked)
                {
                    MessageBox.Show(_equation + ": " + _ex.Message, "GetEquaOfEquation():", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return _R;
        }
  
        private int GetTPfromDB(string _sampleID, string _sampleType)
        {
            int _intTP = -1; double _doubleTP = -1;

            try
            {
                if (_sampleType == "Standard")
                {
                    _doubleTP = Convert.ToDouble(_sampleID.Substring(_sampleID.LastIndexOf('#') + 1));
                }
                else
                {
                    if (_sampleType == "Unknown")
                    {
                        _doubleTP = Convert.ToDouble(_sampleID.Replace('^', '.'));

                        var query_TP_ID = tIME_POINTTableAdapter.GetData().Select("TIME_POINT_VALUE = "
                                                                      + _doubleTP.ToString()
                                                                      + "AND STUDY_ID = "
                                                                      + comboBox_ClinicalStudy.SelectedValue);

                        _intTP = Convert.ToInt32(query_TP_ID[0]["TIME_POINT_ID"]);
                    }

                }

            }
            catch (Exception _err)
            {
                MessageBox.Show(_err.Message + ": " + _sampleID, "GetLevel_TP(): Can't convert to Integer");
            }

            return _intTP;
        }

  
        private void ToolStripMenuItem_Exit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ToolStripMenuItem_About_Click(object sender, EventArgs e)
        {
            fmAbout _fmAbout = new fmAbout();
            _fmAbout.ShowDialog();
        }

        private void button_SelectXQN_Click(object sender, EventArgs e)
        {
            OpenFileDialog _openDlg = new OpenFileDialog();
            _openDlg.Filter = "XQN|*.XQN";
            _openDlg.Multiselect = true;

            if (_openDlg.ShowDialog() == DialogResult.OK)
            {



                if (checkBox_fileName.Checked == true)
                {
                    foreach (ComboBox _com in this.groupBox_ExperInfo.Controls.OfType<ComboBox>())
                    {
                        _com.Text = "";
                    }
                    textBox_subject.Text = "";
                }

                /*MessageBox.Show(_openDlg.FileNames.Count().ToString());*/
                _FileList = _openDlg.FileNames;

                textBox_XQNfile.Text = _openDlg.FileName;
                textBox_Excelfile.Text = Path.ChangeExtension(textBox_XQNfile.Text, "xlsx");

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



            getExperimentInfoType(Path.GetFileNameWithoutExtension(textBox_XQNfile.Text), checkBox_fileName.Checked);

            groupBox_ExperInfo.Visible = true;

        }

        private void button_EditExcelfileName_Click(object sender, EventArgs e)
        {
            SaveFileDialog _saveDlg = new SaveFileDialog();
            _saveDlg.Filter = "xlsx|*.xlsx";

            if (textBox_XQNfile.Text != "")
            {
                _saveDlg.InitialDirectory = Path.GetDirectoryName(textBox_XQNfile.Text);
            }

            if (_saveDlg.ShowDialog() == DialogResult.OK)
            {
                textBox_Excelfile.Text = _saveDlg.FileName;
            }
        }

        private void ExportLOAD()
        {

            List<string> _strExportSQL = new List<string>();
           
            string _strAnalyte = null;
            double _Area = 0, _ISTD_Area = 0, _Area_Ratio = 0, _CalcAmount = 0;
            decimal dcm_Area_Ratio = 0;
            string _strSpecAmount = null, _strRT = null, _strISTD_RT = null;
            string _response = null;

            if (textBox_XQNfile.Text == "")
            {
                MessageBox.Show("Select Input XQN file, please");
                return;
            }


            _XXQN = new XCALIBURFILESLib.XXQNClass();

            try
            {
                _XXQN.Open(textBox_XQNfile.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("File: " + textBox_XQNfile.Text + " is invalid", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {

            }

            XCALIBURFILESLib.IXHeader _IXHeader = (XCALIBURFILESLib.IXHeader)_XXQN.Header;

            ExperimentInfo._ProcessedBy = _IXHeader.ChangedId;

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = _IXHeader.ChangedLogon;
            }

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = "unknown";
            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = true;
            toolStripStatusLabelMain.Visible = true;



            _xlApp = new Excel.ApplicationClass();
            _xlApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;



            Excel.Workbook xlWB;

            try
            {

                xlWB =
                    _xlApp.Workbooks.Add(
                    //@"c:\Project\XQNReader\XQNReader\LOAD_TSQ_Template_Summary.xlsx"
                Application.StartupPath + Path.DirectorySeparatorChar +
                @"\Templates\LOAD_TSQ_Template_Summary.xlsx"
                );

            }

            catch (Exception)
            {
                MessageBox.Show("File: " + Application.StartupPath + Path.DirectorySeparatorChar +
                @"\Templates\LOAD_TSQ_Template_Summary.xlsx", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //MessageBox.Show(_exception.Message, _exception.GetBaseException().Message);
                return;
            }
            finally
            {

            }

            Excel.Worksheet _wsSummary = (Excel.Worksheet)xlWB.ActiveSheet;

            Excel.Worksheet _wsCurrent = null;


            byte Title = 5;
            byte _UnknownCount = 0, _StdBracketCount = 0, _TotalCount = 0;




            XCALIBURFILESLib.XBracketGroups _XBracketGroups
                = (XCALIBURFILESLib.XBracketGroups)_XXQN.BracketGroups;

            XCALIBURFILESLib.XBracketGroup _XBracketGroup
                = (XCALIBURFILESLib.XBracketGroup)_XBracketGroups.Item(1);

            XCALIBURFILESLib.XQuanResults _XQuanResults
                = (XCALIBURFILESLib.XQuanResults)_XBracketGroup.QuanResults;


            //********_XQuanResults*******************************
            foreach (XCALIBURFILESLib.XQuanResult _XQuanResult in _XQuanResults)
            {
                toolStripProgressBarMain.Maximum = _XQuanResults.Count;

                XCALIBURFILESLib.XSample _Sample
                    = (XCALIBURFILESLib.XSample)_XQuanResult.Sample;




                _TotalCount++;
                switch (_Sample.SampleType)
                {
                    case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                        _UnknownCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                        _StdBracketCount++;
                        break;
                    default:
                        break;
                }

                XCALIBURFILESLib.XQuanComponents _XQuanSelectedComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Selected; /*  Method; */
                //  = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.User;  
                ////////////////////////////////////////
                XCALIBURFILESLib.XQuanComponents _XQuanMethodComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Method; /*  Method; */


                ////////////////////////////////////////
                short _comNo = 1;
                foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponent in _XQuanSelectedComponents)
                {
                    XCALIBURFILESLib.XComponent _XComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanComponent.Component;

                    XCALIBURFILESLib.XQuanComponent _XQuanMethodComponent = (XCALIBURFILESLib.XQuanComponent)
                    _XQuanMethodComponents.Item(_comNo);
                    _comNo++;

                    XCALIBURFILESLib.XComponent _XMethodComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanMethodComponent.Component;

                    /*
                    XCALIBURFILESLib.XComponent _XComponentUser =
                        (XCALIBURFILESLib.XComponent)_XQuanComponent.Component;
                    */

                    //****************** 
                    XCALIBURFILESLib.IXDetection2 _IXDetection =
                        (XCALIBURFILESLib.IXDetection2)_XComponent.Detection;
                    //*****************

                    XCALIBURFILESLib.XQuanCalibration _XQuanCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XComponent.Calibration;

                    XCALIBURFILESLib.XQuanCalibration _XQuanMethodCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XMethodComponent.Calibration;

                    string _quanTypeKey = "";

                    // MessageBox.Show(ExperimentInfo._quanType.ToString());

                    if (ExperimentInfo._quanType == "RELATIVE")
                    {
                        _quanTypeKey = "N14:N15";
                    }
                    if (ExperimentInfo._quanType == "ABSOLUTE")
                    {
                        _quanTypeKey = "C13:C12";
                    }

                    //MessageBox.Show(_XQuanCalibration.ComponentType.ToString());
                    //MessageBox.Show(_XComponent.Name.Contains(_quanTypeKey).ToString());

                    if (_XQuanCalibration.ComponentType != XCALIBURFILESLib.XComponentType.XISTD
                        && !_XComponent.Name.Contains(_quanTypeKey))
                    {
                        _strAnalyte = _XComponent.Name.Replace("_", " "); //*******************

                        // ****** 2010-07-23
                        _strAnalyte = _strAnalyte.Replace("AB-Total", "AbetaTotal");
                        // ****** 2010-07-23


                        bool _SheetExist = false;
                        string _ComponentName = _XComponent.Name.Replace(":", "_");
                        _ComponentName = _ComponentName.Replace("-", "_");
                        //**********
                        _ComponentName = _ComponentName.Replace("(", "_");
                        _ComponentName = _ComponentName.Replace(")", "_");
                        //**********

                        foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
                        {

                            if (_wsComponent.Name == _ComponentName
                                )
                            {
                                _wsCurrent = _wsComponent;
                                _SheetExist = true;
                            }
                        }



                        if (!_SheetExist)
                        {
                            _wsCurrent = (Excel.Worksheet)xlWB.Sheets.Add
                          (Type.Missing, Type.Missing, Type.Missing,
                          Application.StartupPath + Path.DirectorySeparatorChar +
                            @"\Templates\LOAD_TSQ_Template.xlsx"
                          );

                            _wsCurrent.Name = _ComponentName;



                            _wsCurrent.Cells[2, 1] = _XComponent.Name.Replace("_", " ");

                            _wsCurrent.Cells[2, 3] =
                                CurveTypeToStr(_XQuanMethodCalibration.CurveType);
                            _wsCurrent.Cells[2, 4] =
                                weightingTypeToStr(_XQuanMethodCalibration.Weighting);
                            _wsCurrent.Cells[2, 6] =
                                OriginTypeToStr(_XQuanMethodCalibration.OriginType);

                            _wsCurrent.Cells[2, 8] = _XQuanMethodCalibration.FullEquation;

                        }



                        _wsCurrent.Cells[_Sample.RowNumber + Title, 1] = _Sample.FileName;
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 2] = SampleTypeToStr(_Sample.SampleType);
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 3] = IntegTypeToStr(_XQuanComponent.Selected_Type);


                        XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeak;


                        //************************************
                        /*
                        if (_Sample.RowNumber > 23 && _ComponentName.Contains("42") )
                        {
                            MessageBox.Show(_IXDetection.PeakWidthHeight.ToString()
                                + "  " + _IXDetection.SNThreshold.ToString());
                        }
                        */
                        //************************************

                        bool _IsPeakFoundANALYTE = false;
                        bool _IsPeakFoundISTD = false;


                        if (_XQuanComponent.DoesDetectedQuanPeakExist > 0)
                        {
                            _IsPeakFoundANALYTE = true;
                            _XDetectedQuanPeak =
                                (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponent.DetectedQuanPeak;

                            _wsCurrent.Cells[_Sample.RowNumber + Title, 4] = _XDetectedQuanPeak.Area;
                            _Area = _XDetectedQuanPeak.Area;

                            //_XDetectedQuanPeak.ISTDValid

                            _wsCurrent.Cells[_Sample.RowNumber + Title, 8] = _XDetectedQuanPeak.CalcAmount;
                            _CalcAmount = _XDetectedQuanPeak.CalcAmount;

                            if (_XDetectedQuanPeak.ISTDArea != 0)
                            {
                                _IsPeakFoundISTD = true;
                                _wsCurrent.Cells[_Sample.RowNumber + Title, 5] = _XDetectedQuanPeak.ISTDArea;
                                _wsCurrent.Cells[_Sample.RowNumber + Title, 6] =
                                _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;

                                _ISTD_Area = _XDetectedQuanPeak.ISTDArea;
                                _Area_Ratio = _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;
                                dcm_Area_Ratio = (decimal)_XDetectedQuanPeak.Area / (decimal)_XDetectedQuanPeak.ISTDArea;
                            }
                            else
                            {
                                _wsCurrent.Cells[_Sample.RowNumber + Title, 5] = "N/F";
                                //_wsCurrent.Cells[_Sample.RowNumber + Title, 6] = "-";
                            }



                            //string _response = "";
                            _response = "NULL"; _strRT = "NULL"; _strISTD_RT = "NULL";

                            if (_XDetectedQuanPeak.IsResponseOK > 0) { _response = "Ok"; }
                            if (_XDetectedQuanPeak.IsResponseHigh > 0) { _response = "High"; }
                            if (_XDetectedQuanPeak.IsResponseLow > 0) { _response = "Low"; }
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 9] = _response;
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 12] = _XDetectedQuanPeak.ApexRT;

                            _strRT = _XDetectedQuanPeak.ApexRT.ToString("F2");
                            //_RT = _XDetectedQuanPeak.ApexRT;
                        }
                        else
                        {
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 9] = "Not Found";

                        }



                        string _sample_ID = "N/A";
                        double _specAmount = 0;
                        switch (_Sample.SampleType)
                        {
                            case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                                _sample_ID = _Sample.Level;
                                XCALIBURFILESLib.XCalibrationCurveData _XCalibrationCurveData =
                                (XCALIBURFILESLib.XCalibrationCurveData)_XQuanComponent.CalibrationCurveData;
                                XCALIBURFILESLib.XReplicateRows _XReplicateRows =
                                (XCALIBURFILESLib.XReplicateRows)_XCalibrationCurveData.ReplicateRows;

                                foreach (XCALIBURFILESLib.XReplicateRow _XReplicateRow in _XReplicateRows)
                                {
                                    if (_XReplicateRow.LevelName == _Sample.Level)
                                    {
                                        _specAmount = _XReplicateRow.Amount;
                                        _strSpecAmount = "'" + _XReplicateRow.Amount.ToString("F5") + "'";
                                    }
                                }

                                break;
                            case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                                _sample_ID = _Sample.SampleId;
                                _specAmount = -1;
                                _strSpecAmount = "NULL";
                                break;
                            default:
                                break;
                        }


                        bool _IsSampleIDDefinded = false;
                        int _intTP = -1;

                        if (_sample_ID == "")
                        {
                            _sample_ID = _Sample.FileName.Substring(_Sample.FileName.LastIndexOf('_') + 1);
                        };

                        try
                        {
                            _intTP = Convert.ToInt16(_sample_ID);
                        }
                        catch (System.FormatException _FoEx)
                        {
                            MessageBox.Show(_FoEx.Message);
                        };

                        if (_intTP < 0)
                        {
                            _sample_ID = "";
                            _IsSampleIDDefinded = false;
                        }
                        else
                        {
                            _sample_ID = _intTP.ToString();
                            _IsSampleIDDefinded = true;
                        }


                        if (_specAmount != -1)
                        {
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 7] = _specAmount;
                        }
                        else
                        {
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 7] = "";
                        }
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 10] = _sample_ID;



                        foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponentISTD in _XQuanSelectedComponents)
                        {
                            XCALIBURFILESLib.XComponent _XISTD =
                            (XCALIBURFILESLib.XComponent)_XQuanComponentISTD.Component;

                            if (_XISTD.Name == _XQuanCalibration.ISTD)
                            {
                                if (_XQuanComponentISTD.DoesDetectedQuanPeakExist > 0)
                                {
                                    XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeakISTD =
                                        (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponentISTD.DetectedQuanPeak;
                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 13] = _XDetectedQuanPeakISTD.ApexRT;

                                    _strISTD_RT = _XDetectedQuanPeakISTD.ApexRT.ToString("F2");
                                }
                                else
                                {
                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 13] = "N/F";
                                }

                            }

                        }

                        // if (GetSlopeOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType)) == -1)
                        // {
                        //     return;
                        // }

                        string _subjectNum = "";

                        if (checkBox_if_multi_subject.Checked && (int)_Sample.SampleType == 0)
                        {
                            _subjectNum = _Sample.FileName.Split('_')[1];
                        }
                        else
                        {
                            if (ExperimentInfo._Subject.Contains("-"))
                            {
                                _subjectNum = ExperimentInfo._Subject.Split('-')[1];
                            }
                            else _subjectNum = ExperimentInfo._Subject;

                        }


                        if (_IsPeakFoundANALYTE && _IsPeakFoundISTD && _IsSampleIDDefinded && checkBox_ExportIntoDB.Checked)
                        {
                            _strExportSQL.Add("INSERT INTO TMP_TSQ_IMPORT ( ANYLYTE_NAME, FILENAME_DESC, SAMPLE_TYPE_ID, INTEG_TYPE_ID, AREA, ISTD_AREA, AREA_RATIO, " +
                                                                          "SPEC_AMOUNT, CALC_AMOUNT, LEVEL_TP, PEAK_SATUS, RT, ISTD_RT, QUAN_TYPE_ID, FLUID_TYPE_ID, " +
                                                                          "SUBJECT_NUM, PROCESS_BY_ID, ASSAY_DATE, ASSAY_TYPE_ID, ANTIBODY_ID, ENZYME_ID, DONE_BY, " +
                                                                          "SAMPLE_PROCESS_BY, EXPER_FILE, CURVE, WEIGHTING, ORIGIN, EQUATION, R, R_SQR, NUM_SLOPE, " +
                                                                          "NUM_INTERSECT )");
                            _strExportSQL.Add(String.Format("VALUES ('{0}', '{1}', {2}, '{3}', {4:F0}, {5:F0}, {6:F5}, {7:F}, {8:F5}, {9}, '{10}', {11}, {12}, {13}, {14}, '{15}', '{16}', '{17}', {18}, {19}, {20}, {21}, {22}, '{23}'," +
                                " '{24}', '{25}', '{26}', '{27}', {28:F5}, {29:F5}, {30:F5}, {31:F5});",
                                   _strAnalyte,                                 // 0
                                   _Sample.FileName,                            // 1
                                   (int)_Sample.SampleType,                     // 2
                                   (int)_XQuanComponent.Selected_Type,          // 3
                                   _Area,                                       // 4
                                   _ISTD_Area,                                  // 5
                                   _Area_Ratio,                                 // 6
                                   _strSpecAmount,                              // 7 
                                   _CalcAmount,                                 // 8
                                   _sample_ID,                                  // 9
                                   _response,                                   // 10
                                   _strRT,                                      // 11 
                                   _strISTD_RT,                                 // 12
                                   QuanTypeToDBInt(ExperimentInfo._quanType),
                                   FluidTypeToDBInt(ExperimentInfo._matrix),    // 14
                                   _subjectNum,       // 15
                                   ExperimentInfo._ProcessedBy,                 // 16
                                   ExperimentInfo._Date,                        // 17
                                   2,                                           // 18      { ASSAY_TYPE_ID   2 - "IP LC/MS (C-term)"    
                                   IPToDBInt(ExperimentInfo._IP),               // 19      { 3, // A-BODY  3 - HJ5.x
                                   EnzymeToDBInt(ExperimentInfo._Enzyme),       // 20      { 2, //Enzyme 2 - LysN
                                   1,                                           // 21      { Done by 1 - ovodvo
                                   1,                                           // 22      { Sample process by 1 - ovodvo
                                   ExperimentInfo._expName,                     // 23

                                   CurveTypeToStr(_XQuanMethodCalibration.CurveType),
                                   weightingTypeToStr(_XQuanMethodCalibration.Weighting),
                                   OriginTypeToStr(_XQuanMethodCalibration.OriginType),
                                   GetEquaOfEquation(_XQuanMethodCalibration.FullEquation),
                                   GetROfEquation(_XQuanMethodCalibration.FullEquation),
                                   GetR_SQROfEquation(_XQuanMethodCalibration.FullEquation),
                                   GetSlopeOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType)),
                                   GetInterOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType))


                                   ));

                        }
                    }
                }

                toolStripProgressBarMain.PerformStep();
                toolStripStatusLabelMain.Text = "Processing ..." + _Sample.FileName;
                toolStripStatusLabelMain.PerformClick();

            }

            //_xlApp.Visible = true;
            toolStripProgressBarMain.Visible = false;
            toolStripProgressBarMain.Value = 0;
            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();
            toolStripProgressBarMain.Visible = true;

            toolStripProgressBarMain.Maximum = xlWB.Worksheets.Count - 1;


            string _AppPath = Application.StartupPath + Path.DirectorySeparatorChar;

            string[] _sql_P1 = File.ReadAllLines(_AppPath + "sql_TSQImport_P1.sql");
            string[] _sql_P2 = File.ReadAllLines(_AppPath + "sql_TSQImport_P2.sql");
            string[] _sql_P3 = File.ReadAllLines(_AppPath + "sql_TSQImport_P3.sql");

            _strExportSQL.Insert(0, " ");
            _strExportSQL.InsertRange(0, _sql_P1.ToList<string>());
            _strExportSQL.AddRange(_sql_P3.ToList<string>());
            //_strExportSQL.AddRange(_sql_P2.ToList<string>());


            File.WriteAllLines(Path.ChangeExtension(textBox_Excelfile.Text, "sql"), _strExportSQL.ToArray());

            foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
            {
                _wsComponent.Name = "_" + _wsComponent.Name;

                toolStripStatusLabelMain.Text = "Formatting ..." + _wsComponent.Name;
                toolStripStatusLabelMain.PerformClick();

                if (_wsComponent.Name.Contains("Summary")) { continue; }

                Excel.ChartObjects _xlCharts = (Excel.ChartObjects)_wsComponent.ChartObjects(Type.Missing);
                string _componentName = _wsComponent.get_Range("A2", "A2").Value2.ToString();

                Excel.Range _xlRange_Compound,
                            _xlRange_StdBracket_Area,
                            _xlRange_StdBracket_ISArea,
                            _xlRange_StdBracket_SpecifiedAmount,
                            _xlRange_StdBracket_Ratio,
                            _xlRange_Unknown_Area,
                            _xlRange_Unknown_ISArea,
                            _xlRange_Unknown_CalcAmount,
                            _xlRange_Unknown_SampleID,
                            _xlRange_STD,
                            _xlRange_Unknowns,
                            _xlRange_All;

                _xlRange_Compound = _wsComponent.get_Range("A2", "A2");

                //All
                _xlRange_All =
                    _wsComponent.get_Range
                    ("A" + Title.ToString(), "M" + (Title - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_All.Name = _wsComponent.Name + "_All";


                //Std
                if (_StdBracketCount > 0)
                {
                    _xlRange_STD =
                        _wsComponent.get_Range
                        ("A" + Title.ToString(), "M" + (Title - 1 + _StdBracketCount).ToString());
                    _xlRange_STD.Name = _wsComponent.Name + "_STD";


                    _xlRange_StdBracket_Ratio =
                        _wsComponent.get_Range
                        ("F" + Title.ToString(), "F" + (Title - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_Ratio.Name = _wsComponent.Name + "_StdBracket_Area";

                    _xlRange_StdBracket_Area =
                        _wsComponent.get_Range
                        ("D" + Title.ToString(), "D" + (Title - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_Area.Name = _wsComponent.Name + "_StdBracket_Area";

                    _xlRange_StdBracket_ISArea =
                        _wsComponent.get_Range
                        ("E" + Title.ToString(), "E" + (Title - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_ISArea.Name = _wsComponent.Name + "_StdBracket_ISArea";

                    _xlRange_StdBracket_SpecifiedAmount =
                       _wsComponent.get_Range
                       ("G" + Title.ToString(), "G" + (Title - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_SpecifiedAmount.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount";


                    _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                    _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                    _xlRange_STD.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);

                    //_xlRange_StdBracket_ISArea.FormatConditions.AddDatabar();


                    //*********Sort Range************
                    _xlRange_STD.Sort("Column J",
                        Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                        Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                        Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                        Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                        Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                        Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                        Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                        Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                        Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);




                    //****************
                    // StdCurve;
                    Excel.ChartObject _xlChartStdCurve = (Excel.ChartObject)_xlCharts.Add(
                                       490, 800, 470, 250);
                    Excel.Chart _xlChartStdCurvePage = _xlChartStdCurve.Chart;
                    _xlChartStdCurvePage.ChartType = Excel.XlChartType.xlXYScatter;

                    _xlChartStdCurvePage.HasTitle = true;
                    _xlChartStdCurvePage.ChartTitle.Font.Size = 11;

                    string _chartTitleStdCurve = String.Format
                        ("Calibration curve {0} {1} {2} {3} ({4})",
                        _componentName,
                        ExperimentInfo._IP,
                        ExperimentInfo._Enzyme,
                        ExperimentInfo._Instument,
                        ExperimentInfo._Date);
                    //_xlSeries_TTR.Name = _componentName;

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
                    _xlSerie_Stds.XValues = _xlRange_StdBracket_SpecifiedAmount.Cells;
                    _xlSerie_Stds.Values = _xlRange_StdBracket_Ratio.Cells;


                    string _serieStdName = "H:L Abeta";

                    string _axeStdXNumFormat = "0.00",
                           _axeStdYNumFormat = "0.00";

                    string _axeStdXCaption = "%, Predicted labeling Abeta",
                           _axeStdYCaption = "%, Measured labeling Abeta";

                    if (ExperimentInfo._quanType == "ABS")
                    {
                        _axeStdXNumFormat = "0.0";
                        _axeStdXNumFormat = "0.00";

                        _axeStdXCaption = "ng/mL, Specified amount";
                        _axeStdYCaption = "AreaRatio";

                        _serieStdName = "Analyte:ISTD Ratio";
                    }

                    _xlAxisX_Std.TickLabels.NumberFormat = _axeStdXNumFormat;
                    _xlAxisY_Std.TickLabels.NumberFormat = _axeStdYNumFormat;
                    _xlAxisX_Std.AxisTitle.Caption = _axeStdXCaption;
                    _xlAxisY_Std.AxisTitle.Caption = _axeStdYCaption;

                    _xlSerie_Stds.Name = _serieStdName;

                    Excel.Trendlines _xlTrendlines_Stds = (Excel.Trendlines)_xlSerie_Stds.Trendlines(Type.Missing);

                    Excel.XlTrendlineType _XlTrendlineType =
                        Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear;
                    string _CurveIndex = _wsComponent.get_Range("C2", "C2").Value2.ToString();

                    if (_CurveIndex == "Quadratic")
                    { _XlTrendlineType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlExponential; }


                    _xlTrendlines_Stds.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, true, true, Type.Missing);

                    //****************

                };


                //Unknowns
                _xlRange_Unknowns =
                    _wsComponent.get_Range
                    ("A" + (Title + _StdBracketCount).ToString(), "M" + (Title - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknowns.Name = _wsComponent.Name + "_Unknowns";

                _xlRange_Unknown_Area =
                    _wsComponent.get_Range
                    ("D" + (Title + _StdBracketCount).ToString(), "D" + (Title - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_Area.Name = _wsComponent.Name + "_Unknown_Area";

                _xlRange_Unknown_ISArea =
                    _wsComponent.get_Range
                    ("E" + (Title + _StdBracketCount).ToString(), "E" + (Title - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_ISArea.Name = _wsComponent.Name + "_Unknown_ISArea";

                string strCalcOrRAtio = "H";
                if (_StdBracketCount == 0) { strCalcOrRAtio = "F"; }
                _xlRange_Unknown_CalcAmount =
                    _wsComponent.get_Range
                    (strCalcOrRAtio + (Title + _StdBracketCount).ToString(), strCalcOrRAtio + (Title - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_CalcAmount.Name = _wsComponent.Name + "_Unknown_CalcAmount";

                _xlRange_Unknown_SampleID =
                    _wsComponent.get_Range
                    ("J" + (Title + _StdBracketCount).ToString(), "J" + (Title - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_SampleID.Name = _wsComponent.Name + "_Unknown_SampleID";







                //*************************Format Ranges**************************
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;

                //******Sort Unknowns
                _xlRange_Unknowns.Sort("Column J",
                    Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                    Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);

                //*************2011-07-19**************************************

                //_xlRange_Unknown_ISArea.FormatConditions.AddDatabar();

                Excel.ColorScale cfColorScale = (Excel.ColorScale)(_xlRange_Unknown_ISArea.FormatConditions.AddColorScale(3));

                // Set the minimum threshold to red (0x000000FF) and maximum threshold
                // to blue (0x00FF0000).
                cfColorScale.ColorScaleCriteria[1].FormatColor.Color = 7039480;  //0x000000FF;  
                cfColorScale.ColorScaleCriteria[2].FormatColor.Color = 8711167;  // 0x00FF0000;
                cfColorScale.ColorScaleCriteria[3].FormatColor.Color = 13011546; // 0x00FF0000;

                //*************************Charts*********************************
                //

                Excel.ChartObject _xlChartTimeCorse = (Excel.ChartObject)_xlCharts.Add(
                                   0, 800, 470, 250);
                Excel.Chart _xlChartTimeCorsePage = _xlChartTimeCorse.Chart;
                _xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlXYScatter;
                _xlChartTimeCorsePage.HasTitle = true;
                _xlChartTimeCorsePage.ChartTitle.Font.Size = 11;



                _xlChartTimeCorsePage.ChartTitle.Caption = _componentName + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_TP = (Excel.Axes)_xlChartTimeCorsePage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_TP.HasMajorGridlines = true;
                _xlAxisX_TP.HasTitle = true;
                _xlAxisY_TP.HasTitle = true;

                _xlChartTimeCorsePage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartTimeCorsePage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_TimeCourse = (Excel.SeriesCollection)_xlChartTimeCorsePage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_TTR = (Excel.Series)_xlSeriesColl_TimeCourse.NewSeries();
                _xlSeries_TTR.XValues = _xlRange_Unknown_SampleID.Cells;
                _xlSeries_TTR.Values = _xlRange_Unknown_CalcAmount;


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

                if (_StdBracketCount == 0)
                {
                    _axeCurveYCaption = "Area Ratio";
                }

                _xlAxisX_TP.TickLabels.NumberFormat = _axeCurveXNumFormat;
                _xlAxisY_TP.TickLabels.NumberFormat = _axeCurveYNumFormat;
                _xlAxisX_TP.AxisTitle.Caption = _axeCurveXCaption;
                _xlAxisY_TP.AxisTitle.Caption = _axeCurveYCaption;

                _xlSeries_TTR.Name = _serieCurveName;
                _xlSeries_TTR.Name = _componentName;








                Excel.ChartObjects _xlChartsSummary = (Excel.ChartObjects)_wsSummary.ChartObjects(Type.Missing);
                Excel.ChartObject _xlChartCurveSummary;
                Excel.Chart _xlChartCurveSummaryPage;

                if (_xlChartsSummary.Count < 1)
                {
                    _xlChartCurveSummary = (Excel.ChartObject)_xlChartsSummary.Add(
                                   10, 10, 850, 500);
                    _xlChartCurveSummaryPage = _xlChartCurveSummary.Chart;
                    _xlChartCurveSummaryPage.ChartType = Excel.XlChartType.xlXYScatter;
                    _xlChartCurveSummaryPage.HasTitle = true;
                    _xlChartCurveSummaryPage.ChartTitle.Font.Size = 11;



                    Excel.Axes _xlAxes_CurveSummary = (Excel.Axes)_xlChartCurveSummaryPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisX_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisY_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                    _xlAxisX_CurveSummary.HasMajorGridlines = true;
                    _xlAxisY_CurveSummary.HasMajorGridlines = true;

                    _xlAxisX_CurveSummary.Format.Line.Weight = 3.0F;
                    _xlAxisY_CurveSummary.Format.Line.Weight = 3.0F;

                    _xlAxisX_CurveSummary.HasTitle = true;
                    _xlAxisY_CurveSummary.HasTitle = true;

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
                _xlSeries_CurveTTR.XValues = _xlRange_Unknown_SampleID.Cells;
                _xlSeries_CurveTTR.Values = _xlRange_Unknown_CalcAmount;
                _xlSeries_CurveTTR.Name = (string)_xlRange_Compound.Cells.Value2;


                toolStripProgressBarMain.PerformStep();

            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = false;

            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();

            _XXQN.Close();

            object misValue = System.Reflection.Missing.Value;

            try
            {
                _XXQN.Close();

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
                    ((Excel._Worksheet)_wsSummary).Activate();
                    
                }
                else
                {
                    _xlApp.Quit();
                }
            }
        }

        private void ExportXcaliburUni_before_May_2016()
        {
            
            List<string> _strExportSQL = new List<string>();

            string _strAnalyte = null;
            double _Area = 0, _ISTD_Area = 0, _AreaRatio = 0, _CalcAmount = 0;
            decimal dcm_AreaRatio = 0;
            string _strSpecAmount = null, _strRT = null, _strISTD_RT = null;
            string _response = null;

            if (textBox_XQNfile.Text == "")
            {
                MessageBox.Show("Select Input XQN file, please");
                return;
            }

            _XXQN = new XCALIBURFILESLib.XXQNClass();

            try
            {
                _XXQN.Open(textBox_XQNfile.Text);
            }
            catch (Exception _exception)
            {
                MessageBox.Show("File: " + textBox_XQNfile.Text + " is invalid XQN-file", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                if (checkBox_ShowDebug.Checked)
                {
                    MessageBox.Show(_exception.Message);
                }
                return;
            }
            finally
            {

            }

            
            XCALIBURFILESLib.IXHeader _IXHeader = (XCALIBURFILESLib.IXHeader)_XXQN.Header;

            ExperimentInfo._ProcessedBy = _IXHeader.ChangedId;

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = _IXHeader.ChangedLogon;
            }

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = "unknown";
            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = true;
            toolStripStatusLabelMain.Visible = true;
            
            _xlApp = new Excel.ApplicationClass();
            _xlApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;
            
            Excel.Workbook xlWB;
            
            try
            {

                xlWB =
                    _xlApp.Workbooks.Add(
                                      Application.StartupPath + Path.DirectorySeparatorChar +
                                      @"\Templates\LOAD_TSQ_Template_Summary.xlsx"
                );

            }

            catch (Exception)
            {
                MessageBox.Show("File: " + Application.StartupPath + Path.DirectorySeparatorChar +
                @"\Templates\LOAD_TSQ_Template_Summary.xlsx", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {

            }

            Excel.Worksheet _wsSummary = (Excel.Worksheet)xlWB.ActiveSheet;

            Excel.Worksheet _wsCurrent = null ;

            byte SkipFirstN = ExperimentInfo.TitleHeight; /*Title height of template in rows*/

            byte _UnknownCount = 0, _StdBracketCount = 0, _BlankCount = 0, _QCCount = 0, _TotalCount = 0;

                       
            XCALIBURFILESLib.XBracketGroups _XBracketGroups
                = (XCALIBURFILESLib.XBracketGroups)_XXQN.BracketGroups;

            XCALIBURFILESLib.XBracketGroup _XBracketGroup
                = (XCALIBURFILESLib.XBracketGroup)_XBracketGroups.Item(1);

            XCALIBURFILESLib.XQuanResults _XQuanResults
                = (XCALIBURFILESLib.XQuanResults)_XBracketGroup.QuanResults;


            //********_XQuanResults*******************************
            foreach (XCALIBURFILESLib.XQuanResult _XQuanResult in _XQuanResults)
            {
                toolStripProgressBarMain.Maximum = _XQuanResults.Count;

                XCALIBURFILESLib.XSample _Sample
                    = (XCALIBURFILESLib.XSample)_XQuanResult.Sample;


                ExperimentInfo._totSampleCnt++;

                _TotalCount++;
                switch (_Sample.SampleType)
                {
                    case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                        _UnknownCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                        _StdBracketCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleBlank:
                        _BlankCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleQC:
                        _QCCount++;
                        break;

                    default:
                        MessageBox.Show("Unususal sample type " + _Sample.SampleType.ToString());
                        break;
                }


                XCALIBURFILESLib.XQuanComponents _XQuanSelectedComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Selected;

                XCALIBURFILESLib.XQuanComponents _XQuanMethodComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Method;


                short _comNo = 1;
                foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponent in _XQuanSelectedComponents)
                {
                    XCALIBURFILESLib.XComponent _XComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanComponent.Component;

                    XCALIBURFILESLib.XQuanComponent _XQuanMethodComponent = (XCALIBURFILESLib.XQuanComponent)
                    _XQuanMethodComponents.Item(_comNo);
                    _comNo++;

                    XCALIBURFILESLib.XComponent _XMethodComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanMethodComponent.Component;

                    XCALIBURFILESLib.IXDetection2 _IXDetection =
                        (XCALIBURFILESLib.IXDetection2)_XComponent.Detection;

                    XCALIBURFILESLib.XQuanCalibration _XQuanCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XComponent.Calibration;

                    XCALIBURFILESLib.XQuanCalibration _XQuanMethodCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XMethodComponent.Calibration;

                    string _quanTypeKey = "All";
                    
                    if (ExperimentInfo._quanType == "RELATIVE")
                    {
                        _quanTypeKey = "N14:N15";
                    }
                    if (ExperimentInfo._quanType == "ABSOLUTE")
                    {
                        _quanTypeKey = "C13:C12";
                    }

                    
                    if (_XQuanCalibration.ComponentType != XCALIBURFILESLib.XComponentType.XISTD
                        && !_XComponent.Name.Contains(_quanTypeKey))
                    {
                        _strAnalyte = _XComponent.Name.Replace("_", " "); 
                                                
                        _strAnalyte = _strAnalyte.Replace("AB-Total", "AbetaTotal");
                        
                        bool _SheetExist = false;
                        string _ComponentName = _XComponent.Name.Replace(":", "_");
                        _ComponentName = _ComponentName.Replace("-", "_");

                        _ComponentName = _ComponentName.Replace("(", "_");
                        _ComponentName = _ComponentName.Replace(")", "_");


                        foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
                        {

                            if (_wsComponent.Name == _ComponentName
                                )
                            {
                                _wsCurrent = _wsComponent;
                                _SheetExist = true;
                            }
                        }



                        if (!_SheetExist)
                        {
                            _wsCurrent = (Excel.Worksheet)xlWB.Sheets.Add
                          (Type.Missing, Type.Missing, Type.Missing,
                          Application.StartupPath + Path.DirectorySeparatorChar +
                            @"\Templates\LOAD_TSQ_Template.xlsx"
                          );

                            _wsCurrent.Name = _ComponentName;


                            _wsCurrent.Cells[2, 1] = _XComponent.Name.Replace("_", " ");

                            _wsCurrent.Cells[2, 3] =
                                CurveTypeToStr(_XQuanMethodCalibration.CurveType);
                            _wsCurrent.Cells[2, 4] =
                                weightingTypeToStr(_XQuanMethodCalibration.Weighting);
                            _wsCurrent.Cells[2, 6] =
                                OriginTypeToStr(_XQuanMethodCalibration.OriginType);

                            _wsCurrent.Cells[2, 8] = _XQuanMethodCalibration.FullEquation;

                        }


                        int _currRowN = _Sample.RowNumber + SkipFirstN;

                        _wsCurrent.Cells[_currRowN, 1] = _Sample.FileName;
                        _wsCurrent.Cells[_currRowN, 2] = SampleTypeToStr(_Sample.SampleType);
                        _wsCurrent.Cells[_currRowN, 3] = IntegTypeToStr(_XQuanComponent.Selected_Type);


                        XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeak;


                        _XDetectedQuanPeak = null;

                        bool _IsPeakFoundANALYTE = false;
                        bool _IsPeakFoundISTD = false;



                        if (_XQuanComponent.DoesDetectedQuanPeakExist > 0)
                        {
                            _IsPeakFoundANALYTE = true;
                            _XDetectedQuanPeak =
                                (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponent.DetectedQuanPeak;


                            _wsCurrent.Cells[_currRowN, 4] = _XDetectedQuanPeak.Area;
                            _Area = _XDetectedQuanPeak.Area;


                            _wsCurrent.Cells[_currRowN, 8] = _XDetectedQuanPeak.CalcAmount;
                            _CalcAmount = _XDetectedQuanPeak.CalcAmount;

                            if (_XDetectedQuanPeak.ISTDArea != 0)
                            {
                                _IsPeakFoundISTD = true;
                                _wsCurrent.Cells[_currRowN, 5] = _XDetectedQuanPeak.ISTDArea;
                                _wsCurrent.Cells[_currRowN, 6] =
                                _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;

                                _ISTD_Area = _XDetectedQuanPeak.ISTDArea;
                                _AreaRatio = _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;
                                dcm_AreaRatio = (decimal)_XDetectedQuanPeak.Area / (decimal)_XDetectedQuanPeak.ISTDArea;
                            }
                            else
                            {
                                _wsCurrent.Cells[_currRowN, 5] = "N/F";

                            }


                            _response = "NULL"; _strRT = "NULL"; _strISTD_RT = "NULL";

                            if (_XDetectedQuanPeak.IsResponseOK > 0) { _response = "Ok"; }
                            if (_XDetectedQuanPeak.IsResponseHigh > 0) { _response = "High"; }
                            if (_XDetectedQuanPeak.IsResponseLow > 0) { _response = "Low"; }
                            _wsCurrent.Cells[_currRowN, 9] = _response;
                            _wsCurrent.Cells[_currRowN, 12] = _XDetectedQuanPeak.ApexRT;

                            _strRT = _XDetectedQuanPeak.ApexRT.ToString("F2");
                        }
                        else
                        {
                            _wsCurrent.Cells[_currRowN, 9] = "Not Found";

                        }



                        string _sample_ID = "N/A";
                        double _specAmount = 0;
                        switch (_Sample.SampleType)
                        {
                            case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                                _sample_ID = _Sample.Level;
                                XCALIBURFILESLib.XCalibrationCurveData _XCalibrationCurveData =
                                (XCALIBURFILESLib.XCalibrationCurveData)_XQuanComponent.CalibrationCurveData;
                                XCALIBURFILESLib.XReplicateRows _XReplicateRows =
                                (XCALIBURFILESLib.XReplicateRows)_XCalibrationCurveData.ReplicateRows;

                                foreach (XCALIBURFILESLib.XReplicateRow _XReplicateRow in _XReplicateRows)
                                {
                                    if (_XReplicateRow.LevelName == _Sample.Level)
                                    {
                                        _specAmount = _XReplicateRow.Amount;
                                        _strSpecAmount = "'" + _XReplicateRow.Amount.ToString("F5") + "'";
                                    }
                                }

                                break;
                            case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                                _sample_ID = _Sample.SampleId;
                                _specAmount = -1;
                                _strSpecAmount = "NULL";
                                break;
                            default:
                                break;
                        }


                        bool _IsSampleIDDefinded = false;
                        int _intTP = -1;

                        if (_sample_ID == "")
                        {
                            _sample_ID = _Sample.FileName.Substring(_Sample.FileName.LastIndexOf('_') + 1);
                        };

                        try
                        {
                            _intTP = Convert.ToInt16(_sample_ID);
                        }
                        catch (System.FormatException _FoEx)
                        {
                            MessageBox.Show(_FoEx.Message);
                        };

                        if (_intTP < 0)
                        {
                            _sample_ID = "";
                            _IsSampleIDDefinded = false;
                        }
                        else
                        {
                            _sample_ID = _intTP.ToString();
                            _IsSampleIDDefinded = true;
                        }


                        if (_specAmount != -1)
                        {
                            _wsCurrent.Cells[_currRowN, 7] = _specAmount;
                        }
                        else
                        {
                            _wsCurrent.Cells[_currRowN, 7] = "";
                        }
                        _wsCurrent.Cells[_currRowN, 10] = _sample_ID;



                        foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponentISTD in _XQuanSelectedComponents)
                        {
                            XCALIBURFILESLib.XComponent _XISTD =
                            (XCALIBURFILESLib.XComponent)_XQuanComponentISTD.Component;

                            if (_XISTD.Name == _XQuanCalibration.ISTD)
                            {
                                if (_XQuanComponentISTD.DoesDetectedQuanPeakExist > 0)
                                {
                                    XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeakISTD =
                                        (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponentISTD.DetectedQuanPeak;
                                    _wsCurrent.Cells[_currRowN, 13] = _XDetectedQuanPeakISTD.ApexRT;

                                    _wsCurrent.Cells[_currRowN, 5] = _XDetectedQuanPeakISTD.Area;

                                    _strISTD_RT = _XDetectedQuanPeakISTD.ApexRT.ToString("F2");
                                }
                                else
                                {
                                    _wsCurrent.Cells[_currRowN, 13] = "N/F";

                                }

                            }

                        }

                        // if (GetSlopeOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType)) == -1)
                        // {
                        //     return;
                        // }

                        string _subjectNum = "";

                        if (checkBox_if_multi_subject.Checked && (int)_Sample.SampleType == 0)
                        {
                            _subjectNum = _Sample.FileName.Split('_')[1];
                        }
                        else
                        {
                            if (ExperimentInfo._Subject.Contains("-"))
                            {
                                _subjectNum = ExperimentInfo._Subject.Split('-')[1];
                            }
                            else _subjectNum = ExperimentInfo._Subject;

                        }


                        if (_IsPeakFoundANALYTE && _IsPeakFoundISTD && _IsSampleIDDefinded && checkBox_ExportIntoDB.Checked)
                        {
                            
                            _strExportSQL.Add("INSERT INTO TMP_TSQ_IMPORT ( ANYLYTE_NAME, FILENAME_DESC, SAMPLE_TYPE_ID, INTEG_TYPE_ID, AREA, ISTD_AREA, AREA_RATIO, " +
                                                                          "SPEC_AMOUNT, CALC_AMOUNT, LEVEL_TP, PEAK_SATUS, RT, ISTD_RT, QUAN_TYPE_ID, FLUID_TYPE_ID, " +
                                                                          "SUBJECT_NUM, PROCESS_BY_ID, ASSAY_DATE, ASSAY_TYPE_ID, ANTIBODY_ID, ENZYME_ID, DONE_BY, " +
                                                                          "SAMPLE_PROCESS_BY, EXPER_FILE, CURVE, WEIGHTING, ORIGIN, EQUATION, R, R_SQR, NUM_SLOPE, " +
                                                                          "NUM_INTERSECT )");
                            _strExportSQL.Add(String.Format("VALUES ('{0}', '{1}', {2}, '{3}', {4:F0}, {5:F0}, {6:F5}, {7:F}, {8:F5}, {9}, '{10}', {11}, {12}, {13}, {14}, '{15}', {16}, '{17}', {18}, {19}, {20}, {21}, {22}, '{23}'," +
                                " '{24}', '{25}', '{26}', '{27}', {28:F5}, {29:F5}, {30:F5}, {31:F5});",
                                   _strAnalyte,                                 // 0
                                   _Sample.FileName,                            // 1
                                   (int)_Sample.SampleType,                     // 2
                                   (int)_XQuanComponent.Selected_Type,          // 3
                                   _Area,                                       // 4
                                   _ISTD_Area,                                  // 5
                                   _AreaRatio,                                  // 6
                                   _strSpecAmount,                              // 7 
                                   _CalcAmount,                                 // 8
                                   _sample_ID,                                  // 9
                                   _response,                                   // 10
                                   _strRT,                                      // 11 
                                   _strISTD_RT,                                 // 12
                                   QuanTypeToDBInt(ExperimentInfo._quanType),
                                   FluidTypeToDBInt(ExperimentInfo._matrix),    // 14
                                   _subjectNum,       // 15
                                   ExperimentInfo._SampleQuantitatedBy_ID,      // 16
                                   ExperimentInfo._Date,                        // 17
                                   2,                                           // 18      { ASSAY_TYPE_ID   2 - "IP LC/MS (C-term)"    
                                   IPToDBInt(ExperimentInfo._IP),               // 19      { 3, // A-BODY  3 - HJ5.x
                                   EnzymeToDBInt(ExperimentInfo._Enzyme),       // 20      { 2, //Enzyme 2 - LysN
                                   ExperimentInfo._SampleRanBy_ID,              // 21      { Ran by 1 - //ovodvo
                                   ExperimentInfo._SamplePreparedBy_ID,         // 22      { Sample prepared by //1 - ovodvo
                                   ExperimentInfo._expName,                     // 23

                                   CurveTypeToStr(_XQuanMethodCalibration.CurveType),
                                   weightingTypeToStr(_XQuanMethodCalibration.Weighting),
                                   OriginTypeToStr(_XQuanMethodCalibration.OriginType),
                                   GetEquaOfEquation(_XQuanMethodCalibration.FullEquation),
                                   GetROfEquation(_XQuanMethodCalibration.FullEquation),
                                   GetR_SQROfEquation(_XQuanMethodCalibration.FullEquation),
                                   GetSlopeOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType)),
                                   GetInterOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType))


                                   ));

                        }
                    }
                }

                toolStripProgressBarMain.PerformStep();
                toolStripStatusLabelMain.Text = "Processing ..." + _Sample.FileName;
                toolStripStatusLabelMain.PerformClick();

            }

            //_xlApp.Visible = true;
            toolStripProgressBarMain.Visible = false;
            toolStripProgressBarMain.Value = 0;
            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();
            toolStripProgressBarMain.Visible = true;

            toolStripProgressBarMain.Maximum = xlWB.Worksheets.Count - 1;


            string _AppPath = Application.StartupPath + Path.DirectorySeparatorChar;

            string[] _sql_P1 = File.ReadAllLines(_AppPath + "sql_TSQImport_P1.sql");
            string[] _sql_P2 = File.ReadAllLines(_AppPath + "sql_TSQImport_P2.sql");
            string[] _sql_P3 = File.ReadAllLines(_AppPath + "sql_TSQImport_P3.sql");

            _strExportSQL.Insert(0, " ");

            _strExportSQL.InsertRange(0, _sql_P1.ToList<string>());


            _strExportSQL.Insert(0, "/*"); //Assembly.GetExecutingAssembly().

            _strExportSQL.Insert(1, String.Format("Generated by XQN Reader {0} from {1}", Assembly.GetExecutingAssembly().GetName().Version.ToString(), Convert.ToDateTime(Properties.Resources.BuildDate).ToShortDateString()));

            _strExportSQL.Insert(2, String.Format("       Machine: {0} ran by {1}@{2}", Environment.MachineName, Environment.UserName, Environment.UserDomainName));
            _strExportSQL.Insert(3, String.Format("     Timestamp: {0} @ {1}", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString()));
            _strExportSQL.Insert(4, String.Format("        Source: {0} ", textBox_XQNfile.Text));

            _strExportSQL.Insert(5, "*/");


            _strExportSQL.Add(" ");

            _strExportSQL.AddRange(_sql_P3.ToList<string>());


            File.WriteAllLines(Path.ChangeExtension(textBox_Excelfile.Text, "sql"), _strExportSQL.ToArray());

            foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
            {
                _wsComponent.Name = "_" + _wsComponent.Name;

                toolStripStatusLabelMain.Text = "Formatting ..." + _wsComponent.Name;
                toolStripStatusLabelMain.PerformClick();

                if (_wsComponent.Name.Contains("Summary")) { continue; }

                Excel.ChartObjects _xlCharts = (Excel.ChartObjects)_wsComponent.ChartObjects(Type.Missing);
                string _componentName = _wsComponent.get_Range("A2", "A2").Value2.ToString();

                Excel.Range _xlRange_Compound,
                            _xlRange_StdBracket_Area,
                            _xlRange_StdBracket_ISArea,
                            _xlRange_StdBracket_SpecifiedAmount,
                            _xlRange_StdBracket_Ratio,
                            _xlRange_Unknown_Area,
                            _xlRange_Unknown_ISArea,
                            _xlRange_Unknown_CalcAmount,
                            _xlRange_Unknown_SampleID,
                            _xlRange_STD,
                            _xlRange_Unknowns,
                            _xlRange_All;

                _xlRange_Compound = _wsComponent.get_Range("A2", "A2");

                //All
                _xlRange_All =
                    _wsComponent.get_Range
                    ("A" + SkipFirstN.ToString(), "M" + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_All.Name = _wsComponent.Name + "_All";


                //Std
                if (_StdBracketCount > 0)
                {
                    _xlRange_STD =
                        _wsComponent.get_Range
                        ("A" + SkipFirstN.ToString(), "M" + (SkipFirstN - 1 + _StdBracketCount).ToString());
                    _xlRange_STD.Name = _wsComponent.Name + "_STD";


                    _xlRange_StdBracket_Ratio =
                        _wsComponent.get_Range
                        ("F" + SkipFirstN.ToString(), "F" + (SkipFirstN - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_Ratio.Name = _wsComponent.Name + "_StdBracket_Area";

                    _xlRange_StdBracket_Area =
                        _wsComponent.get_Range
                        ("D" + SkipFirstN.ToString(), "D" + (SkipFirstN - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_Area.Name = _wsComponent.Name + "_StdBracket_Area";

                    _xlRange_StdBracket_ISArea =
                        _wsComponent.get_Range
                        ("E" + SkipFirstN.ToString(), "E" + (SkipFirstN - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_ISArea.Name = _wsComponent.Name + "_StdBracket_ISArea";

                    _xlRange_StdBracket_SpecifiedAmount =
                       _wsComponent.get_Range
                       ("G" + SkipFirstN.ToString(), "G" + (SkipFirstN - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_SpecifiedAmount.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount";


                    _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                    _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                    _xlRange_STD.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);

                    //_xlRange_StdBracket_ISArea.FormatConditions.AddDatabar();


                    //*********Sort Range************
                  /*  _xlRange_STD.Sort("Column J",
                        Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                        Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                        Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                        Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                        Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                        Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                        Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                        Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                        Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal); */




                    //****************
                    // StdCurve;
                    Excel.ChartObject _xlChartStdCurve = (Excel.ChartObject)_xlCharts.Add(
                                       490, 800, 470, 250);
                    Excel.Chart _xlChartStdCurvePage = _xlChartStdCurve.Chart;
                    _xlChartStdCurvePage.ChartType = Excel.XlChartType.xlXYScatter;

                    _xlChartStdCurvePage.HasTitle = true;
                    _xlChartStdCurvePage.ChartTitle.Font.Size = 11;

                    string _chartTitleStdCurve = String.Format
                        ("Calibration curve {0} {1} {2} {3} ({4})",
                        _componentName,
                        ExperimentInfo._IP,
                        ExperimentInfo._Enzyme,
                        ExperimentInfo._Instument,
                        ExperimentInfo._Date);
                    //_xlSeries_TTR.Name = _componentName;

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
                    _xlSerie_Stds.XValues = _xlRange_StdBracket_SpecifiedAmount.Cells;
                    _xlSerie_Stds.Values = _xlRange_StdBracket_Ratio.Cells;


                    string _serieStdName = "H:L Abeta";

                    string _axeStdXNumFormat = "0.00",
                           _axeStdYNumFormat = "0.00";

                    string _axeStdXCaption = "%, Predicted labeling Abeta",
                           _axeStdYCaption = "%, Measured labeling Abeta";

                    if (ExperimentInfo._quanType == "ABS")
                    {
                        _axeStdXNumFormat = "0.0";
                        _axeStdXNumFormat = "0.00";

                        _axeStdXCaption = "ng/mL, Specified amount";
                        _axeStdYCaption = "AreaRatio";

                        _serieStdName = "Analyte:ISTD Ratio";
                    }

                    _xlAxisX_Std.TickLabels.NumberFormat = _axeStdXNumFormat;
                    _xlAxisY_Std.TickLabels.NumberFormat = _axeStdYNumFormat;
                    _xlAxisX_Std.AxisTitle.Caption = _axeStdXCaption;
                    _xlAxisY_Std.AxisTitle.Caption = _axeStdYCaption;

                    _xlSerie_Stds.Name = _serieStdName;

                    Excel.Trendlines _xlTrendlines_Stds = (Excel.Trendlines)_xlSerie_Stds.Trendlines(Type.Missing);

                    Excel.XlTrendlineType _XlTrendlineType =
                        Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear;
                    string _CurveIndex = _wsComponent.get_Range("C2", "C2").Value2.ToString();

                    if (_CurveIndex == "Quadratic")
                    { _XlTrendlineType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlExponential; }


                    _xlTrendlines_Stds.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, true, true, Type.Missing);

                    //****************

                };


                //Unknowns
                _xlRange_Unknowns =
                    _wsComponent.get_Range
                    ("A" + (SkipFirstN + _StdBracketCount).ToString(), "M" + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknowns.Name = _wsComponent.Name + "_Unknowns";

                _xlRange_Unknown_Area =
                    _wsComponent.get_Range
                    ("D" + (SkipFirstN + _StdBracketCount).ToString(), "D" + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_Area.Name = _wsComponent.Name + "_Unknown_Area";

                _xlRange_Unknown_ISArea =
                    _wsComponent.get_Range
                    ("E" + (SkipFirstN + _StdBracketCount).ToString(), "E" + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_ISArea.Name = _wsComponent.Name + "_Unknown_ISArea";

                string strCalcOrRAtio = "H";
                if (_StdBracketCount == 0) { strCalcOrRAtio = "F"; }
                _xlRange_Unknown_CalcAmount =
                    _wsComponent.get_Range
                    (strCalcOrRAtio + (SkipFirstN + _StdBracketCount).ToString(), strCalcOrRAtio + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_CalcAmount.Name = _wsComponent.Name + "_Unknown_CalcAmount";

                _xlRange_Unknown_SampleID =
                    _wsComponent.get_Range
                    ("J" + (SkipFirstN + _StdBracketCount).ToString(), "J" + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_SampleID.Name = _wsComponent.Name + "_Unknown_SampleID";







                //*************************Format Ranges**************************
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;

                //******Sort Unknowns
                /*
                _xlRange_Unknowns.Sort("Column J",
                    Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                    Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal); */

                //*************2011-07-19**************************************

                //_xlRange_Unknown_ISArea.FormatConditions.AddDatabar();

                Excel.ColorScale cfColorScale = (Excel.ColorScale)(_xlRange_Unknown_ISArea.FormatConditions.AddColorScale(3));

                // Set the minimum threshold to red (0x000000FF) and maximum threshold
                // to blue (0x00FF0000).
                cfColorScale.ColorScaleCriteria[1].FormatColor.Color = 7039480;  //0x000000FF;  
                cfColorScale.ColorScaleCriteria[2].FormatColor.Color = 8711167;  // 0x00FF0000;
                cfColorScale.ColorScaleCriteria[3].FormatColor.Color = 13011546; // 0x00FF0000;

                //*************************Charts*********************************
                //

                string cellAddress = "A" + (ExperimentInfo.TitleHeight + ExperimentInfo._totSampleCnt + 2).ToString();

                double plotOffset = (double)_wsCurrent.get_Range(cellAddress, cellAddress).Top;
                double plotWidth = 470; if (ExperimentInfo._StdBracketCount == 0) { plotWidth = (double)_xlRange_All.Width; }

                Excel.ChartObject _xlChartTimeCorse = (Excel.ChartObject)_xlCharts.Add(
                                   0, plotOffset, plotWidth, 250);
                Excel.Chart _xlChartTimeCorsePage = _xlChartTimeCorse.Chart;
                _xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlXYScatter;
                _xlChartTimeCorsePage.HasTitle = true; _xlChartTimeCorsePage.ChartTitle.Font.Size = 11;

                _xlChartTimeCorsePage.ChartTitle.Caption = _componentName + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_TP = (Excel.Axes)_xlChartTimeCorsePage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_TP.HasMajorGridlines = true; _xlAxisX_TP.HasTitle = true; _xlAxisY_TP.HasTitle = true;

                _xlChartTimeCorsePage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartTimeCorsePage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_TimeCourse = (Excel.SeriesCollection)_xlChartTimeCorsePage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_TTR = (Excel.Series)_xlSeriesColl_TimeCourse.NewSeries();
                _xlSeries_TTR.XValues = _xlRange_Unknown_SampleID.Cells; _xlSeries_TTR.Values = _xlRange_Unknown_CalcAmount;


                string _serieCurveName = "H:L Abeta";

                string _axeCurveXNumFormat = "0",
                       _axeCurveYNumFormat = "0.0%";

                string _axeCurveXCaption = "h, Time",
                       _axeCurveYCaption = "H:L ratio";

                if (ExperimentInfo._quanType == "ABSOLUTE")
                {
                    _axeCurveYNumFormat = "0.0";
                    _axeCurveYCaption = "ng/mL, Concentration";
                    _serieCurveName = "Abeta level";
                }

                if (_StdBracketCount == 0) { _axeCurveYCaption = "Area Ratio"; }

                _xlAxisX_TP.TickLabels.NumberFormat = _axeCurveXNumFormat; _xlAxisY_TP.TickLabels.NumberFormat = _axeCurveYNumFormat;
                _xlAxisX_TP.AxisTitle.Caption = _axeCurveXCaption; _xlAxisY_TP.AxisTitle.Caption = _axeCurveYCaption;

                _xlSeries_TTR.Name = _serieCurveName; _xlSeries_TTR.Name = _componentName;







                //**********************************************************************************************************************************************
                Excel.ChartObjects _xlChartsSummary = (Excel.ChartObjects)_wsSummary.ChartObjects(Type.Missing);
                Excel.ChartObject _xlChartCurveSummary;
                Excel.Chart _xlChartCurveSummaryPage;

                if (_xlChartsSummary.Count < 1)
                {
                    _xlChartCurveSummary = (Excel.ChartObject)_xlChartsSummary.Add(
                                   10, 10, 850, 500);
                    _xlChartCurveSummaryPage = _xlChartCurveSummary.Chart;
                    _xlChartCurveSummaryPage.ChartType = Excel.XlChartType.xlXYScatter;
                    _xlChartCurveSummaryPage.HasTitle = true;
                    _xlChartCurveSummaryPage.ChartTitle.Font.Size = 11;



                    Excel.Axes _xlAxes_CurveSummary = (Excel.Axes)_xlChartCurveSummaryPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisX_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisY_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                    _xlAxisX_CurveSummary.HasMajorGridlines = true; _xlAxisY_CurveSummary.HasMajorGridlines = true;

                    _xlAxisX_CurveSummary.Format.Line.Weight = 3.0F; _xlAxisY_CurveSummary.Format.Line.Weight = 3.0F;

                    _xlAxisX_CurveSummary.HasTitle = true; _xlAxisY_CurveSummary.HasTitle = true;

                    _xlAxisX_CurveSummary.AxisTitle.Caption = _axeCurveXCaption; _xlAxisY_CurveSummary.AxisTitle.Caption = _axeCurveYCaption;

                    _xlAxisX_CurveSummary.TickLabels.NumberFormat = _axeCurveXNumFormat; _xlAxisY_CurveSummary.TickLabels.NumberFormat = _axeCurveYNumFormat;

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
                _xlSeries_CurveTTR.XValues = _xlRange_Unknown_SampleID.Cells;
                _xlSeries_CurveTTR.Values = _xlRange_Unknown_CalcAmount;
                _xlSeries_CurveTTR.Name = (string)_xlRange_Compound.Cells.Value2;


                toolStripProgressBarMain.PerformStep();

            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = false;

            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();

            _XXQN.Close();

            object misValue = System.Reflection.Missing.Value;

            try
            {
                _XXQN.Close();

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
                    ((Excel._Worksheet)_wsSummary).Activate();

                }
                else
                {
                    _xlApp.Quit();
                }
            }
            if (this.checkBox_ShowDebug.Checked)
            {
                MessageBox.Show("ExportXcaliburUni()");
            }
        }


        //******************************************************* 2014-02-26 for ABS ApoE CSF-curve + Syhtn
        private void ExportAPOE_ABS_old()
        {

            List<string> _strExportSQL = new List<string>();
            /*05-07-2010*/
            string _strAnalyte = null;
            double _Area = 0, _ISTD_Area = 0, _Area_Ratio = 0, _CalcAmount = 0;
            decimal dcm_Area_Ratio = 0;
            string _strSpecAmount = null, _strRT = null, _strISTD_RT = null;
            string _response = null;

            /*bool IsPeakFound = false;*/

            /*05-07-2010*/
            if (textBox_XQNfile.Text == "")
            {
                MessageBox.Show("Select Input XQN file, please");
                return;
            }


            _XXQN = new XCALIBURFILESLib.XXQNClass();

            try
            {
                _XXQN.Open(textBox_XQNfile.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("File: " + textBox_XQNfile.Text + " is invalid", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //MessageBox.Show(_exception.Message, _exception.GetBaseException().Message);
                return;
            }
            finally
            {

            }

            XCALIBURFILESLib.IXHeader _IXHeader = (XCALIBURFILESLib.IXHeader)_XXQN.Header;

            ExperimentInfo._ProcessedBy = _IXHeader.ChangedId;

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = _IXHeader.ChangedLogon;
            }

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = "unknown";
            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = true;
            toolStripStatusLabelMain.Visible = true;



            _xlApp = new Excel.ApplicationClass();
            _xlApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;



            Excel.Workbook xlWB;

            try
            {

                xlWB =
                    _xlApp.Workbooks.Add(
                    //@"c:\Project\XQNReader\XQNReader\LOAD_TSQ_Template_Summary.xlsx"
                Application.StartupPath + Path.DirectorySeparatorChar +
                @"\Templates\ApoE_TSQ_ABS_Template_Summary.xlsx"
                );

            }

            catch (Exception)
            {
                MessageBox.Show("File: " + Application.StartupPath + Path.DirectorySeparatorChar +
                @"\Templates\ApoE_TSQ_ABS_Template_Summary.xlsx", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //MessageBox.Show(_exception.Message, _exception.GetBaseException().Message);
                return;
            }
            finally
            {

            }

            Excel.Worksheet _wsSummary = (Excel.Worksheet)xlWB.ActiveSheet;

            Excel.Worksheet _wsCurrent = null;


            byte Title = 5;
            byte _UnknownCount = 0, _StdBracketCount = 0, _BlankCount = 0, _TotalCount = 0;




            XCALIBURFILESLib.XBracketGroups _XBracketGroups
                = (XCALIBURFILESLib.XBracketGroups)_XXQN.BracketGroups;

            XCALIBURFILESLib.XBracketGroup _XBracketGroup
                = (XCALIBURFILESLib.XBracketGroup)_XBracketGroups.Item(1);

            XCALIBURFILESLib.XQuanResults _XQuanResults
                = (XCALIBURFILESLib.XQuanResults)_XBracketGroup.QuanResults;


            var _XQuanResultsSorted = from XCALIBURFILESLib.XQuanResult t in _XQuanResults
                                      //orderby t.Sample
                                      select t;


            //********_XQuanResults*******************************

            foreach (XCALIBURFILESLib.XQuanResult _XQuanResult in _XQuanResultsSorted /*_XQuanResults*/)
            {
                toolStripProgressBarMain.Maximum = _XQuanResults.Count;

                XCALIBURFILESLib.XSample _Sample
                    = (XCALIBURFILESLib.XSample)_XQuanResult.Sample;


                //_Sample.




                _TotalCount++;
                switch (_Sample.SampleType)
                {
                    case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                        _UnknownCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                        _StdBracketCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleBlank:
                        _BlankCount++;
                        break;
                    default:
                        break;
                }

                XCALIBURFILESLib.XQuanComponents _XQuanSelectedComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Selected; /*  Method; */
                //  = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.User;  
                ////////////////////////////////////////
                XCALIBURFILESLib.XQuanComponents _XQuanMethodComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Method; /*  Method; */


                ////////////////////////////////////////
                short _comNo = 1;
                foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponent in _XQuanSelectedComponents)
                {
                    XCALIBURFILESLib.XComponent _XComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanComponent.Component;

                    XCALIBURFILESLib.XQuanComponent _XQuanMethodComponent = (XCALIBURFILESLib.XQuanComponent)
                    _XQuanMethodComponents.Item(_comNo);
                    _comNo++;

                    XCALIBURFILESLib.XComponent _XMethodComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanMethodComponent.Component;

                    /*
                    XCALIBURFILESLib.XComponent _XComponentUser =
                        (XCALIBURFILESLib.XComponent)_XQuanComponent.Component;
                    */

                    //****************** 
                    XCALIBURFILESLib.IXDetection2 _IXDetection =
                        (XCALIBURFILESLib.IXDetection2)_XComponent.Detection;
                    //*****************

                    XCALIBURFILESLib.XQuanCalibration _XQuanCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XComponent.Calibration;

                    XCALIBURFILESLib.XQuanCalibration _XQuanMethodCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XMethodComponent.Calibration;

                    string _quanTypeKey = "";

                    // MessageBox.Show(ExperimentInfo._quanType.ToString());

                    if (ExperimentInfo._quanType == "RELATIVE")
                    {
                        _quanTypeKey = "R0:N14";
                    }
                    if (ExperimentInfo._quanType == "ABSOLUTE")
                    {
                        _quanTypeKey = "L6:C13";
                    }

                    //MessageBox.Show(_XQuanCalibration.ComponentType.ToString());
                    //MessageBox.Show(_XComponent.Name.Contains(_quanTypeKey).ToString());

                    if (_XQuanCalibration.ComponentType != XCALIBURFILESLib.XComponentType.XISTD
                        && !_XComponent.Name.Contains(_quanTypeKey))
                    {

                        // MessageBox.Show(_quanTypeKey);

                        _strAnalyte = _XComponent.Name.Replace("_", " "); //*******************

                        // ****** 2010-07-23
                        _strAnalyte = _strAnalyte.Replace("AB-Total", "AbetaTotal");
                        // ****** 2010-07-23


                        bool _SheetExist = false;
                        string _ComponentName = _XComponent.Name.Replace(":", "_");
                        _ComponentName = _ComponentName.Replace("-", "_");
                        //**********
                        _ComponentName = _ComponentName.Replace("(", "_");
                        _ComponentName = _ComponentName.Replace(")", "_");
                        //**********

                        foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
                        {

                            if (_wsComponent.Name == _ComponentName
                                )
                            {
                                _wsCurrent = _wsComponent;
                                _SheetExist = true;
                            }
                        }



                        if (!_SheetExist)
                        {
                            _wsCurrent = (Excel.Worksheet)xlWB.Sheets.Add
                          (Type.Missing, Type.Missing, Type.Missing,
                          Application.StartupPath + Path.DirectorySeparatorChar +
                            @"\Templates\ApoE_TSQ_ABS_Template.xlsx"
                          );

                            _wsCurrent.Name = _ComponentName;



                            _wsCurrent.Cells[2, 1] = _XComponent.Name.Replace("_", " ");

                            /*  _wsCurrent.Cells[2, 3] =
                                  CurveTypeToStr(_XQuanMethodCalibration.CurveType);
                              _wsCurrent.Cells[2, 4] =
                                  weightingTypeToStr(_XQuanMethodCalibration.Weighting);
                              _wsCurrent.Cells[2, 6] =
                                  OriginTypeToStr(_XQuanMethodCalibration.OriginType);

                              _wsCurrent.Cells[2, 8] = _XQuanMethodCalibration.FullEquation; */

                        }

                        string _subjectNum = "";

                        if (checkBox_if_multi_subject.Checked && (int)_Sample.SampleType == 0)
                        {
                            _subjectNum = _Sample.FileName.Split('_')[1].Replace("CSF", "");
                        }
                        else
                        {
                            if (ExperimentInfo._Subject.Contains("-"))
                            {
                                _subjectNum = ExperimentInfo._Subject.Split('-')[1];
                            }
                            else _subjectNum = ExperimentInfo._Subject;

                        }

                        _wsCurrent.Cells[_Sample.RowNumber + Title, 1] = _Sample.FileName;
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 2] = SampleTypeToStr(_Sample.SampleType);
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 3] = IntegTypeToStr(_XQuanComponent.Selected_Type);


                        XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeak;


                        //************************************
                        /*
                        if (_Sample.RowNumber > 23 && _ComponentName.Contains("42") )
                        {
                            MessageBox.Show(_IXDetection.PeakWidthHeight.ToString()
                                + "  " + _IXDetection.SNThreshold.ToString());
                        }
                        */
                        //************************************

                        bool _IsPeakFoundANALYTE = false;
                        bool _IsPeakFoundISTD = false;
                        int _shift = 2;

                        if (_XQuanComponent.DoesDetectedQuanPeakExist > 0)
                        {
                            _IsPeakFoundANALYTE = true;
                            _XDetectedQuanPeak =
                                (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponent.DetectedQuanPeak;

                            _wsCurrent.Cells[_Sample.RowNumber + Title, 4] = _XDetectedQuanPeak.Area;
                            _Area = _XDetectedQuanPeak.Area;

                            //_XDetectedQuanPeak.ISTDValid


                            _wsCurrent.Cells[_Sample.RowNumber + Title, 8] = _XDetectedQuanPeak.CalcAmount;

                            _CalcAmount = _XDetectedQuanPeak.CalcAmount;


                            /////
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 20 + _shift] = _XDetectedQuanPeak.LeftRT;
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 21 + _shift] = _XDetectedQuanPeak.RightRT;


                            _wsCurrent.Cells[_Sample.RowNumber + Title, 19 + _shift] = String.Format("=L{0}-M{0}", (_Sample.RowNumber + Title).ToString());
                            // _wsCurrent.Cells[_Sample.RowNumber + Title, 22 + _shift] = String.Format("=U{0}-T{0}", (_Sample.RowNumber + Title).ToString());
                            // _wsCurrent.Cells[_Sample.RowNumber + Title, 25 + _shift] = String.Format("=X{0}-W{0}", (_Sample.RowNumber + Title).ToString());


                            /*
                            XCALIBURFILESLib.IXSpectrum  t1;
 
                            
                            XCALIBURFILESLib.IXComponent t2;

                            t2 =   (XCALIBURFILESLib.IXComponent) _XQuanComponent.Component;

                             XCALIBURFILESLib.IXDetection2 t3;
                            
                            t3=                 (XCALIBURFILESLib.IXDetection2)        t2.Detection;

                             XCALIBURFILESLib.IXFindParams t4;

                            t4 =        (XCALIBURFILESLib.IXFindParams)           t3.FindParams;

                            XCALIBURFILESLib.IXSpectrum4 t5;
                            t5 =
                             (XCALIBURFILESLib.IXSpectrum4)  t4.Spectrum;


                            //XCALIBURFILESLib.IXTraceSelector  t7;

                            XCALIBURFILESLib.IXParentScans  t8;

                            //t8.Count

                            
                           // t8 = (XCALIBURFILESLib.IXParentScans) t5.ParentScans;

                            
                            

                            //System.Array t6;

                            System.Double[,] t6;

                            //t6 =   (double[,]) t5.get_DataInRange(1);

                            t6 =   (System.Double[,])t5.LabelData;

                           // object t9 = t5.Data.;

                           // MessageBox.Show(t5.LabelData.GetType().ToString() );
                           
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 30] =   t6.Length;

                           // _XQuanComponent.Component

                          //  _wsCurrent.Cells[_Sample.RowNumber + Title, 22] = _XDetectedQuanPeak.get_IonRatioQualifierIon(1). .g .LeftEdge;


                           // _wsCurrent.Cells[_Sample.RowNumber + Title, 23] = _XDetectedQuanPeak.RightEdge;

                           // _wsCurrent.Cells[_Sample.RowNumber + Title, 24] = _XDetectedQuanPeak.LeftBase;
                          //  _wsCurrent.Cells[_Sample.RowNumber + Title, 25] = _XDetectedQuanPeak.RightBase;


                            */
                            /////


                            if (_XDetectedQuanPeak.ISTDArea != 0)
                            {
                                _IsPeakFoundISTD = true;
                                _wsCurrent.Cells[_Sample.RowNumber + Title, 5] = _XDetectedQuanPeak.ISTDArea;
                                _wsCurrent.Cells[_Sample.RowNumber + Title, 6] =
                                _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;



                                _ISTD_Area = _XDetectedQuanPeak.ISTDArea;
                                _Area_Ratio = _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;
                                dcm_Area_Ratio = (decimal)_XDetectedQuanPeak.Area / (decimal)_XDetectedQuanPeak.ISTDArea;
                            }
                            else
                            {
                                _wsCurrent.Cells[_Sample.RowNumber + Title, 5] = "N/F";
                                //_wsCurrent.Cells[_Sample.RowNumber + Title, 6] = "-";
                            }



                            //string _response = "";
                            _response = "NULL"; _strRT = "NULL"; _strISTD_RT = "NULL";

                            if (_XDetectedQuanPeak.IsResponseOK > 0) { _response = "Ok"; }
                            if (_XDetectedQuanPeak.IsResponseHigh > 0) { _response = "High"; }
                            if (_XDetectedQuanPeak.IsResponseLow > 0) { _response = "Low"; }
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 9] = _response;
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 12] = _XDetectedQuanPeak.ApexRT;

                            _strRT = _XDetectedQuanPeak.ApexRT.ToString("F2");
                            //_RT = _XDetectedQuanPeak.ApexRT;
                        }
                        else
                        {
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 9] = "Not Found";

                        }



                        string _sample_ID = "N/A";
                        double _specAmount = 0;
                        switch (_Sample.SampleType)
                        {
                            case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                                _sample_ID = _Sample.Level;
                                XCALIBURFILESLib.XCalibrationCurveData _XCalibrationCurveData =
                                (XCALIBURFILESLib.XCalibrationCurveData)_XQuanComponent.CalibrationCurveData;
                                XCALIBURFILESLib.XReplicateRows _XReplicateRows =
                                (XCALIBURFILESLib.XReplicateRows)_XCalibrationCurveData.ReplicateRows;

                                foreach (XCALIBURFILESLib.XReplicateRow _XReplicateRow in _XReplicateRows)
                                {
                                    if (_XReplicateRow.LevelName == _Sample.Level)
                                    {
                                        _specAmount = _XReplicateRow.Amount;
                                        _strSpecAmount = "'" + _XReplicateRow.Amount.ToString("F5") + "'";
                                    }
                                }

                                _wsCurrent.Cells[_Sample.RowNumber + Title, 11] = _Sample.FileName.Split('_')[1];

                                break;

                            case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                                _sample_ID = _Sample.FileName.Split('_')[2].Replace("Hr", "");
                                _specAmount = -1;
                                _strSpecAmount = "NULL";

                                _wsCurrent.Cells[_Sample.RowNumber + Title, 11] = _subjectNum;
                                //   _wsCurrent.Cells[_Sample.RowNumber + Title, 11] = _Sample.FileName.Substring(12, 5);

                                break;

                            case XCALIBURFILESLib.XSampleTypes.XSampleBlank:
                                _sample_ID = _Sample.SampleId;
                                _specAmount = -1;
                                _strSpecAmount = "NULL";
                                break;
                            default:
                                break;
                        }

                        //**********
                        //_Sample.
                        //*********
                        bool _IsSampleIDDefinded = false;
                        int _intTP = -1;

                        if (_sample_ID == "")
                        {
                            _sample_ID = _Sample.FileName.Split('_')[2].Replace("Hr", "");
                        };

                        try
                        {
                            _intTP = Convert.ToInt16(_sample_ID);
                        }
                        catch (System.FormatException _FoEx)
                        {
                            MessageBox.Show(_sample_ID + " || " + _FoEx.Message);
                        };

                        if (_intTP < 0)
                        {
                            _sample_ID = "";
                            _IsSampleIDDefinded = false;
                        }
                        else
                        {
                            _sample_ID = _intTP.ToString();
                            _IsSampleIDDefinded = true;
                        }


                        if (_specAmount != -1)
                        {
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 7] = _specAmount;
                        }
                        else
                        {
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 7] = "";
                        }
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 10] = _sample_ID;




                        foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponentISTD in _XQuanSelectedComponents)
                        {
                            XCALIBURFILESLib.XComponent _XISTD =
                            (XCALIBURFILESLib.XComponent)_XQuanComponentISTD.Component;

                            if (_XISTD.Name == _XQuanCalibration.ISTD)
                            {
                                if (_XQuanComponentISTD.DoesDetectedQuanPeakExist > 0)
                                {
                                    XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeakISTD =
                                        (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponentISTD.DetectedQuanPeak;
                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 13] = _XDetectedQuanPeakISTD.ApexRT;

                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 23 + _shift] = _XDetectedQuanPeakISTD.LeftRT;
                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 24 + _shift] = _XDetectedQuanPeakISTD.RightRT;



                                    _strISTD_RT = _XDetectedQuanPeakISTD.ApexRT.ToString("F2");
                                }
                                else
                                {
                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 13] = "N/F";
                                }

                            }

                        }

                        _wsCurrent.Cells[_Sample.RowNumber + Title, 14] = _Sample.SampleId;
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 15] = _Sample.SampleName;

                        //_wsCurrent.Cells[_Sample.RowNumber + Title, 16] = _Sample.SampleName;





                        // if (GetSlopeOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType)) == -1)
                        // {
                        //     return;
                        // }




                        if (_IsPeakFoundANALYTE && _IsPeakFoundISTD && _IsSampleIDDefinded && checkBox_ExportIntoDB.Checked)
                        {
                            _strExportSQL.Add("INSERT INTO TMP_TSQ_IMPORT ( ANYLYTE_NAME, FILENAME_DESC, SAMPLE_TYPE_ID, INTEG_TYPE_ID, AREA, ISTD_AREA, AREA_RATIO, " +
                                                                          "SPEC_AMOUNT, CALC_AMOUNT, LEVEL_TP, PEAK_SATUS, RT, ISTD_RT, QUAN_TYPE_ID, FLUID_TYPE_ID, " +
                                                                          "SUBJECT_NUM, PROCESS_BY_ID, ASSAY_DATE, ASSAY_TYPE_ID, ANTIBODY_ID, ENZYME_ID, DONE_BY, " +
                                                                          "SAMPLE_PROCESS_BY, EXPER_FILE, CURVE, WEIGHTING, ORIGIN, EQUATION, R, R_SQR, NUM_SLOPE, " +
                                                                          "NUM_INTERSECT )");
                            _strExportSQL.Add(String.Format("VALUES ('{0}', '{1}', {2}, '{3}', {4:F0}, {5:F0}, {6:F5}, {7:F}, {8:F5}, {9}, '{10}', {11}, {12}, {13}, {14}, '{15}', '{16}', '{17}', {18}, {19}, {20}, {21}, {22}, '{23}'," +
                                " '{24}', '{25}', '{26}', '{27}', {28:F5}, {29:F5}, {30:F5}, {31:F5});",
                                   _strAnalyte,                                 // 0
                                   _Sample.FileName,                            // 1
                                   (int)_Sample.SampleType,                     // 2
                                   (int)_XQuanComponent.Selected_Type,          // 3
                                   _Area,                                       // 4
                                   _ISTD_Area,                                  // 5
                                   _Area_Ratio,                                 // 6
                                   _strSpecAmount,                              // 7 
                                   _CalcAmount,                                 // 8  "NULL", //
                                   _sample_ID,                                  // 9
                                   _response,                                   // 10
                                   _strRT,                                      // 11 
                                   _strISTD_RT,                                 // 12
                                   QuanTypeToDBInt(ExperimentInfo._quanType),
                                   FluidTypeToDBInt(ExperimentInfo._matrix),    // 14
                                   _subjectNum,       // 15
                                   ExperimentInfo._ProcessedBy,                 // 16
                                   ExperimentInfo._Date,                        // 17
                                   11,                                           // 18      { ASSAY_TYPE_ID   2 - "IP LC/MS (C-term)"    
                                   comboBox_Abody.SelectedValue,                // 19      { 3, // A-BODY  3 - HJ5.x
                                   comboBox_Enzyme.SelectedValue,       // 20      { 2, //Enzyme 2 - LysN
                                   1,                                           // 21      { Done by 1 - ovodvo
                                   1,                                           // 22      { Sample process by 1 - ovodvo
                                   ExperimentInfo._expName,                     // 23

                                   "NULL", //CurveTypeToStr(_XQuanMethodCalibration.CurveType),
                                   "NULL", //weightingTypeToStr(_XQuanMethodCalibration.Weighting),
                                   "NULL", //OriginTypeToStr(_XQuanMethodCalibration.OriginType),
                                   "NULL", //GetEquaOfEquation(_XQuanMethodCalibration.FullEquation, _strAnalyte),
                                   "NULL", //GetROfEquation(_XQuanMethodCalibration.FullEquation),
                                   "NULL", //GetR_SQROfEquation(_XQuanMethodCalibration.FullEquation),
                                   "NULL", //GetSlopeOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType)),
                                   "NULL" //GetInterOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType))


                                   ));

                        }
                    }
                }

                toolStripProgressBarMain.PerformStep();
                toolStripStatusLabelMain.Text = "Processing ..." + _Sample.FileName;
                toolStripStatusLabelMain.PerformClick();

            }

            //_xlApp.Visible = true;
            toolStripProgressBarMain.Visible = false;
            toolStripProgressBarMain.Value = 0;
            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();
            toolStripProgressBarMain.Visible = true;

            toolStripProgressBarMain.Maximum = xlWB.Worksheets.Count - 1;


            if (checkBox_ExportIntoDB.Checked)
            {

                string _AppPath = Application.StartupPath + Path.DirectorySeparatorChar;

                string[] _sql_P1 = File.ReadAllLines(_AppPath + "sql_TSQImport_P1.sql");
                string[] _sql_P2 = File.ReadAllLines(_AppPath + "sql_TSQImport_P2.sql");
                string[] _sql_P3 = File.ReadAllLines(_AppPath + "sql_TSQImport_P3.sql");

                _strExportSQL.Insert(0, " ");
                _strExportSQL.InsertRange(0, _sql_P1.ToList<string>());
                _strExportSQL.AddRange(_sql_P3.ToList<string>());
                //_strExportSQL.AddRange(_sql_P2.ToList<string>());


                File.WriteAllLines(Path.ChangeExtension(textBox_Excelfile.Text, "sql"), _strExportSQL.ToArray());
            }

            foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
            {
                _wsComponent.Name = "_" + _wsComponent.Name;

                toolStripStatusLabelMain.Text = "Formatting ..." + _wsComponent.Name;
                toolStripStatusLabelMain.PerformClick();

                if (_wsComponent.Name.Contains("Summary")) { continue; }

                Excel.ChartObjects _xlCharts = (Excel.ChartObjects)_wsComponent.ChartObjects(Type.Missing);
                string _componentName = _wsComponent.get_Range("A2", "A2").Value2.ToString();

                Excel.Range _xlRange_Compound,
                            _xlRange_StdBracket_Area,
                            _xlRange_StdBracket_ISArea,
                            _xlRange_StdBracket_SpecifiedAmount,
                            _xlRange_StdBracket_Ratio,
                            _xlRange_Unknown_Area,
                            _xlRange_Unknown_ISArea,
                            _xlRange_Unknown_Ratio,
                            _xlRange_Unknown_CalcAmount,
                            _xlRange_Unknown_SampleID,
                            _xlRange_STD,
                            _xlRange_Unknowns,
                            _xlRange_All,

                           _xlRange_StdBracket_SpecifiedAmount1,
                           _xlRange_StdBracket_SpecifiedAmount2,
                           _xlRange_StdBracket_SpecifiedAmount3,
                           _xlRange_StdBracket_SpecifiedAmount4,
                           
                    
                           _xlRange_StdBracket_Ratio1,
                           _xlRange_StdBracket_Ratio2,
                           _xlRange_StdBracket_Ratio3,
                           _xlRange_StdBracket_Ratio4;
                    

                _xlRange_Compound = _wsComponent.get_Range("A2", "A2");

                //All
                _xlRange_All =
                    _wsComponent.get_Range
                    ("A" + Title.ToString(), "Y" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_All.Name = _wsComponent.Name + "_All";


                double _rowHeight = (double)_wsCurrent.get_Range("A1", "A1").RowHeight;

                double _graphOffset = (double)_xlRange_All.Top + (double)_xlRange_All.Height +

                      (_rowHeight * 3);



                //Std
                // if (_StdBracketCount > 0)
                // {
                _xlRange_STD =
                    _wsComponent.get_Range
                    ("A" + Title.ToString(), "AA" + (Title - 1 + _StdBracketCount).ToString());
                _xlRange_STD.Name = _wsComponent.Name + "_STD";


                _xlRange_StdBracket_Ratio =
                    _wsComponent.get_Range
                    ("F" + Title.ToString(), "F" + (Title - 1 + _StdBracketCount).ToString());
                _xlRange_StdBracket_Ratio.Name = _wsComponent.Name + "_StdBracket_Ratio";

                _xlRange_StdBracket_Area =
                    _wsComponent.get_Range
                    ("D" + Title.ToString(), "D" + (Title - 1 + _StdBracketCount).ToString());
                _xlRange_StdBracket_Area.Name = _wsComponent.Name + "_StdBracket_Area";

                _xlRange_StdBracket_ISArea =
                    _wsComponent.get_Range
                    ("E" + Title.ToString(), "E" + (Title - 1 + _StdBracketCount).ToString());
                _xlRange_StdBracket_ISArea.Name = _wsComponent.Name + "_StdBracket_ISArea";

                _xlRange_StdBracket_SpecifiedAmount =
                   _wsComponent.get_Range
                   ("G" + Title.ToString(), "G" + (Title - 1 + _StdBracketCount).ToString());
                _xlRange_StdBracket_SpecifiedAmount.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount";


                _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                _xlRange_STD.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);

                //Unknowns
                /*   _xlRange_Unknowns = _wsComponent.get_Range
                       ("A" + (Title - 1 + _StdBracketCount).ToString(), "O" + (Title - 1 + _StdBracketCount + _UnknownCount).ToString());
                   _xlRange_STD.Name = _wsComponent.Name + "_Unknown"; */

                // v

                //*********Sort Range************
                _xlRange_STD.Sort(_xlRange_STD.Columns[14, Type.Missing], Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Microsoft.Office.Interop.Excel.XlYesNoGuess.xlGuess,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                    Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);

                /* _xlRange_Unknowns.Sort(_xlRange_Unknowns.Columns[10, Type.Missing],
                                        , Microsoft.Office.Interop.Excel.XlSortOrder.xlDescending,
                                        //_xlRange_Unknowns.Columns[10, Type.Missing], Type.Missing, Excel.XlSortOrder.xlAscending,

                                        Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,

                                         Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                                         Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                                         Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                                         Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                                         Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                                         Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                                         Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);*/


                //****************
                // StdCurve;


                Excel.ChartObject _xlChartStdCurve = (Excel.ChartObject)_xlCharts.Add(
                                   970, _graphOffset, 470, 250);
                Excel.Chart _xlChartStdCurvePage = _xlChartStdCurve.Chart;
                _xlChartStdCurvePage.ChartType = Excel.XlChartType.xlXYScatter;

                _xlChartStdCurvePage.HasTitle = true;
                _xlChartStdCurvePage.ChartTitle.Font.Size = 11;

                string _chartTitleStdCurve = String.Format
                    ("Calibration curve {0} {1} {2} {3} ({4})",
                    _componentName,
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

                Excel.Series _xlSerie_Stds1 = (Excel.Series)_xlSerColl_StdCurve.NewSeries();
                Excel.Series _xlSerie_Stds2 = (Excel.Series)_xlSerColl_StdCurve.NewSeries();
                Excel.Series _xlSerie_Stds3 = (Excel.Series)_xlSerColl_StdCurve.NewSeries();

                Excel.Series _xlSerie_Stds4 = (Excel.Series)_xlSerColl_StdCurve.NewSeries();
                // Excel.Series _xlSerie_Stds5 = (Excel.Series)_xlSerColl_StdCurve.NewSeries();
                // Excel.Series _xlSerie_Stds6 = (Excel.Series)_xlSerColl_StdCurve.NewSeries();


                //String Std1 = 

                _xlRange_StdBracket_SpecifiedAmount1 = _xlRange_StdBracket_SpecifiedAmount.get_Range("A1", "A5");
                //_xlRange_StdBracket_SpecifiedAmount1 = _xlRange_StdBracket_SpecifiedAmount.get_Range("A1", "A6");
                _xlRange_StdBracket_SpecifiedAmount1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);

                _xlRange_StdBracket_SpecifiedAmount2 = _xlRange_StdBracket_SpecifiedAmount.get_Range("A7", "A10");
                _xlRange_StdBracket_SpecifiedAmount2.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);


                _xlRange_StdBracket_SpecifiedAmount3 = _xlRange_StdBracket_SpecifiedAmount.get_Range("A11", "A16");
                _xlRange_StdBracket_SpecifiedAmount3.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);


                _xlRange_StdBracket_SpecifiedAmount4 = _xlRange_StdBracket_SpecifiedAmount.get_Range("A17", "A23");

                //_xlRange_StdBracket_SpecifiedAmount5 = _xlRange_StdBracket_SpecifiedAmount.get_Range("A25", "A30");
                //_xlRange_StdBracket_SpecifiedAmount6 = _xlRange_StdBracket_SpecifiedAmount.get_Range("A31", "A36");

                _xlRange_StdBracket_Ratio1 = _xlRange_StdBracket_Ratio.get_Range("A1", "A5");
                //_xlRange_StdBracket_Ratio1 = _xlRange_StdBracket_Ratio.get_Range("A1", "A6");

                _xlRange_StdBracket_Ratio2 = _xlRange_StdBracket_Ratio.get_Range("A7", "A10");
                _xlRange_StdBracket_Ratio3 = _xlRange_StdBracket_Ratio.get_Range("A11", "A16");

                _xlRange_StdBracket_Ratio4 = _xlRange_StdBracket_Ratio.get_Range("A17", "A23");
                //_xlRange_StdBracket_Ratio5 = _xlRange_StdBracket_Ratio.get_Range("A25", "A30");
                //_xlRange_StdBracket_Ratio6 = _xlRange_StdBracket_Ratio.get_Range("A31", "A36");

                _xlRange_StdBracket_SpecifiedAmount1.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount1";
                _xlRange_StdBracket_SpecifiedAmount2.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount2";
                _xlRange_StdBracket_SpecifiedAmount3.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount3";

                _xlRange_StdBracket_SpecifiedAmount4.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount4";
                //_xlRange_StdBracket_SpecifiedAmount5.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount5";
                //_xlRange_StdBracket_SpecifiedAmount6.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount6";

                _xlRange_StdBracket_Ratio1.Name = _wsComponent.Name + "_StdBracket_Ratio1";
                _xlRange_StdBracket_Ratio2.Name = _wsComponent.Name + "_StdBracket_Ratio2";
                _xlRange_StdBracket_Ratio3.Name = _wsComponent.Name + "_StdBracket_Ratio3";

                _xlRange_StdBracket_Ratio4.Name = _wsComponent.Name + "_StdBracket_Ratio4";
                //_xlRange_StdBracket_Ratio5.Name = _wsComponent.Name + "_StdBracket_Ratio5";
                //_xlRange_StdBracket_Ratio6.Name = _wsComponent.Name + "_StdBracket_Ratio6";




                _xlSerie_Stds1.XValues = _xlRange_StdBracket_SpecifiedAmount1.Cells;
                _xlSerie_Stds2.XValues = _xlRange_StdBracket_SpecifiedAmount2.Cells;
                _xlSerie_Stds3.XValues = _xlRange_StdBracket_SpecifiedAmount3.Cells;
                _xlSerie_Stds4.XValues = _xlRange_StdBracket_SpecifiedAmount4.Cells;
                //_xlSerie_Stds5.XValues = _xlRange_StdBracket_SpecifiedAmount5.Cells;
                //_xlSerie_Stds6.XValues = _xlRange_StdBracket_SpecifiedAmount6.Cells;

                _xlSerie_Stds1.Values = _xlRange_StdBracket_Ratio1.Cells;
                _xlSerie_Stds2.Values = _xlRange_StdBracket_Ratio2.Cells;
                _xlSerie_Stds3.Values = _xlRange_StdBracket_Ratio3.Cells;
                _xlSerie_Stds4.Values = _xlRange_StdBracket_Ratio4.Cells;
                //_xlSerie_Stds5.Values = _xlRange_StdBracket_Ratio5.Cells;
                //_xlSerie_Stds6.Values = _xlRange_StdBracket_Ratio6.Cells;

                string _serieStdName1 = "", _serieStdName2 = "", _serieStdName3 = "", _serieStdName4 = ""; //, _serieStdName5 = "", _serieStdName6 = "";

                string _axeStdXNumFormat = "0.00",
                       _axeStdYNumFormat = "0.00";

                string _axeStdXCaption = "%, Predicted ratio",
                       _axeStdYCaption = "%, Area Ratio";

                if (ExperimentInfo._quanType == "ABSOLUTE")
                {
                    _axeStdXNumFormat = "0.0"; _axeStdYNumFormat = "0.0";

                    _axeStdXCaption = "Predicted ratio"; _axeStdYCaption = "Area Ratio";
                }

                //_serieStdName1 = "E3 (4) 4-28-07";
                //_serieStdName2 = "E3 (4) 5-16-07";
                //_serieStdName3 = "E3 (4) 5-17-07";

                _serieStdName1 = "E33 CSF";
                _serieStdName2 = "E44 CSF";
                _serieStdName3 = "EMix CSF";

                _serieStdName4 = "Synth";
                //_serieStdName5 = "E3 (4) 5-16-07";
                //_serieStdName6 = "E4 5-05-08";



                _xlAxisX_Std.TickLabels.NumberFormat = _axeStdXNumFormat; _xlAxisY_Std.TickLabels.NumberFormat = _axeStdYNumFormat;
                _xlAxisX_Std.AxisTitle.Caption = _axeStdXCaption; _xlAxisY_Std.AxisTitle.Caption = _axeStdYCaption;

                _xlSerie_Stds1.Name = _serieStdName1;
                _xlSerie_Stds2.Name = _serieStdName2;
                _xlSerie_Stds3.Name = _serieStdName3;
                _xlSerie_Stds4.Name = _serieStdName4;
                //_xlSerie_Stds5.Name = _serieStdName5;
                //_xlSerie_Stds6.Name = _serieStdName6;

                Excel.Trendlines _xlTrendlines_Stds1 = (Excel.Trendlines)_xlSerie_Stds1.Trendlines(Type.Missing);
                Excel.Trendlines _xlTrendlines_Stds2 = (Excel.Trendlines)_xlSerie_Stds2.Trendlines(Type.Missing);
                Excel.Trendlines _xlTrendlines_Stds3 = (Excel.Trendlines)_xlSerie_Stds3.Trendlines(Type.Missing);
                Excel.Trendlines _xlTrendlines_Stds4 = (Excel.Trendlines)_xlSerie_Stds4.Trendlines(Type.Missing);
                //Excel.Trendlines _xlTrendlines_Stds5 = (Excel.Trendlines)_xlSerie_Stds5.Trendlines(Type.Missing);
                //Excel.Trendlines _xlTrendlines_Stds6 = (Excel.Trendlines)_xlSerie_Stds6.Trendlines(Type.Missing);


                Excel.XlTrendlineType _XlTrendlineType =
                    Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear;


                _xlTrendlines_Stds1.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, true, true, Type.Missing);
                _xlTrendlines_Stds2.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, true, true, Type.Missing);
                _xlTrendlines_Stds3.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, true, true, Type.Missing);
                //_xlTrendlines_Stds4.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                //    Type.Missing, true, true, Type.Missing);
                //_xlTrendlines_Stds5.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                //    Type.Missing, true, true, Type.Missing);

                //****************

                // };


                //Unknowns
                _xlRange_Unknowns =
                    _wsComponent.get_Range
                    ("A" + (Title + _StdBracketCount + _BlankCount).ToString(), "AA" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknowns.Name = _wsComponent.Name + "_Unknowns";

                _xlRange_Unknown_Area =
                    _wsComponent.get_Range
                    ("D" + (Title + _StdBracketCount + _BlankCount).ToString(), "D" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_Area.Name = _wsComponent.Name + "_Unknown_Area";

                _xlRange_Unknown_ISArea =
                    _wsComponent.get_Range
                    ("E" + (Title + _StdBracketCount + _BlankCount).ToString(), "E" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_ISArea.Name = _wsComponent.Name + "_Unknown_ISArea";

                _xlRange_Unknown_Ratio =
                   _wsComponent.get_Range
                   ("F" + (Title + _StdBracketCount + _BlankCount).ToString(), "F" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_Ratio.Name = _wsComponent.Name + "_Unknown_Ratio";

                string strCalcOrRAtio = "H";
                if (_StdBracketCount == 0) { strCalcOrRAtio = "F"; }
                _xlRange_Unknown_CalcAmount =
                    _wsComponent.get_Range
                    (strCalcOrRAtio + (Title + _StdBracketCount + _BlankCount).ToString(), strCalcOrRAtio + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_CalcAmount.Name = _wsComponent.Name + "_Unknown_CalcAmount";

                _xlRange_Unknown_SampleID =
                    _wsComponent.get_Range
                    ("J" + (Title + _StdBracketCount + _BlankCount).ToString(), "J" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_SampleID.Name = _wsComponent.Name + "_Unknown_SampleID";



                // 2014-01-21

                Excel.Range _xlRange_byStd1, _xlRange_byStd2, _xlRange_byStd3, _xlRange_byStd4, _xlRange_Subject;

                _xlRange_Subject = _xlRange_Unknown_SampleID.get_Offset(0, 1);
                _xlRange_Subject.Name = _wsComponent.Name + "_Subjects";

                _xlRange_byStd1 = _xlRange_Unknown_Ratio.get_Offset(0, 10);
                _xlRange_byStd1.Formula =
                    String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                             _wsComponent.Name + "_Unknown_Ratio",
                             _wsComponent.Name + "_StdBracket_Ratio1",
                             _wsComponent.Name + "_StdBracket_SpecifiedAmount1");
                _xlRange_byStd1.Name = _wsComponent.Name + "_byStd1";

                _xlRange_byStd2 = _xlRange_Unknown_Ratio.get_Offset(0, 11);
                _xlRange_byStd2.Formula =
                    String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                             _wsComponent.Name + "_Unknown_Ratio",
                             _wsComponent.Name + "_StdBracket_Ratio2",
                             _wsComponent.Name + "_StdBracket_SpecifiedAmount2");
                _xlRange_byStd2.Name = _wsComponent.Name + "_byStd2";

                _xlRange_byStd3 = _xlRange_Unknown_Ratio.get_Offset(0, 12);
                _xlRange_byStd3.Formula =
                String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                         _wsComponent.Name + "_Unknown_Ratio",
                         _wsComponent.Name + "_StdBracket_Ratio3",
                         _wsComponent.Name + "_StdBracket_SpecifiedAmount3");
                _xlRange_byStd3.Name = _wsComponent.Name + "_byStd3";

                _xlRange_byStd4 = _xlRange_Unknown_Ratio.get_Offset(0, 13);
                _xlRange_byStd4.Formula =
                String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                         _wsComponent.Name + "_Unknown_Ratio",
                         _wsComponent.Name + "_StdBracket_Ratio4",
                         _wsComponent.Name + "_StdBracket_SpecifiedAmount4");
                _xlRange_byStd4.Name = _wsComponent.Name + "_byStd4";
                /*
                                _xlRange_byStd5 = _xlRange_StdBracket_Ratio.get_Offset(0, 14);
                                _xlRange_byStd5.Formula =
                                String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                                         _wsComponent.Name + "_StdBracket_Ratio",
                                         _wsComponent.Name + "_StdBracket_Ratio5",
                                         _wsComponent.Name + "_StdBracket_SpecifiedAmount5");
                                _xlRange_byStd5.Name = _wsComponent.Name + "_byStd5";
                */
                // 2013-12-10
                /*
                                Excel.Range _xlRange_byStd1, _xlRange_byStd2, _xlRange_byStd3, _xlRange_Subject;

                                _xlRange_Subject = _xlRange_Unknown_SampleID.get_Offset(0, 1);
                                _xlRange_Subject.Name = _wsComponent.Name + "_Subjects";

                                _xlRange_byStd1 = _xlRange_StdBracket_Ratio.get_Offset(0, 10);
                                _xlRange_byStd1.Formula =
                                    String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                                             _wsComponent.Name + "_Unknown_Ratio",
                                             _wsComponent.Name + "_StdBracket_Ratio1",
                                             _wsComponent.Name + "_StdBracket_SpecifiedAmount1" ) ;
                                _xlRange_byStd1.Name = _wsComponent.Name + "_byStd1";

                                _xlRange_byStd2 = _xlRange_StdBracket_Ratio.get_Offset(0, 11);
                                _xlRange_byStd2.Formula =
                                    String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                                             _wsComponent.Name + "_Unknown_Ratio",
                                             _wsComponent.Name + "_StdBracket_Ratio2",
                                             _wsComponent.Name + "_StdBracket_SpecifiedAmount2");
                                _xlRange_byStd2.Name = _wsComponent.Name + "_byStd2";

                                _xlRange_byStd3 = _xlRange_StdBracket_Ratio.get_Offset(0, 12);
                                _xlRange_byStd3.Formula =
                                String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                                         _wsComponent.Name + "_Unknown_Ratio",
                                         _wsComponent.Name + "_StdBracket_Ratio3",
                                         _wsComponent.Name + "_StdBracket_SpecifiedAmount3");
                                _xlRange_byStd3.Name = _wsComponent.Name + "_byStd3";
                */

                //*************************Format Ranges**************************
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;


                _xlRange_Unknowns.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                _xlRange_Unknowns.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                // _xlRange_Unknowns.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color. .Chocolate);

                //******Sort Unknowns
                _xlRange_Unknowns.Sort("Column N",
                    Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    "Column J", Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                    Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);

                //*************2011-07-19**************************************

                //_xlRange_Unknown_ISArea.FormatConditions.AddDatabar();

                Excel.ColorScale cfColorScale = (Excel.ColorScale)(_xlRange_Unknown_ISArea.FormatConditions.AddColorScale(3));

                // Set the minimum threshold to red (0x000000FF) and maximum threshold
                // to blue (0x00FF0000).
                cfColorScale.ColorScaleCriteria[1].FormatColor.Color = 7039480;  //0x000000FF;  
                cfColorScale.ColorScaleCriteria[2].FormatColor.Color = 8711167;  // 0x00FF0000;
                cfColorScale.ColorScaleCriteria[3].FormatColor.Color = 13011546; // 0x00FF0000;

                //*************************Charts*********************************
                //
                // Area Ratio
                Excel.ChartObject _xlChartTimeCorse = (Excel.ChartObject)_xlCharts.Add(
                                   0, _graphOffset, 950, 250);
                Excel.Chart _xlChartTimeCorsePage = _xlChartTimeCorse.Chart;

                _xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlColumnClustered;
                //_xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlXYScatter;

                _xlChartTimeCorsePage.HasTitle = true;
                _xlChartTimeCorsePage.ChartTitle.Font.Size = 11;

                _xlChartTimeCorsePage.ChartTitle.Caption = _componentName + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_TP = (Excel.Axes)_xlChartTimeCorsePage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_TP.HasMajorGridlines = true;

                //_xlAxisX_TP.HasTitle = true;
                //_xlAxisY_TP.HasTitle = true;

                _xlChartTimeCorsePage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartTimeCorsePage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_TimeCourse = (Excel.SeriesCollection)_xlChartTimeCorsePage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_TTR = (Excel.Series)_xlSeriesColl_TimeCourse.NewSeries();


                _xlSeries_TTR.XValues = _xlRange_Subject.Cells;
                _xlSeries_TTR.Values = _xlRange_Unknown_Ratio;


                //_xlSeries_TTR.XValues = _xlRange_StdBracket_Ratio.get_Offset(0, 5).Cells;
                //_xlSeries_TTR.Values = _xlRange_StdBracket_Ratio;



                string _serieCurveName = "H:L Abeta";

                string _axeCurveXNumFormat = "0",
                       _axeCurveYNumFormat = "0.0%";


                if (ExperimentInfo._quanType == "ABSOLUTE")
                {
                    _axeCurveYNumFormat = "0.0";
                    _serieCurveName = "Abeta level";
                }

                
                _xlAxisX_TP.TickLabels.NumberFormat = _axeCurveXNumFormat;
                _xlAxisY_TP.TickLabels.NumberFormat = _axeCurveYNumFormat;
                                
                _xlSeries_TTR.Name = _serieCurveName;
                _xlSeries_TTR.Name = _componentName;
                
              
                // Normalized
                
                Excel.ChartObject _xlChartTimeCorseNorm = (Excel.ChartObject)_xlCharts.Add(
                                   0, _graphOffset + _xlChartTimeCorse.Height, 950, 250);
                Excel.Chart _xlChartTimeCorseNormPage = _xlChartTimeCorseNorm.Chart;

                _xlChartTimeCorseNormPage.ChartType = Excel.XlChartType.xlColumnClustered;
                
                _xlChartTimeCorseNormPage.HasTitle = true;
                _xlChartTimeCorseNormPage.ChartTitle.Font.Size = 11;

                _xlChartTimeCorseNormPage.ChartTitle.Caption = _componentName + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_TP_Norm = (Excel.Axes)_xlChartTimeCorseNormPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_TP_Norm = _xlAxes_TP_Norm.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_TP_Norm = _xlAxes_TP_Norm.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_TP_Norm.HasMajorGridlines = true;

                _xlAxisX_TP_Norm.TickLabels.NumberFormat = "0";
                _xlAxisY_TP_Norm.TickLabels.NumberFormat = "0.0E+0";

                _xlChartTimeCorseNormPage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartTimeCorseNormPage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_TimeCourseNorm = (Excel.SeriesCollection)_xlChartTimeCorseNormPage.SeriesCollection(Type.Missing);



                Excel.Series _xlSeries_CalcAm_byStd1 = (Excel.Series)_xlSeriesColl_TimeCourseNorm.NewSeries();

                _xlSeries_CalcAm_byStd1.XValues = _xlRange_Subject.Cells;
                _xlSeries_CalcAm_byStd1.Values = _xlRange_byStd1.Cells;
                _xlSeries_CalcAm_byStd1.Name = "by E33 CSF";

                Excel.Series _xlSeries_CalcAm_byStd2 = (Excel.Series)_xlSeriesColl_TimeCourseNorm.NewSeries();
                _xlSeries_CalcAm_byStd2.XValues = _xlRange_Subject.Cells;
                _xlSeries_CalcAm_byStd2.Values = _xlRange_byStd2.Cells;
                _xlSeries_CalcAm_byStd2.Name = "by E44 CSF";

                Excel.Series _xlSeries_CalcAm_byStd3 = (Excel.Series)_xlSeriesColl_TimeCourseNorm.NewSeries();
                _xlSeries_CalcAm_byStd3.XValues = _xlRange_Subject.Cells;
                _xlSeries_CalcAm_byStd3.Values = _xlRange_byStd3.Cells;
                _xlSeries_CalcAm_byStd3.Name = "by Mix CSF";

                Excel.Series _xlSeries_CalcAm_byStd4 = (Excel.Series)_xlSeriesColl_TimeCourseNorm.NewSeries();
                _xlSeries_CalcAm_byStd4.XValues = _xlRange_Subject.Cells;
                _xlSeries_CalcAm_byStd4.Values = _xlRange_byStd4.Cells;
                _xlSeries_CalcAm_byStd4.Name = "by Syn pep";
                  
                _xlRange_byStd1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                _xlRange_byStd2.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                _xlRange_byStd3.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);

                Excel.ChartObjects _xlChartsSummary = (Excel.ChartObjects)_wsSummary.ChartObjects(Type.Missing);
                Excel.ChartObject _xlChartCurveSummary;
                Excel.Chart _xlChartCurveSummaryPage;

                if (_xlChartsSummary.Count < 1)
                {
                    _xlChartCurveSummary = (Excel.ChartObject)_xlChartsSummary.Add(
                                   10, 10, 850, 500);
                    _xlChartCurveSummaryPage = _xlChartCurveSummary.Chart;
                    
                    _xlChartCurveSummaryPage.ChartType = Excel.XlChartType.xlXYScatterLines;
                    
                    _xlChartCurveSummaryPage.HasTitle = true;
                    _xlChartCurveSummaryPage.ChartTitle.Font.Size = 11;
                    
                    Excel.Axes _xlAxes_CurveSummary = (Excel.Axes)_xlChartCurveSummaryPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisX_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisY_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                    _xlAxisX_CurveSummary.HasMajorGridlines = true;
                    _xlAxisY_CurveSummary.HasMajorGridlines = true;

                    _xlAxisX_CurveSummary.Format.Line.Weight = 3.0F;
                    _xlAxisY_CurveSummary.Format.Line.Weight = 3.0F;

                    
                    _xlAxisX_CurveSummary.TickLabels.NumberFormat = _axeCurveXNumFormat;
                    _xlAxisY_CurveSummary.TickLabels.NumberFormat = _axeCurveYNumFormat;


                    _xlChartCurveSummaryPage.ChartTitle.Caption = "H:L Abeta ";

                    if (ExperimentInfo._quanType == "ABSOLUTE")
                    {
                        _xlChartCurveSummaryPage.ChartTitle.Caption = "ApoE levels in ";
                        try
                        {
                            // _xlAxisY_CurveSummary.ScaleType = Excel.XlScaleType.xlScaleLogarithmic;
                            // _xlAxisY_CurveSummary.LogBase = 10;
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

                //_xlSeries_CurveTTR.XValues = _xlRange_StdBracket_Ratio.get_Offset(0, 5).Cells;

                _xlSeries_CurveTTR.XValues = _xlRange_Unknown_SampleID.get_Offset(0, 1).Cells; ;

                //_xlSeries_CurveTTR.Values = _xlRange_Unknown_CalcAmount;

                _xlSeries_CurveTTR.Values = _xlRange_Unknown_Ratio;

                //_xlSeries_CurveTTR.Values = _xlRange_StdBracket_Ratio;

                _xlSeries_CurveTTR.Name = (string)_xlRange_Compound.Cells.Value2;


                toolStripProgressBarMain.PerformStep();

            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = false;

            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();

            _XXQN.Close();

            object misValue = System.Reflection.Missing.Value;

            try
            {
                _XXQN.Close();

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
                    ((Excel._Worksheet)_wsSummary).Activate();

                }
                else
                {
                    _xlApp.Quit();
                }
            }
        }


        private void ExportAPOE_ABS()
        {

            List<string> _strExportSQL = new List<string>();
            /*05-07-2010*/
            string _strAnalyte = null;
            double _Area = 0, _ISTD_Area = 0, _Area_Ratio = 0, _CalcAmount = 0;
            decimal dcm_Area_Ratio = 0;
            string _strSpecAmount = null, _strRT = null, _strISTD_RT = null;
            string _response = null;

            /*bool IsPeakFound = false;*/

            /*05-07-2010*/
            if (textBox_XQNfile.Text == "")
            {
                MessageBox.Show("Select Input XQN file, please");
                return;
            }


            _XXQN = new XCALIBURFILESLib.XXQNClass();

            try
            {
                _XXQN.Open(textBox_XQNfile.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("File: " + textBox_XQNfile.Text + " is invalid", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //MessageBox.Show(_exception.Message, _exception.GetBaseException().Message);
                return;
            }
            finally
            {

            }

            XCALIBURFILESLib.IXHeader _IXHeader = (XCALIBURFILESLib.IXHeader)_XXQN.Header;

            ExperimentInfo._ProcessedBy = _IXHeader.ChangedId;

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = _IXHeader.ChangedLogon;
            }

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = "unknown";
            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = true;
            toolStripStatusLabelMain.Visible = true;



            _xlApp = new Excel.ApplicationClass();
            _xlApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;



            Excel.Workbook xlWB;

            try
            {

                xlWB =
                    _xlApp.Workbooks.Add(
                    //@"c:\Project\XQNReader\XQNReader\LOAD_TSQ_Template_Summary.xlsx"
                Application.StartupPath + Path.DirectorySeparatorChar +
                @"\Templates\ApoE_TSQ_ABS_Template_Summary.xlsx"
                );

            }

            catch (Exception)
            {
                MessageBox.Show("File: " + Application.StartupPath + Path.DirectorySeparatorChar +
                @"\Templates\ApoE_TSQ_ABS_Template_Summary.xlsx", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //MessageBox.Show(_exception.Message, _exception.GetBaseException().Message);
                return;
            }
            finally
            {

            }

            Excel.Worksheet _wsSummary = (Excel.Worksheet)xlWB.ActiveSheet;

            Excel.Worksheet _wsCurrent = null;


            byte Title = 5;
            byte _UnknownCount = 0, _StdBracketCount = 0, _BlankCount = 0, _TotalCount = 0;




            XCALIBURFILESLib.XBracketGroups _XBracketGroups
                = (XCALIBURFILESLib.XBracketGroups)_XXQN.BracketGroups;

            XCALIBURFILESLib.XBracketGroup _XBracketGroup
                = (XCALIBURFILESLib.XBracketGroup)_XBracketGroups.Item(1);

            XCALIBURFILESLib.XQuanResults _XQuanResults
                = (XCALIBURFILESLib.XQuanResults)_XBracketGroup.QuanResults;


            var _XQuanResultsSorted = from XCALIBURFILESLib.XQuanResult t in _XQuanResults
                                      //orderby t.Sample
                                      select t;


            //********_XQuanResults*******************************

            foreach (XCALIBURFILESLib.XQuanResult _XQuanResult in _XQuanResultsSorted /*_XQuanResults*/)
            {
                toolStripProgressBarMain.Maximum = _XQuanResults.Count;

                XCALIBURFILESLib.XSample _Sample
                    = (XCALIBURFILESLib.XSample)_XQuanResult.Sample;


                //_Sample.




                _TotalCount++;
                switch (_Sample.SampleType)
                {
                    case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                        _UnknownCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                        _StdBracketCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleBlank:
                        _BlankCount++;
                        break;
                    default:
                        break;
                }

                XCALIBURFILESLib.XQuanComponents _XQuanSelectedComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Selected; /*  Method; */
                //  = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.User;  
                ////////////////////////////////////////
                XCALIBURFILESLib.XQuanComponents _XQuanMethodComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Method; /*  Method; */


                ////////////////////////////////////////
                short _comNo = 1;
                foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponent in _XQuanSelectedComponents)
                {
                    XCALIBURFILESLib.XComponent _XComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanComponent.Component;

                    XCALIBURFILESLib.XQuanComponent _XQuanMethodComponent = (XCALIBURFILESLib.XQuanComponent)
                    _XQuanMethodComponents.Item(_comNo);
                    _comNo++;

                    XCALIBURFILESLib.XComponent _XMethodComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanMethodComponent.Component;

                    /*
                    XCALIBURFILESLib.XComponent _XComponentUser =
                        (XCALIBURFILESLib.XComponent)_XQuanComponent.Component;
                    */

                    //****************** 
                    XCALIBURFILESLib.IXDetection2 _IXDetection =
                        (XCALIBURFILESLib.IXDetection2)_XComponent.Detection;
                    //*****************

                    XCALIBURFILESLib.XQuanCalibration _XQuanCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XComponent.Calibration;

                    XCALIBURFILESLib.XQuanCalibration _XQuanMethodCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XMethodComponent.Calibration;

                    string _quanTypeKey = "";

                    // MessageBox.Show(ExperimentInfo._quanType.ToString());

                    if (ExperimentInfo._quanType == "RELATIVE")
                    {
                        _quanTypeKey = "R0:N14";
                    }
                    if (ExperimentInfo._quanType == "ABSOLUTE")
                    {
                        _quanTypeKey = "L6:C13";
                    }

                    //MessageBox.Show(_XQuanCalibration.ComponentType.ToString());
                    //MessageBox.Show(_XComponent.Name.Contains(_quanTypeKey).ToString());

                    if (_XQuanCalibration.ComponentType != XCALIBURFILESLib.XComponentType.XISTD
                        && !_XComponent.Name.Contains(_quanTypeKey))
                    {

                        // MessageBox.Show(_quanTypeKey);

                        _strAnalyte = _XComponent.Name.Replace("_", " "); //*******************

                        // ****** 2010-07-23
                        _strAnalyte = _strAnalyte.Replace("AB-Total", "AbetaTotal");
                        // ****** 2010-07-23


                        bool _SheetExist = false;
                        string _ComponentName = _XComponent.Name.Replace(":", "_");
                        _ComponentName = _ComponentName.Replace("-", "_");
                        //**********
                        _ComponentName = _ComponentName.Replace("(", "_");
                        _ComponentName = _ComponentName.Replace(")", "_");
                        //**********

                        foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
                        {

                            if (_wsComponent.Name == _ComponentName
                                )
                            {
                                _wsCurrent = _wsComponent;
                                _SheetExist = true;
                            }
                        }



                        if (!_SheetExist)
                        {
                            _wsCurrent = (Excel.Worksheet)xlWB.Sheets.Add
                          (Type.Missing, Type.Missing, Type.Missing,
                          Application.StartupPath + Path.DirectorySeparatorChar +
                            @"\Templates\ApoE_TSQ_ABS_Template.xlsx"
                          );

                            _wsCurrent.Name = _ComponentName;



                            _wsCurrent.Cells[2, 1] = _XComponent.Name.Replace("_", " ");

                            _wsCurrent.Cells[2, 3] =
                                CurveTypeToStr(_XQuanMethodCalibration.CurveType);
                            _wsCurrent.Cells[2, 4] =
                                weightingTypeToStr(_XQuanMethodCalibration.Weighting);
                            _wsCurrent.Cells[2, 6] =
                                OriginTypeToStr(_XQuanMethodCalibration.OriginType);

                            _wsCurrent.Cells[2, 8] = _XQuanMethodCalibration.FullEquation;

                        }

                        string _subjectNum = "";

                        if (checkBox_if_multi_subject.Checked && (int)_Sample.SampleType == 0)
                        {
                            //_subjectNum = _Sample.FileName.Split('_')[1].Replace("CSF", "");
                            //_subjectNum = _Sample.FileName.Split('_')[0].Replace("CSF", "");
                            //_subjectNum = _Sample.SampleId;
                            _subjectNum = _Sample.SampleName;
                        }
                        else
                        {
                            /*
                            if (ExperimentInfo._Subject.Contains("-"))
                            {
                                _subjectNum = ExperimentInfo._Subject.Split('-')[1];
                            }
                            else _subjectNum = ExperimentInfo._Subject;

                            */

                        }

                        _wsCurrent.Cells[_Sample.RowNumber + Title, 1] = _Sample.FileName;
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 2] = SampleTypeToStr(_Sample.SampleType);
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 3] = IntegTypeToStr(_XQuanComponent.Selected_Type);


                        XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeak;


                        bool _IsPeakFoundANALYTE = false;
                        bool _IsPeakFoundISTD = false;
                        int _shift = 2;

                        if (_XQuanComponent.DoesDetectedQuanPeakExist > 0)
                        {
                            _IsPeakFoundANALYTE = true;
                            _XDetectedQuanPeak =
                                (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponent.DetectedQuanPeak;

                            _wsCurrent.Cells[_Sample.RowNumber + Title, 4] = _XDetectedQuanPeak.Area;
                            _Area = _XDetectedQuanPeak.Area;

                            //_XDetectedQuanPeak.ISTDValid


                            _wsCurrent.Cells[_Sample.RowNumber + Title, 8] = _XDetectedQuanPeak.CalcAmount;

                            _CalcAmount = _XDetectedQuanPeak.CalcAmount;


                            /////
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 20 + _shift] = _XDetectedQuanPeak.LeftRT;
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 21 + _shift] = _XDetectedQuanPeak.RightRT;

                            _wsCurrent.Cells[_Sample.RowNumber + Title, 19 + _shift] = String.Format("=L{0}-M{0}", (_Sample.RowNumber + Title).ToString());


                            if (_XDetectedQuanPeak.ISTDArea != 0)
                            {
                                _IsPeakFoundISTD = true;
                                _wsCurrent.Cells[_Sample.RowNumber + Title, 5] = _XDetectedQuanPeak.ISTDArea;
                                _wsCurrent.Cells[_Sample.RowNumber + Title, 6] =
                                _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;

                                _ISTD_Area = _XDetectedQuanPeak.ISTDArea;
                                _Area_Ratio = _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;
                                dcm_Area_Ratio = (decimal)_XDetectedQuanPeak.Area / (decimal)_XDetectedQuanPeak.ISTDArea;
                            }
                            else
                            {
                                _wsCurrent.Cells[_Sample.RowNumber + Title, 5] = "N/F";
                            }



                            //string _response = "";
                            _response = "NULL"; _strRT = "NULL"; _strISTD_RT = "NULL";

                            if (_XDetectedQuanPeak.IsResponseOK > 0) { _response = "Ok"; }
                            if (_XDetectedQuanPeak.IsResponseHigh > 0) { _response = "High"; }
                            if (_XDetectedQuanPeak.IsResponseLow > 0) { _response = "Low"; }
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 9] = _response;
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 12] = _XDetectedQuanPeak.ApexRT;

                            _strRT = _XDetectedQuanPeak.ApexRT.ToString("F2");
                            //_RT = _XDetectedQuanPeak.ApexRT;
                        }
                        else
                        {
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 9] = "Not Found";

                        }



                        string _sample_ID = "N/A";
                        double _specAmount = 0;
                        switch (_Sample.SampleType)
                        {
                            case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                                _sample_ID = _Sample.Level;
                                XCALIBURFILESLib.XCalibrationCurveData _XCalibrationCurveData =
                                (XCALIBURFILESLib.XCalibrationCurveData)_XQuanComponent.CalibrationCurveData;
                                XCALIBURFILESLib.XReplicateRows _XReplicateRows =
                                (XCALIBURFILESLib.XReplicateRows)_XCalibrationCurveData.ReplicateRows;

                                foreach (XCALIBURFILESLib.XReplicateRow _XReplicateRow in _XReplicateRows)
                                {
                                    if (_XReplicateRow.LevelName == _Sample.Level)
                                    {
                                        _specAmount = _XReplicateRow.Amount;
                                        _strSpecAmount = "'" + _XReplicateRow.Amount.ToString("F5") + "'";
                                    }
                                }

                                _wsCurrent.Cells[_Sample.RowNumber + Title, 11] = _Sample.FileName.Split('_')[1];

                                break;

                            case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:

                                //_sample_ID = _Sample.FileName.Split('_')[2].Replace("Hr", "");

                                //_sample_ID = _Sample.SampleName;

                                //_sample_ID = "0"; 

                                _sample_ID = "-1";

                                _specAmount = -1;
                                _strSpecAmount = "NULL";

                                _wsCurrent.Cells[_Sample.RowNumber + Title, 11] = _subjectNum;
                                //   _wsCurrent.Cells[_Sample.RowNumber + Title, 11] = _Sample.FileName.Substring(12, 5);

                                break;

                            case XCALIBURFILESLib.XSampleTypes.XSampleBlank:
                                _sample_ID = _Sample.SampleId;
                                _specAmount = -1;
                                _strSpecAmount = "NULL";
                                break;
                            default:
                                break;
                        }

                        //**********
                        //_Sample.
                        //*********
                        bool _IsSampleIDDefinded = false;
                        int _intTP = -2;



                        if (_sample_ID == "")
                        {
                            _sample_ID = _Sample.FileName.Split('_')[2].Replace("Hr", "");

                            //_Sample.Level

                        };



                        try
                        {
                            _intTP = Convert.ToInt32(_sample_ID);

                            //_intTP = -1;
                        }
                        catch (System.FormatException _FoEx)
                        {
                            MessageBox.Show(_sample_ID + " || " + _FoEx.Message);
                        };


                        if (_intTP < -1)
                        {
                            _sample_ID = "";
                            _IsSampleIDDefinded = false;
                        }
                        else
                        {
                            _sample_ID = _intTP.ToString();
                            _IsSampleIDDefinded = true;
                        }




                        if (_specAmount != -1)
                        {
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 7] = _specAmount;
                        }
                        else
                        {
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 7] = "";
                        }
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 10] = _sample_ID;




                        foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponentISTD in _XQuanSelectedComponents)
                        {
                            XCALIBURFILESLib.XComponent _XISTD =
                            (XCALIBURFILESLib.XComponent)_XQuanComponentISTD.Component;

                            if (_XISTD.Name == _XQuanCalibration.ISTD)
                            {
                                if (_XQuanComponentISTD.DoesDetectedQuanPeakExist > 0)
                                {
                                    XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeakISTD =
                                        (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponentISTD.DetectedQuanPeak;
                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 13] = _XDetectedQuanPeakISTD.ApexRT;

                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 23 + _shift] = _XDetectedQuanPeakISTD.LeftRT;
                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 24 + _shift] = _XDetectedQuanPeakISTD.RightRT;



                                    _strISTD_RT = _XDetectedQuanPeakISTD.ApexRT.ToString("F2");
                                }
                                else
                                {
                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 13] = "N/F";
                                }

                            }

                        }

                        _wsCurrent.Cells[_Sample.RowNumber + Title, 14] = _Sample.SampleId;
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 15] = _Sample.SampleName;

                        //_wsCurrent.Cells[_Sample.RowNumber + Title, 16] = _Sample.SampleName;





                        // if (GetSlopeOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType)) == -1)
                        // {
                        //     return;
                        // }


                        //*****
                        //_IsSampleIDDefinded = true;



                        if (_IsPeakFoundANALYTE && _IsPeakFoundISTD && _IsSampleIDDefinded && checkBox_ExportIntoDB.Checked)
                        {
                            _strExportSQL.Add("INSERT INTO TMP_TSQ_IMPORT ( ANYLYTE_NAME, FILENAME_DESC, SAMPLE_TYPE_ID, INTEG_TYPE_ID, AREA, ISTD_AREA, AREA_RATIO, " +
                                                                          "SPEC_AMOUNT, CALC_AMOUNT, LEVEL_TP, PEAK_SATUS, RT, ISTD_RT, QUAN_TYPE_ID, FLUID_TYPE_ID, " +
                                                                          "SUBJECT_NUM, PROCESS_BY_ID, ASSAY_DATE, ASSAY_TYPE_ID, ANTIBODY_ID, ENZYME_ID, DONE_BY, " +
                                                                          "SAMPLE_PROCESS_BY, EXPER_FILE, CURVE, WEIGHTING, ORIGIN, EQUATION, R, R_SQR, NUM_SLOPE, " +
                                                                          "NUM_INTERSECT )");
                            _strExportSQL.Add(String.Format("VALUES ('{0}', '{1}', {2}, '{3}', {4:F0}, {5:F0}, {6:F5}, {7:F}, {8:F5}, {9}, '{10}', {11}, {12}, {13}, {14}, '{15}', {16}, '{17}', {18}, {19}, {20}, {21}, {22}, '{23}'," +
                                " '{24}', '{25}', '{26}', '{27}', {28:F5}, {29:F5}, {30:F5}, {31:F5});",
                                   _strAnalyte,                                 // 0
                                   _Sample.FileName,                            // 1
                                   (int)_Sample.SampleType,                     // 2
                                   (int)_XQuanComponent.Selected_Type,          // 3
                                   _Area,                                       // 4
                                   _ISTD_Area,                                  // 5
                                   _Area_Ratio,                                 // 6
                                   _strSpecAmount,                              // 7 
                                   _CalcAmount,                                 // 8  "NULL", //
                                   _sample_ID,                                  // 9
                                   _response,                                   // 10
                                   _strRT,                                      // 11 
                                   _strISTD_RT,                                 // 12
                                   QuanTypeToDBInt(ExperimentInfo._quanType),
                                   FluidTypeToDBInt(ExperimentInfo._matrix),    // 14
                                   _subjectNum,       // 15
                                   ExperimentInfo._SampleQuantitatedBy_ID,      // 16
                                   ExperimentInfo._Date,                        // 17
                                   11,                                          // 18      { ASSAY_TYPE_ID   2 - "IP LC/MS (C-term)"    
                                   comboBox_Abody.SelectedValue,                // 19      { 3, // A-BODY  3 - HJ5.x
                                   comboBox_Enzyme.SelectedValue,               // 20      { 2, //Enzyme 2 - LysN
                                   ExperimentInfo._SampleRanBy_ID,              // 21      { Done by 1 - ovodvo
                                   ExperimentInfo._SamplePreparedBy_ID,         // 22      { Sample process by 1 - ovodvo
                                   ExperimentInfo._expName,                     // 23

                                   CurveTypeToStr(_XQuanMethodCalibration.CurveType),
                                   weightingTypeToStr(_XQuanMethodCalibration.Weighting),
                                   OriginTypeToStr(_XQuanMethodCalibration.OriginType),
                                   GetEquaOfEquation(_XQuanMethodCalibration.FullEquation, _strAnalyte),
                                   GetROfEquation(_XQuanMethodCalibration.FullEquation),
                                   GetR_SQROfEquation(_XQuanMethodCalibration.FullEquation),
                                   GetSlopeOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType)),
                                   GetInterOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType))


                                   ));

                        }
                    }
                }

                toolStripProgressBarMain.PerformStep();
                toolStripStatusLabelMain.Text = "Processing ..." + _Sample.FileName;
                toolStripStatusLabelMain.PerformClick();

            }

            //_xlApp.Visible = true;
            toolStripProgressBarMain.Visible = false;
            toolStripProgressBarMain.Value = 0;
            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();
            toolStripProgressBarMain.Visible = true;

            toolStripProgressBarMain.Maximum = xlWB.Worksheets.Count - 1;


            if (checkBox_ExportIntoDB.Checked)
            {

                string _AppPath = Application.StartupPath + Path.DirectorySeparatorChar;

                string[] _sql_P1 = File.ReadAllLines(_AppPath + "sql_TSQImport_P1.sql");
                string[] _sql_P2 = File.ReadAllLines(_AppPath + "sql_TSQImport_P2.sql");
                string[] _sql_P3 = File.ReadAllLines(_AppPath + "sql_TSQImport_P3.sql");

                _strExportSQL.Insert(0, " ");
                _strExportSQL.InsertRange(0, _sql_P1.ToList<string>());
                _strExportSQL.AddRange(_sql_P3.ToList<string>());
                //_strExportSQL.AddRange(_sql_P2.ToList<string>());


                File.WriteAllLines(Path.ChangeExtension(textBox_Excelfile.Text, "sql"), _strExportSQL.ToArray());
            }

            foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
            {
                _wsComponent.Name = "_" + _wsComponent.Name;

                toolStripStatusLabelMain.Text = "Formatting ..." + _wsComponent.Name;
                toolStripStatusLabelMain.PerformClick();

                if (_wsComponent.Name.Contains("Summary")) { continue; }

                Excel.ChartObjects _xlCharts = (Excel.ChartObjects)_wsComponent.ChartObjects(Type.Missing);
                string _componentName = _wsComponent.get_Range("A2", "A2").Value2.ToString();

                Excel.Range _xlRange_Compound,
                            _xlRange_StdBracket_Area,
                            _xlRange_StdBracket_ISArea,
                            _xlRange_StdBracket_SpecifiedAmount,
                            _xlRange_StdBracket_Ratio,
                            _xlRange_Unknown_Area,
                            _xlRange_Unknown_ISArea,
                            _xlRange_Unknown_Ratio,
                            _xlRange_Unknown_CalcAmount,
                            _xlRange_Unknown_SampleID,
                            _xlRange_STD,
                            _xlRange_Unknowns,
                            _xlRange_All,


                           _xlRange_StdBracket_SpecifiedAmount1,
                           _xlRange_StdBracket_SpecifiedAmount2,
                           _xlRange_StdBracket_SpecifiedAmount3,
                           _xlRange_StdBracket_SpecifiedAmount4,

                           _xlRange_StdBracket_Ratio1,
                           _xlRange_StdBracket_Ratio2,
                           _xlRange_StdBracket_Ratio3,
                           _xlRange_StdBracket_Ratio4;
                           


                _xlRange_Compound = _wsComponent.get_Range("A2", "A2");

                //All
                _xlRange_All =
                    _wsComponent.get_Range
                    ("A" + Title.ToString(), "Y" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_All.Name = _wsComponent.Name + "_All";


                double _rowHeight = (double)_wsCurrent.get_Range("A1", "A1").RowHeight;

                double _graphOffset = (double)_xlRange_All.Top + (double)_xlRange_All.Height +

                      (_rowHeight * 3);



                //Std
                // if (_StdBracketCount > 0)
                // {
                _xlRange_STD =
                    _wsComponent.get_Range
                    ("A" + Title.ToString(), "AA" + (Title - 1 + _StdBracketCount).ToString());
                _xlRange_STD.Name = _wsComponent.Name + "_STD";


                _xlRange_StdBracket_Ratio =
                    _wsComponent.get_Range
                    ("F" + Title.ToString(), "F" + (Title - 1 + _StdBracketCount).ToString());
                _xlRange_StdBracket_Ratio.Name = _wsComponent.Name + "_StdBracket_Ratio";

                _xlRange_StdBracket_Area =
                    _wsComponent.get_Range
                    ("D" + Title.ToString(), "D" + (Title - 1 + _StdBracketCount).ToString());
                _xlRange_StdBracket_Area.Name = _wsComponent.Name + "_StdBracket_Area";

                _xlRange_StdBracket_ISArea =
                    _wsComponent.get_Range
                    ("E" + Title.ToString(), "E" + (Title - 1 + _StdBracketCount).ToString());
                _xlRange_StdBracket_ISArea.Name = _wsComponent.Name + "_StdBracket_ISArea";

                _xlRange_StdBracket_SpecifiedAmount =
                   _wsComponent.get_Range
                   ("G" + Title.ToString(), "G" + (Title - 1 + _StdBracketCount).ToString());
                _xlRange_StdBracket_SpecifiedAmount.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount";


                _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                _xlRange_STD.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);

                //Unknowns
                /*   _xlRange_Unknowns = _wsComponent.get_Range
                       ("A" + (Title - 1 + _StdBracketCount).ToString(), "O" + (Title - 1 + _StdBracketCount + _UnknownCount).ToString());
                   _xlRange_STD.Name = _wsComponent.Name + "_Unknown"; */

                // v

                //*********Sort Range************
                _xlRange_STD.Sort(_xlRange_STD.Columns[14, Type.Missing], Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Microsoft.Office.Interop.Excel.XlYesNoGuess.xlGuess,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                    Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);

                /* _xlRange_Unknowns.Sort(_xlRange_Unknowns.Columns[10, Type.Missing],
                                        , Microsoft.Office.Interop.Excel.XlSortOrder.xlDescending,
                                        //_xlRange_Unknowns.Columns[10, Type.Missing], Type.Missing, Excel.XlSortOrder.xlAscending,

                                        Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,

                                         Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                                         Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                                         Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                                         Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                                         Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                                         Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                                         Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);*/


                //****************
                // StdCurve;


                Excel.ChartObject _xlChartStdCurve = (Excel.ChartObject)_xlCharts.Add(
                                   970, _graphOffset, 470, 250);
                Excel.Chart _xlChartStdCurvePage = _xlChartStdCurve.Chart;
                _xlChartStdCurvePage.ChartType = Excel.XlChartType.xlXYScatter;

                _xlChartStdCurvePage.HasTitle = true;
                _xlChartStdCurvePage.ChartTitle.Font.Size = 11;

                string _chartTitleStdCurve = String.Format
                    ("Calibration curve {0} {1} {2} {3} ({4})",
                    _componentName,
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

                Excel.Series _xlSerie_Stds1 = (Excel.Series)_xlSerColl_StdCurve.NewSeries();
                Excel.Series _xlSerie_Stds2 = (Excel.Series)_xlSerColl_StdCurve.NewSeries();
                Excel.Series _xlSerie_Stds3 = (Excel.Series)_xlSerColl_StdCurve.NewSeries();

                Excel.Series _xlSerie_Stds4 = (Excel.Series)_xlSerColl_StdCurve.NewSeries();
                // Excel.Series _xlSerie_Stds5 = (Excel.Series)_xlSerColl_StdCurve.NewSeries();
                // Excel.Series _xlSerie_Stds6 = (Excel.Series)_xlSerColl_StdCurve.NewSeries();


                //String Std1 = 

                _xlRange_StdBracket_SpecifiedAmount1 = _xlRange_StdBracket_SpecifiedAmount.get_Range("A1", "A5");
                //_xlRange_StdBracket_SpecifiedAmount1 = _xlRange_StdBracket_SpecifiedAmount.get_Range("A1", "A6");
                _xlRange_StdBracket_SpecifiedAmount1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);

                _xlRange_StdBracket_SpecifiedAmount2 = _xlRange_StdBracket_SpecifiedAmount.get_Range("A7", "A10");
                _xlRange_StdBracket_SpecifiedAmount2.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);


                _xlRange_StdBracket_SpecifiedAmount3 = _xlRange_StdBracket_SpecifiedAmount.get_Range("A11", "A16");
                _xlRange_StdBracket_SpecifiedAmount3.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);


                _xlRange_StdBracket_SpecifiedAmount4 = _xlRange_StdBracket_SpecifiedAmount.get_Range("A17", "A23");

                //_xlRange_StdBracket_SpecifiedAmount5 = _xlRange_StdBracket_SpecifiedAmount.get_Range("A25", "A30");
                //_xlRange_StdBracket_SpecifiedAmount6 = _xlRange_StdBracket_SpecifiedAmount.get_Range("A31", "A36");

                _xlRange_StdBracket_Ratio1 = _xlRange_StdBracket_Ratio.get_Range("A1", "A5");
                //_xlRange_StdBracket_Ratio1 = _xlRange_StdBracket_Ratio.get_Range("A1", "A6");

                _xlRange_StdBracket_Ratio2 = _xlRange_StdBracket_Ratio.get_Range("A7", "A10");
                _xlRange_StdBracket_Ratio3 = _xlRange_StdBracket_Ratio.get_Range("A11", "A16");

                _xlRange_StdBracket_Ratio4 = _xlRange_StdBracket_Ratio.get_Range("A17", "A23");
                //_xlRange_StdBracket_Ratio5 = _xlRange_StdBracket_Ratio.get_Range("A25", "A30");
                //_xlRange_StdBracket_Ratio6 = _xlRange_StdBracket_Ratio.get_Range("A31", "A36");

                _xlRange_StdBracket_SpecifiedAmount1.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount1";
                _xlRange_StdBracket_SpecifiedAmount2.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount2";
                _xlRange_StdBracket_SpecifiedAmount3.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount3";

                _xlRange_StdBracket_SpecifiedAmount4.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount4";
                //_xlRange_StdBracket_SpecifiedAmount5.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount5";
                //_xlRange_StdBracket_SpecifiedAmount6.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount6";

                _xlRange_StdBracket_Ratio1.Name = _wsComponent.Name + "_StdBracket_Ratio1";
                _xlRange_StdBracket_Ratio2.Name = _wsComponent.Name + "_StdBracket_Ratio2";
                _xlRange_StdBracket_Ratio3.Name = _wsComponent.Name + "_StdBracket_Ratio3";

                _xlRange_StdBracket_Ratio4.Name = _wsComponent.Name + "_StdBracket_Ratio4";
                //_xlRange_StdBracket_Ratio5.Name = _wsComponent.Name + "_StdBracket_Ratio5";
                //_xlRange_StdBracket_Ratio6.Name = _wsComponent.Name + "_StdBracket_Ratio6";




                _xlSerie_Stds1.XValues = _xlRange_StdBracket_SpecifiedAmount1.Cells;
                _xlSerie_Stds2.XValues = _xlRange_StdBracket_SpecifiedAmount2.Cells;
                _xlSerie_Stds3.XValues = _xlRange_StdBracket_SpecifiedAmount3.Cells;
                _xlSerie_Stds4.XValues = _xlRange_StdBracket_SpecifiedAmount4.Cells;
                //_xlSerie_Stds5.XValues = _xlRange_StdBracket_SpecifiedAmount5.Cells;
                //_xlSerie_Stds6.XValues = _xlRange_StdBracket_SpecifiedAmount6.Cells;

                _xlSerie_Stds1.Values = _xlRange_StdBracket_Ratio1.Cells;
                _xlSerie_Stds2.Values = _xlRange_StdBracket_Ratio2.Cells;
                _xlSerie_Stds3.Values = _xlRange_StdBracket_Ratio3.Cells;
                _xlSerie_Stds4.Values = _xlRange_StdBracket_Ratio4.Cells;
                //_xlSerie_Stds5.Values = _xlRange_StdBracket_Ratio5.Cells;
                //_xlSerie_Stds6.Values = _xlRange_StdBracket_Ratio6.Cells;

                string _serieStdName1 = "", _serieStdName2 = "", _serieStdName3 = "", _serieStdName4 = ""; //, _serieStdName5 = "", _serieStdName6 = "";

                string _axeStdXNumFormat = "0.00",
                       _axeStdYNumFormat = "0.00";

                string _axeStdXCaption = "%, Predicted ratio",
                       _axeStdYCaption = "%, Area Ratio";

                if (ExperimentInfo._quanType == "ABSOLUTE")
                {
                    _axeStdXNumFormat = "0.0"; _axeStdYNumFormat = "0.0";

                    _axeStdXCaption = "Predicted ratio"; _axeStdYCaption = "Area Ratio";
                }

                //_serieStdName1 = "E3 (4) 4-28-07";
                //_serieStdName2 = "E3 (4) 5-16-07";
                //_serieStdName3 = "E3 (4) 5-17-07";

                _serieStdName1 = "E33 CSF";
                _serieStdName2 = "E44 CSF";
                _serieStdName3 = "EMix CSF";

                _serieStdName4 = "Synth";
                //_serieStdName5 = "E3 (4) 5-16-07";
                //_serieStdName6 = "E4 5-05-08";



                _xlAxisX_Std.TickLabels.NumberFormat = _axeStdXNumFormat; _xlAxisY_Std.TickLabels.NumberFormat = _axeStdYNumFormat;
                _xlAxisX_Std.AxisTitle.Caption = _axeStdXCaption; _xlAxisY_Std.AxisTitle.Caption = _axeStdYCaption;

                _xlSerie_Stds1.Name = _serieStdName1;
                _xlSerie_Stds2.Name = _serieStdName2;
                _xlSerie_Stds3.Name = _serieStdName3;
                _xlSerie_Stds4.Name = _serieStdName4;
                //_xlSerie_Stds5.Name = _serieStdName5;
                //_xlSerie_Stds6.Name = _serieStdName6;

                Excel.Trendlines _xlTrendlines_Stds1 = (Excel.Trendlines)_xlSerie_Stds1.Trendlines(Type.Missing);
                Excel.Trendlines _xlTrendlines_Stds2 = (Excel.Trendlines)_xlSerie_Stds2.Trendlines(Type.Missing);
                Excel.Trendlines _xlTrendlines_Stds3 = (Excel.Trendlines)_xlSerie_Stds3.Trendlines(Type.Missing);
                Excel.Trendlines _xlTrendlines_Stds4 = (Excel.Trendlines)_xlSerie_Stds4.Trendlines(Type.Missing);
                //Excel.Trendlines _xlTrendlines_Stds5 = (Excel.Trendlines)_xlSerie_Stds5.Trendlines(Type.Missing);
                //Excel.Trendlines _xlTrendlines_Stds6 = (Excel.Trendlines)_xlSerie_Stds6.Trendlines(Type.Missing);


                Excel.XlTrendlineType _XlTrendlineType =
                    Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear;


                _xlTrendlines_Stds1.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, true, true, Type.Missing);
                _xlTrendlines_Stds2.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, true, true, Type.Missing);
                _xlTrendlines_Stds3.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, true, true, Type.Missing);
                //_xlTrendlines_Stds4.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                //    Type.Missing, true, true, Type.Missing);
                //_xlTrendlines_Stds5.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                //    Type.Missing, true, true, Type.Missing);

                //****************

                // };


                //Unknowns
                _xlRange_Unknowns =
                    _wsComponent.get_Range
                    ("A" + (Title + _StdBracketCount + _BlankCount).ToString(), "AA" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknowns.Name = _wsComponent.Name + "_Unknowns";

                _xlRange_Unknown_Area =
                    _wsComponent.get_Range
                    ("D" + (Title + _StdBracketCount + _BlankCount).ToString(), "D" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_Area.Name = _wsComponent.Name + "_Unknown_Area";

                _xlRange_Unknown_ISArea =
                    _wsComponent.get_Range
                    ("E" + (Title + _StdBracketCount + _BlankCount).ToString(), "E" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_ISArea.Name = _wsComponent.Name + "_Unknown_ISArea";

                _xlRange_Unknown_Ratio =
                   _wsComponent.get_Range
                   ("F" + (Title + _StdBracketCount + _BlankCount).ToString(), "F" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_Ratio.Name = _wsComponent.Name + "_Unknown_Ratio";

                string strCalcOrRAtio = "H";
                if (_StdBracketCount == 0) { strCalcOrRAtio = "F"; }
                _xlRange_Unknown_CalcAmount =
                    _wsComponent.get_Range
                    (strCalcOrRAtio + (Title + _StdBracketCount + _BlankCount).ToString(), strCalcOrRAtio + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_CalcAmount.Name = _wsComponent.Name + "_Unknown_CalcAmount";

                _xlRange_Unknown_SampleID =
                    _wsComponent.get_Range
                    ("J" + (Title + _StdBracketCount + _BlankCount).ToString(), "J" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_SampleID.Name = _wsComponent.Name + "_Unknown_SampleID";



                // 2014-01-21

                Excel.Range _xlRange_byStd1, _xlRange_byStd2, _xlRange_byStd3, _xlRange_byStd4, _xlRange_Subject;

                _xlRange_Subject = _xlRange_Unknown_SampleID.get_Offset(0, 1);
                _xlRange_Subject.Name = _wsComponent.Name + "_Subjects";

                _xlRange_byStd1 = _xlRange_Unknown_Ratio.get_Offset(0, 10);
                _xlRange_byStd1.Formula =
                    String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                             _wsComponent.Name + "_Unknown_Ratio",
                             _wsComponent.Name + "_StdBracket_Ratio1",
                             _wsComponent.Name + "_StdBracket_SpecifiedAmount1");
                _xlRange_byStd1.Name = _wsComponent.Name + "_byStd1";

                _xlRange_byStd2 = _xlRange_Unknown_Ratio.get_Offset(0, 11);
                _xlRange_byStd2.Formula =
                    String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                             _wsComponent.Name + "_Unknown_Ratio",
                             _wsComponent.Name + "_StdBracket_Ratio2",
                             _wsComponent.Name + "_StdBracket_SpecifiedAmount2");
                _xlRange_byStd2.Name = _wsComponent.Name + "_byStd2";

                _xlRange_byStd3 = _xlRange_Unknown_Ratio.get_Offset(0, 12);
                _xlRange_byStd3.Formula =
                String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                         _wsComponent.Name + "_Unknown_Ratio",
                         _wsComponent.Name + "_StdBracket_Ratio3",
                         _wsComponent.Name + "_StdBracket_SpecifiedAmount3");
                _xlRange_byStd3.Name = _wsComponent.Name + "_byStd3";

                _xlRange_byStd4 = _xlRange_Unknown_Ratio.get_Offset(0, 13);
                _xlRange_byStd4.Formula =
                String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                         _wsComponent.Name + "_Unknown_Ratio",
                         _wsComponent.Name + "_StdBracket_Ratio4",
                         _wsComponent.Name + "_StdBracket_SpecifiedAmount4");
                _xlRange_byStd4.Name = _wsComponent.Name + "_byStd4";
                /*
                                _xlRange_byStd5 = _xlRange_StdBracket_Ratio.get_Offset(0, 14);
                                _xlRange_byStd5.Formula =
                                String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                                         _wsComponent.Name + "_StdBracket_Ratio",
                                         _wsComponent.Name + "_StdBracket_Ratio5",
                                         _wsComponent.Name + "_StdBracket_SpecifiedAmount5");
                                _xlRange_byStd5.Name = _wsComponent.Name + "_byStd5";
                */
                // 2013-12-10
                /*
                                Excel.Range _xlRange_byStd1, _xlRange_byStd2, _xlRange_byStd3, _xlRange_Subject;

                                _xlRange_Subject = _xlRange_Unknown_SampleID.get_Offset(0, 1);
                                _xlRange_Subject.Name = _wsComponent.Name + "_Subjects";

                                _xlRange_byStd1 = _xlRange_StdBracket_Ratio.get_Offset(0, 10);
                                _xlRange_byStd1.Formula =
                                    String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                                             _wsComponent.Name + "_Unknown_Ratio",
                                             _wsComponent.Name + "_StdBracket_Ratio1",
                                             _wsComponent.Name + "_StdBracket_SpecifiedAmount1" ) ;
                                _xlRange_byStd1.Name = _wsComponent.Name + "_byStd1";

                                _xlRange_byStd2 = _xlRange_StdBracket_Ratio.get_Offset(0, 11);
                                _xlRange_byStd2.Formula =
                                    String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                                             _wsComponent.Name + "_Unknown_Ratio",
                                             _wsComponent.Name + "_StdBracket_Ratio2",
                                             _wsComponent.Name + "_StdBracket_SpecifiedAmount2");
                                _xlRange_byStd2.Name = _wsComponent.Name + "_byStd2";

                                _xlRange_byStd3 = _xlRange_StdBracket_Ratio.get_Offset(0, 12);
                                _xlRange_byStd3.Formula =
                                String.Format("=({0}-INTERCEPT({1}, {2}))/SLOPE({1}, {2})",
                                         _wsComponent.Name + "_Unknown_Ratio",
                                         _wsComponent.Name + "_StdBracket_Ratio3",
                                         _wsComponent.Name + "_StdBracket_SpecifiedAmount3");
                                _xlRange_byStd3.Name = _wsComponent.Name + "_byStd3";
                */

                //*************************Format Ranges**************************
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;


                _xlRange_Unknowns.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                _xlRange_Unknowns.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                // _xlRange_Unknowns.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color. .Chocolate);

                //******Sort Unknowns
                _xlRange_Unknowns.Sort("Column N",
                    Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    "Column J", Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                    Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);

                //*************2011-07-19**************************************

                //_xlRange_Unknown_ISArea.FormatConditions.AddDatabar();

                Excel.ColorScale cfColorScale = (Excel.ColorScale)(_xlRange_Unknown_ISArea.FormatConditions.AddColorScale(3));

                // Set the minimum threshold to red (0x000000FF) and maximum threshold
                // to blue (0x00FF0000).
                cfColorScale.ColorScaleCriteria[1].FormatColor.Color = 7039480;  //0x000000FF;  
                cfColorScale.ColorScaleCriteria[2].FormatColor.Color = 8711167;  // 0x00FF0000;
                cfColorScale.ColorScaleCriteria[3].FormatColor.Color = 13011546; // 0x00FF0000;

                //*************************Charts*********************************
                //
                // Area Ratio
                Excel.ChartObject _xlChartTimeCorse = (Excel.ChartObject)_xlCharts.Add(
                                   0, _graphOffset, 950, 250);
                Excel.Chart _xlChartTimeCorsePage = _xlChartTimeCorse.Chart;

                _xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlColumnClustered;
                //_xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlXYScatter;

                _xlChartTimeCorsePage.HasTitle = true;
                _xlChartTimeCorsePage.ChartTitle.Font.Size = 11;

                _xlChartTimeCorsePage.ChartTitle.Caption = _componentName + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_TP = (Excel.Axes)_xlChartTimeCorsePage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_TP.HasMajorGridlines = true;

                //_xlAxisX_TP.HasTitle = true;
                //_xlAxisY_TP.HasTitle = true;

                _xlChartTimeCorsePage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartTimeCorsePage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_TimeCourse = (Excel.SeriesCollection)_xlChartTimeCorsePage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_TTR = (Excel.Series)_xlSeriesColl_TimeCourse.NewSeries();


                _xlSeries_TTR.XValues = _xlRange_Subject.Cells;
                _xlSeries_TTR.Values = _xlRange_Unknown_Ratio;


                //_xlSeries_TTR.XValues = _xlRange_StdBracket_Ratio.get_Offset(0, 5).Cells;
                //_xlSeries_TTR.Values = _xlRange_StdBracket_Ratio;



                string _serieCurveName = "H:L Abeta";

                string _axeCurveXNumFormat = "0",
                       _axeCurveYNumFormat = "0.0%";

                
                if (ExperimentInfo._quanType == "ABSOLUTE")
                {
                    _axeCurveYNumFormat = "0.0";
                    _serieCurveName = "Abeta level";
                }

                
                _xlAxisX_TP.TickLabels.NumberFormat = _axeCurveXNumFormat;
                _xlAxisY_TP.TickLabels.NumberFormat = _axeCurveYNumFormat;
                               

                _xlSeries_TTR.Name = _serieCurveName;
                _xlSeries_TTR.Name = _componentName;

                
                // Normalized

                
                Excel.ChartObject _xlChartTimeCorseNorm = (Excel.ChartObject)_xlCharts.Add(
                                   0, _graphOffset + _xlChartTimeCorse.Height, 950, 250);
                Excel.Chart _xlChartTimeCorseNormPage = _xlChartTimeCorseNorm.Chart;

                _xlChartTimeCorseNormPage.ChartType = Excel.XlChartType.xlColumnClustered;
                
                _xlChartTimeCorseNormPage.HasTitle = true;
                _xlChartTimeCorseNormPage.ChartTitle.Font.Size = 11;

                _xlChartTimeCorseNormPage.ChartTitle.Caption = _componentName + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_TP_Norm = (Excel.Axes)_xlChartTimeCorseNormPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_TP_Norm = _xlAxes_TP_Norm.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_TP_Norm = _xlAxes_TP_Norm.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_TP_Norm.HasMajorGridlines = true;

                _xlAxisX_TP_Norm.TickLabels.NumberFormat = "0";
                _xlAxisY_TP_Norm.TickLabels.NumberFormat = "0.0E+0";

                _xlChartTimeCorseNormPage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartTimeCorseNormPage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_TimeCourseNorm = (Excel.SeriesCollection)_xlChartTimeCorseNormPage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_CalcAm_byStd1 = (Excel.Series)_xlSeriesColl_TimeCourseNorm.NewSeries();

                _xlSeries_CalcAm_byStd1.XValues = _xlRange_Subject.Cells;
                _xlSeries_CalcAm_byStd1.Values = _xlRange_byStd1.Cells;
                _xlSeries_CalcAm_byStd1.Name = "by E33 CSF";


                Excel.Series _xlSeries_CalcAm_byStd2 = (Excel.Series)_xlSeriesColl_TimeCourseNorm.NewSeries();
                _xlSeries_CalcAm_byStd2.XValues = _xlRange_Subject.Cells;
                _xlSeries_CalcAm_byStd2.Values = _xlRange_byStd2.Cells;
                _xlSeries_CalcAm_byStd2.Name = "by E44 CSF";

                Excel.Series _xlSeries_CalcAm_byStd3 = (Excel.Series)_xlSeriesColl_TimeCourseNorm.NewSeries();
                _xlSeries_CalcAm_byStd3.XValues = _xlRange_Subject.Cells;
                _xlSeries_CalcAm_byStd3.Values = _xlRange_byStd3.Cells;
                _xlSeries_CalcAm_byStd3.Name = "by Mix CSF";

                Excel.Series _xlSeries_CalcAm_byStd4 = (Excel.Series)_xlSeriesColl_TimeCourseNorm.NewSeries();
                _xlSeries_CalcAm_byStd4.XValues = _xlRange_Subject.Cells;
                _xlSeries_CalcAm_byStd4.Values = _xlRange_byStd4.Cells;
                _xlSeries_CalcAm_byStd4.Name = "by Syn pep";

                _xlRange_byStd1.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                _xlRange_byStd2.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                _xlRange_byStd3.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);


                Excel.ChartObjects _xlChartsSummary = (Excel.ChartObjects)_wsSummary.ChartObjects(Type.Missing);
                Excel.ChartObject _xlChartCurveSummary;
                Excel.Chart _xlChartCurveSummaryPage;

                if (_xlChartsSummary.Count < 1)
                {
                    _xlChartCurveSummary = (Excel.ChartObject)_xlChartsSummary.Add(
                                   10, 10, 850, 500);
                    _xlChartCurveSummaryPage = _xlChartCurveSummary.Chart;

                    _xlChartCurveSummaryPage.ChartType = Excel.XlChartType.xlXYScatterLines;

                    _xlChartCurveSummaryPage.HasTitle = true;
                    _xlChartCurveSummaryPage.ChartTitle.Font.Size = 11;

                    Excel.Axes _xlAxes_CurveSummary = (Excel.Axes)_xlChartCurveSummaryPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisX_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisY_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                    _xlAxisX_CurveSummary.HasMajorGridlines = true;
                    _xlAxisY_CurveSummary.HasMajorGridlines = true;

                    _xlAxisX_CurveSummary.Format.Line.Weight = 3.0F;
                    _xlAxisY_CurveSummary.Format.Line.Weight = 3.0F;


                    _xlAxisX_CurveSummary.TickLabels.NumberFormat = _axeCurveXNumFormat;
                    _xlAxisY_CurveSummary.TickLabels.NumberFormat = _axeCurveYNumFormat;

                    _xlChartCurveSummaryPage.ChartTitle.Caption = "H:L Abeta ";

                    if (ExperimentInfo._quanType == "ABSOLUTE")
                    {
                        _xlChartCurveSummaryPage.ChartTitle.Caption = "ApoE levels in ";
                        try
                        {
                            // _xlAxisY_CurveSummary.ScaleType = Excel.XlScaleType.xlScaleLogarithmic;
                            // _xlAxisY_CurveSummary.LogBase = 10;
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

                //_xlSeries_CurveTTR.XValues = _xlRange_StdBracket_Ratio.get_Offset(0, 5).Cells;

                _xlSeries_CurveTTR.XValues = _xlRange_Unknown_SampleID.get_Offset(0, 1).Cells; ;

                //_xlSeries_CurveTTR.Values = _xlRange_Unknown_CalcAmount;

                _xlSeries_CurveTTR.Values = _xlRange_Unknown_Ratio;

                //_xlSeries_CurveTTR.Values = _xlRange_StdBracket_Ratio;

                _xlSeries_CurveTTR.Name = (string)_xlRange_Compound.Cells.Value2;


                toolStripProgressBarMain.PerformStep();

            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = false;

            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();

            _XXQN.Close();

            object misValue = System.Reflection.Missing.Value;

            try
            {
                _XXQN.Close();

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
                    ((Excel._Worksheet)_wsSummary).Activate();

                }
                else
                {
                    _xlApp.Quit();
                }
            }
            
        }



        //******************************************************* 2014-04-21 for REL ApoE CSF-curve + Media
        private void ExportAPOE_REL()
        {

            List<string> _strExportSQL = new List<string>();
            /*05-07-2010*/
            string _strAnalyte = null;
            double _Area = 0, _ISTD_Area = 0, _Area_Ratio = 0, _CalcAmount = 0;
            decimal dcm_Area_Ratio = 0;
            string _strSpecAmount = null, _strRT = null, _strISTD_RT = null;
            string _response = null;

            bool _ifRel;

            /*bool IsPeakFound = false;*/

            /*05-07-2010*/
            if (textBox_XQNfile.Text == "")
            {
                MessageBox.Show("Select Input XQN file, please");
                return;
            }


            _XXQN = new XCALIBURFILESLib.XXQNClass();

            try
            {
                _XXQN.Open(textBox_XQNfile.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("File: " + textBox_XQNfile.Text + " is invalid", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //MessageBox.Show(_exception.Message, _exception.GetBaseException().Message);
                return;
            }
            finally
            {

            }

            XCALIBURFILESLib.IXHeader _IXHeader = (XCALIBURFILESLib.IXHeader)_XXQN.Header;

            ExperimentInfo._ProcessedBy = _IXHeader.ChangedId;

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = _IXHeader.ChangedLogon;
            }

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = "unknown";
            }


            
            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = true;
            toolStripStatusLabelMain.Visible = true;



            _xlApp = new Excel.ApplicationClass();
            _xlApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;



            Excel.Workbook xlWB;

            try
            {

                xlWB =
                    _xlApp.Workbooks.Add(
                    //@"c:\Project\XQNReader\XQNReader\LOAD_TSQ_Template_Summary.xlsx"
                Application.StartupPath + Path.DirectorySeparatorChar +
                @"\Templates\ApoE_TSQ_REL_Template_Summary.xlsx"
                );

            }

            catch (Exception)
            {
                MessageBox.Show("File: " + Application.StartupPath + Path.DirectorySeparatorChar +
                @"\Templates\ApoE_TSQ_REL_Template_Summary.xlsx", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //MessageBox.Show(_exception.Message, _exception.GetBaseException().Message);
                return;
            }
            finally
            {

            }

            //_xlApp.Visible = true;   /// 222222222222222


            Excel.Worksheet _wsSummary = (Excel.Worksheet)xlWB.ActiveSheet;

            Excel.Worksheet _wsCurrent = null;


            byte Title = 5;
            byte _UnknownCount = 0, _StdBracketCount = 0, _BlankCount = 0, _TotalCount = 0;




            XCALIBURFILESLib.XBracketGroups _XBracketGroups
                = (XCALIBURFILESLib.XBracketGroups)_XXQN.BracketGroups;

            XCALIBURFILESLib.XBracketGroup _XBracketGroup
                = (XCALIBURFILESLib.XBracketGroup)_XBracketGroups.Item(1);

            XCALIBURFILESLib.XQuanResults _XQuanResults
                = (XCALIBURFILESLib.XQuanResults)_XBracketGroup.QuanResults;




            var _XQuanResultsSorted = from XCALIBURFILESLib.XQuanResult t in _XQuanResults
                                      orderby ((XCALIBURFILESLib.XSample)t.Sample).SampleId descending
                                      select t;


            //********_XQuanResults*******************************

            foreach (XCALIBURFILESLib.XQuanResult _XQuanResult in _XQuanResultsSorted /*_XQuanResults*/)
            {
                toolStripProgressBarMain.Maximum = _XQuanResults.Count;

                XCALIBURFILESLib.XSample _Sample
                    = (XCALIBURFILESLib.XSample)_XQuanResult.Sample;



                //MessageBox.Show(_Sample.FileName + " - " + _Sample.RowNumber.ToString() );

                _TotalCount++;
                switch (_Sample.SampleType)
                {
                    case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                        _UnknownCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                        _StdBracketCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleBlank:
                        _BlankCount++;
                        break;
                    default:
                        break;
                }

                XCALIBURFILESLib.XQuanComponents _XQuanSelectedComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Selected; /*  Method; */

                XCALIBURFILESLib.XQuanComponents _XQuanMethodComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Method;   /*  Method; */


                ////////////////////////////////////////
                short _comNo = 1;
                foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponent in _XQuanSelectedComponents)
                {
                    XCALIBURFILESLib.XComponent _XComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanComponent.Component;

                    XCALIBURFILESLib.XQuanComponent _XQuanMethodComponent = (XCALIBURFILESLib.XQuanComponent)
                    _XQuanMethodComponents.Item(_comNo);
                    _comNo++;

                    XCALIBURFILESLib.XComponent _XMethodComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanMethodComponent.Component;

                    /*
                    XCALIBURFILESLib.XComponent _XComponentUser =
                        (XCALIBURFILESLib.XComponent)_XQuanComponent.Component;
                    */


                    XCALIBURFILESLib.IXDetection2 _IXDetection =
                        (XCALIBURFILESLib.IXDetection2)_XComponent.Detection;


                    XCALIBURFILESLib.XQuanCalibration _XQuanCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XComponent.Calibration;

                    XCALIBURFILESLib.XQuanCalibration _XQuanMethodCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XMethodComponent.Calibration;

                    string _quanTypeKey = "LALALA";




                    if (ExperimentInfo._quanType == "RELATIVE")
                    {
                        _quanTypeKey = "R0:N14";
                    }
                    if (ExperimentInfo._quanType == "ABSOLUTE")
                    {
                        _quanTypeKey = "L6:C13";
                    }


                    if (_XQuanCalibration.ComponentType != XCALIBURFILESLib.XComponentType.XISTD
                        && !_XComponent.Name.Contains(_quanTypeKey))
                    {


                        _strAnalyte = _XComponent.Name.Replace("_", " ");


                        bool _SheetExist = false; _ifRel = false;

                        string _ComponentName = _XComponent.Name.Replace(":", "_").Replace("-", "_").Replace("(", "_").Replace(")", "_");


                        if (comboBox_Project.Text == "ApoE REL KW")
                        {
                            switch (_strAnalyte)
                            {
                                case "13C-LAVY(E34)":
                                    _strAnalyte = "LAVYQAGAR L6:C13";
                                    break;

                                case "13C LGADMEDVcGR (E32)":
                                    _strAnalyte = "LGADMEDVCGR L6:C13";
                                    break;
                                case "13C-CLAVY(E2)":
                                    _strAnalyte = "CLAVYQAGAR L6:C13";
                                    break;

                                case "13C-LGADMEDVR (E4)":
                                    _strAnalyte = "LGADMEDVR L6:C13";
                                    break;

                                default:
                                    _strAnalyte = "Don't freaking know";
                                    break;
                            }
 
                        }


                        if (_ComponentName.Contains("L6_C13") || _ComponentName.Contains("13C")) { _ifRel = true; }

                        //MessageBox.Show("Compound - "+ _ComponentName +  "_ifRel = " + _ifRel.ToString() );

                        foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
                        {

                            if (_wsComponent.Name == _ComponentName
                                )
                            {
                                _wsCurrent = _wsComponent;
                                _SheetExist = true;
                            }
                        }



                        if (!_SheetExist)
                        {
                            _wsCurrent = (Excel.Worksheet)xlWB.Sheets.Add
                          (Type.Missing, Type.Missing, Type.Missing,
                          Application.StartupPath + Path.DirectorySeparatorChar +
                            @"\Templates\ApoE_TSQ_REL_Template.xlsx"
                          );

                            _wsCurrent.Name = _ComponentName;



                            _wsCurrent.Cells[2, 1] = _XComponent.Name.Replace("_", " ");

                            _wsCurrent.Cells[2, 3] = CurveTypeToStr(_XQuanMethodCalibration.CurveType);

                            _wsCurrent.Cells[2, 4] = weightingTypeToStr(_XQuanMethodCalibration.Weighting);

                            _wsCurrent.Cells[2, 6] = OriginTypeToStr(_XQuanMethodCalibration.OriginType);

                            _wsCurrent.Cells[2, 8] = _XQuanMethodCalibration.FullEquation;

                        }



                        _wsCurrent.Cells[_Sample.RowNumber + Title, 1] = _Sample.FileName;
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 2] = SampleTypeToStr(_Sample.SampleType);
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 3] = IntegTypeToStr(_XQuanComponent.Selected_Type);


                        XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeak;

                        bool _IsPeakFoundANALYTE = false;
                        bool _IsPeakFoundISTD = false;
                        int _shift = 2;

                        if (_XQuanComponent.DoesDetectedQuanPeakExist > 0)
                        {


                            // MessageBox.Show("*");


                            _IsPeakFoundANALYTE = true;
                            _XDetectedQuanPeak =
                                (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponent.DetectedQuanPeak;

                            _wsCurrent.Cells[_Sample.RowNumber + Title, 4] = _XDetectedQuanPeak.Area;
                            _Area = _XDetectedQuanPeak.Area;



                            // _wsCurrent.Cells[_Sample.RowNumber + Title, 8] =
                            //     _ifRel == true ? _XDetectedQuanPeak.CalcAmount : 0 ;

                            _CalcAmount = _XDetectedQuanPeak.CalcAmount;


                            _wsCurrent.Cells[_Sample.RowNumber + Title, 15 + _shift] = _XDetectedQuanPeak.LeftRT;
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 16 + _shift] = _XDetectedQuanPeak.RightRT;


                            _wsCurrent.Cells[_Sample.RowNumber + Title, 14 + _shift] = String.Format("=L{0}-M{0}", (_Sample.RowNumber + Title).ToString());
                            // _wsCurrent.Cells[_Sample.RowNumber + Title, 22 + _shift] = String.Format("=U{0}-T{0}", (_Sample.RowNumber + Title).ToString());
                            // _wsCurrent.Cells[_Sample.RowNumber + Title, 25 + _shift] = String.Format("=X{0}-W{0}", (_Sample.RowNumber + Title).ToString());



                            if (_XDetectedQuanPeak.ISTDArea != 0)
                            {
                                _IsPeakFoundISTD = true;
                                _wsCurrent.Cells[_Sample.RowNumber + Title, 5] = _XDetectedQuanPeak.ISTDArea;
                                _wsCurrent.Cells[_Sample.RowNumber + Title, 6] =
                                _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;


                                _wsCurrent.Cells[_Sample.RowNumber + Title, 8] =
                                    _ifRel == true ? _XDetectedQuanPeak.CalcAmount : _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;



                                _ISTD_Area = _XDetectedQuanPeak.ISTDArea;
                                _Area_Ratio = _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;
                                dcm_Area_Ratio = (decimal)_XDetectedQuanPeak.Area / (decimal)_XDetectedQuanPeak.ISTDArea;
                            }
                            else
                            {
                                _wsCurrent.Cells[_Sample.RowNumber + Title, 5] = "N/F";
                            }


                            _response = "NULL"; _strRT = "NULL"; _strISTD_RT = "NULL";

                            if (_XDetectedQuanPeak.IsResponseOK > 0) { _response = "Ok"; }
                            if (_XDetectedQuanPeak.IsResponseHigh > 0) { _response = "High"; }
                            if (_XDetectedQuanPeak.IsResponseLow > 0) { _response = "Low"; }

                            _wsCurrent.Cells[_Sample.RowNumber + Title, 9] = _response;
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 12] = _XDetectedQuanPeak.ApexRT;

                            _strRT = _XDetectedQuanPeak.ApexRT.ToString("F2");

                        }
                        else
                        {
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 9] = "Not Found";

                        }

                        string _sample_ID = "N/A";
                        double _specAmount = 0;
                        switch (_Sample.SampleType)
                        {
                            case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                                _sample_ID = _Sample.Level;
                                XCALIBURFILESLib.XCalibrationCurveData _XCalibrationCurveData =
                                (XCALIBURFILESLib.XCalibrationCurveData)_XQuanComponent.CalibrationCurveData;
                                XCALIBURFILESLib.XReplicateRows _XReplicateRows =
                                (XCALIBURFILESLib.XReplicateRows)_XCalibrationCurveData.ReplicateRows;

                                foreach (XCALIBURFILESLib.XReplicateRow _XReplicateRow in _XReplicateRows)
                                {
                                    if (_XReplicateRow.LevelName == _Sample.Level)
                                    {
                                        _specAmount = _XReplicateRow.Amount;
                                        _strSpecAmount = "'" + _XReplicateRow.Amount.ToString("F5") + "'";
                                    }
                                }

                                _wsCurrent.Cells[_Sample.RowNumber + Title, 11] = _Sample.FileName.Split('_')[1];

                                break;

                            case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                                _sample_ID = _Sample.SampleId;
                                _specAmount = -1;
                                _strSpecAmount = "NULL";

                                _wsCurrent.Cells[_Sample.RowNumber + Title, 11] = _Sample.FileName.Split('_')[1];


                                break;

                            case XCALIBURFILESLib.XSampleTypes.XSampleBlank:
                                _sample_ID = _Sample.SampleId;
                                _specAmount = -1;
                                _strSpecAmount = "NULL";
                                break;
                            default:
                                break;
                        }


                        bool _IsSampleIDDefinded = false;
                        int _intTP = -1;

                        if (_sample_ID == "")
                        {
                            _sample_ID = _Sample.FileName.Substring(_Sample.FileName.LastIndexOf('_') + 1);
                        };

                        try
                        {
                            _intTP = Convert.ToInt16(_sample_ID);
                        }
                        catch (System.FormatException)
                        {
                            if (checkBox_ShowError.Checked) 
                            {
                                MessageBox.Show(_sample_ID + " - cannot be converted to time point");
                            }
                           
                        };

                        if (_intTP < 0)
                        {
                            _sample_ID = "";
                            _IsSampleIDDefinded = false;
                        }
                        else
                        {
                            _sample_ID = _intTP.ToString();
                            _IsSampleIDDefinded = true;
                        }


                        if (_specAmount != -1)
                        {
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 7] = _specAmount;
                        }
                        else
                        {
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 7] = "";
                        }
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 10] = _sample_ID;




                        foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponentISTD in _XQuanSelectedComponents)
                        {
                            XCALIBURFILESLib.XComponent _XISTD =
                            (XCALIBURFILESLib.XComponent)_XQuanComponentISTD.Component;

                            if (_XISTD.Name == _XQuanCalibration.ISTD)
                            {
                                if (_XQuanComponentISTD.DoesDetectedQuanPeakExist > 0)
                                {
                                    XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeakISTD =
                                        (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponentISTD.DetectedQuanPeak;
                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 13] = _XDetectedQuanPeakISTD.ApexRT;

                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 18 + _shift] = _XDetectedQuanPeakISTD.LeftRT;
                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 19 + _shift] = _XDetectedQuanPeakISTD.RightRT;



                                    _strISTD_RT = _XDetectedQuanPeakISTD.ApexRT.ToString("F2");
                                }
                                else
                                {
                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 13] = "N/F";
                                }

                            }

                        }

                        _wsCurrent.Cells[_Sample.RowNumber + Title, 14] = _Sample.SampleId;
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 15] = _Sample.SampleName;


                        // if (GetSlopeOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType)) == -1)
                        // {
                        //     return;
                        // }

                        string _subjectNum = "";

                        if (checkBox_if_multi_subject.Checked && (int)_Sample.SampleType == 0)
                        {
                            _subjectNum = _Sample.FileName.Split('_')[1];
                        }
                        else
                        {
                            if (ExperimentInfo._Subject.Contains("-"))
                            {
                                _subjectNum = ExperimentInfo._Subject.Split('-')[1];
                            }
                            else if (ExperimentInfo._Subject == "") { _subjectNum = ExperimentInfo._fileName.Substring(ExperimentInfo._fileName.Length - 6); }
                                else {_subjectNum = ExperimentInfo._Subject;}

                        }

                        //MessageBox.Show(ExperimentInfo._SampleQuantitatedBy_ID.ToString());

                        if (_IsPeakFoundANALYTE && _IsPeakFoundISTD && _IsSampleIDDefinded && checkBox_ExportIntoDB.Checked)
                        {
                            _strExportSQL.Add("INSERT INTO TMP_TSQ_IMPORT ( ANYLYTE_NAME, FILENAME_DESC, SAMPLE_TYPE_ID, INTEG_TYPE_ID, AREA, ISTD_AREA, AREA_RATIO, " +
                                                                          "SPEC_AMOUNT, CALC_AMOUNT, LEVEL_TP, PEAK_SATUS, RT, ISTD_RT, QUAN_TYPE_ID, FLUID_TYPE_ID, " +
                                                                          "SUBJECT_NUM, PROCESS_BY_ID, ASSAY_DATE, ASSAY_TYPE_ID, ANTIBODY_ID, ENZYME_ID, DONE_BY, " +
                                                                          "SAMPLE_PROCESS_BY, EXPER_FILE, CURVE, WEIGHTING, ORIGIN, EQUATION, R, R_SQR, NUM_SLOPE, " +
                                                                          "NUM_INTERSECT, ASSAY_DESIGN_ID, PROJECT_ID )");
                            _strExportSQL.Add(String.Format("VALUES ('{0}', '{1}', {2}, '{3}', {4:F0}, {5:F0}, {6:F5}, {7:F}, {8:F5}, {9}, '{10}', {11}, {12}, {13}, {14}, '{15}', {16}, '{17}', {18}, {19}, {20}, {21}, {22}, '{23}'," +
                                " '{24}', '{25}', '{26}', '{27}', {28:F5}, {29:F5}, {30:F5}, {31:F5}, {32}, {33});",
                                   _strAnalyte,                                 // 0
                                   _Sample.FileName,                            // 1
                                   (int)_Sample.SampleType,                     // 2
                                   (int)_XQuanComponent.Selected_Type,          // 3
                                   _Area,                                       // 4
                                   _ISTD_Area,                                  // 5
                                   _Area_Ratio,                                 // 6
                                   _strSpecAmount,                              // 7 

                                   _ifRel == true ? _CalcAmount : _Area_Ratio,
                                //_CalcAmount,                                 // 8

                                   _sample_ID,                                  // 9
                                   _response,                                   // 10
                                   _strRT,                                      // 11 
                                   _strISTD_RT,                                 // 12

                                   _ifRel == true ? 1 : 2,

                                  // QuanTypeToDBInt(ExperimentInfo._quanType),

                                   FluidTypeToDBInt(ExperimentInfo._matrix),    // 14
                                   _subjectNum,       // 15
                                   ExperimentInfo._SampleQuantitatedBy_ID,      // 16
                                   ExperimentInfo._Date,                        // 17
                                   11,                                          // 18      { ASSAY_TYPE_ID   11 - "ApoE IP LC/MS (C-term)"    
                                   comboBox_Abody.SelectedValue,                // 19      { 3, // A-BODY  3 - HJ5.x
                                   comboBox_Enzyme.SelectedValue,               // 20      { 2, //Enzyme 2 - LysN
                                   ExperimentInfo._SampleRanBy_ID,              // 21      { Done by 1 - ovodvo
                                   ExperimentInfo._SamplePreparedBy_ID,         // 22      { Sample process by 1 - ovodvo
                                   ExperimentInfo._expName,                     // 23

                                   _ifRel == true ? CurveTypeToStr(_XQuanMethodCalibration.CurveType) : "",
                                   _ifRel == true ? weightingTypeToStr(_XQuanMethodCalibration.Weighting) : "",
                                   _ifRel == true ? OriginTypeToStr(_XQuanMethodCalibration.OriginType) : "",
                                   _ifRel == true ? GetEquaOfEquation(_XQuanMethodCalibration.FullEquation) : "",
                                   _ifRel == true ? GetROfEquation(_XQuanMethodCalibration.FullEquation).ToString() : "NULL",
                                   _ifRel == true ? GetR_SQROfEquation(_XQuanMethodCalibration.FullEquation).ToString() : "NULL",
                                   _ifRel == true ? GetSlopeOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType)).ToString() : "NULL",
                                   _ifRel == true ? GetInterOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType)).ToString() : "NULL",

                                   1, // time course

                                   comboBox_Project.SelectedValue

                                   ));

                        }
                    }
                }

                toolStripProgressBarMain.PerformStep();
                toolStripStatusLabelMain.Text = "Processing ..." + _Sample.FileName;
                toolStripStatusLabelMain.PerformClick();

            }

            //_xlApp.Visible = true;
            toolStripProgressBarMain.Visible = false;
            toolStripProgressBarMain.Value = 0;
            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();
            toolStripProgressBarMain.Visible = true;

            toolStripProgressBarMain.Maximum = xlWB.Worksheets.Count - 1;


            if (checkBox_ExportIntoDB.Checked)
            {

                string _AppPath = Application.StartupPath + Path.DirectorySeparatorChar;

                string[] _sql_P1 = File.ReadAllLines(_AppPath + "sql_TSQImport_P1.sql");
                string[] _sql_P2 = File.ReadAllLines(_AppPath + "sql_TSQImport_P2.sql");
                string[] _sql_P3 = File.ReadAllLines(_AppPath + "sql_TSQImport_P3.sql");

                _strExportSQL.Insert(0, " ");
                _strExportSQL.InsertRange(0, _sql_P1.ToList<string>());
                _strExportSQL.AddRange(_sql_P3.ToList<string>());
                //_strExportSQL.AddRange(_sql_P2.ToList<string>());


                File.WriteAllLines(Path.ChangeExtension(textBox_Excelfile.Text, "sql"), _strExportSQL.ToArray());
            }

            foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
            {
                _wsComponent.Name = "_" + _wsComponent.Name;

                toolStripStatusLabelMain.Text = "Formatting ..." + _wsComponent.Name;
                toolStripStatusLabelMain.PerformClick();

                if (_wsComponent.Name.Contains("Summary")) { continue; }


                _ifRel = false;

                if (_wsComponent.Name.Contains("L6_C13")) { _ifRel = true; }


                Excel.ChartObjects _xlCharts = (Excel.ChartObjects)_wsComponent.ChartObjects(Type.Missing);
                string _componentName = _wsComponent.get_Range("A2", "A2").Value2.ToString();

                Excel.Range _xlRange_Compound,
                            _xlRange_StdBracket_Area,
                            _xlRange_StdBracket_ISArea,
                            _xlRange_StdBracket_SpecifiedAmount,
                            _xlRange_StdBracket_Ratio,
                            _xlRange_Unknown_Area,
                            _xlRange_Unknown_ISArea,
                            _xlRange_Unknown_Ratio,
                            _xlRange_Unknown_CalcAmount,
                            _xlRange_Unknown_SampleID,
                            _xlRange_STD,
                            _xlRange_Unknowns,
                            _xlRange_Unknown_Subject,
                            _xlRange_All;

/*
                           _xlRange_StdBracket_SpecifiedAmount1,
                           _xlRange_StdBracket_SpecifiedAmount2,
                           _xlRange_StdBracket_SpecifiedAmount3,
                           _xlRange_StdBracket_SpecifiedAmount4,
                           _xlRange_StdBracket_SpecifiedAmount5,
                           _xlRange_StdBracket_SpecifiedAmount6,
                           _xlRange_StdBracket_Ratio1,
                           _xlRange_StdBracket_Ratio2,
                           _xlRange_StdBracket_Ratio3,
                           _xlRange_StdBracket_Ratio4,
                           _xlRange_StdBracket_Ratio5,
                           _xlRange_StdBracket_Ratio6;
*/

                _xlRange_Compound = _wsComponent.get_Range("A2", "A2");

                //All
                _xlRange_All =
                    _wsComponent.get_Range
                    ("A" + Title.ToString(), "V" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_All.Name = _wsComponent.Name + "_All";


                double _rowHeight = (double)_wsCurrent.get_Range("A1", "A1").RowHeight;

                double _graphOffset = (double)_xlRange_All.Top + (double)_xlRange_All.Height +

                      (_rowHeight * 3);



                //Std
                // if (_StdBracketCount > 0)
                // {
                _xlRange_STD =
                    _wsComponent.get_Range
                    ("A" + Title.ToString(), "V" + (Title - 1 + _StdBracketCount).ToString());
                _xlRange_STD.Name = _wsComponent.Name + "_STD";


                _xlRange_StdBracket_Ratio =
                    _wsComponent.get_Range
                    ("F" + Title.ToString(), "F" + (Title - 1 + _StdBracketCount).ToString());
                _xlRange_StdBracket_Ratio.Name = _wsComponent.Name + "_StdBracket_Ratio";

                _xlRange_StdBracket_Area =
                    _wsComponent.get_Range
                    ("D" + Title.ToString(), "D" + (Title - 1 + _StdBracketCount).ToString());
                _xlRange_StdBracket_Area.Name = _wsComponent.Name + "_StdBracket_Area";

                _xlRange_StdBracket_ISArea =
                    _wsComponent.get_Range
                    ("E" + Title.ToString(), "E" + (Title - 1 + _StdBracketCount).ToString());
                _xlRange_StdBracket_ISArea.Name = _wsComponent.Name + "_StdBracket_ISArea";

                _xlRange_StdBracket_SpecifiedAmount =
                   _wsComponent.get_Range
                   ("G" + Title.ToString(), "G" + (Title - 1 + _StdBracketCount).ToString());
                _xlRange_StdBracket_SpecifiedAmount.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount";


                _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                _xlRange_STD.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);

                //Unknowns
                /*   _xlRange_Unknowns = _wsComponent.get_Range
                       ("A" + (Title - 1 + _StdBracketCount).ToString(), "O" + (Title - 1 + _StdBracketCount + _UnknownCount).ToString());
                   _xlRange_STD.Name = _wsComponent.Name + "_Unknown"; */

                // v

                //*********Sort Range************
                _xlRange_STD.Sort(_xlRange_STD.Columns[14, Type.Missing], Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Microsoft.Office.Interop.Excel.XlYesNoGuess.xlGuess,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                    Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);



                // StdCurve;

                if (_ifRel)
                {

                    Excel.ChartObject _xlChartStdCurve = (Excel.ChartObject)_xlCharts.Add(
                                       770, _graphOffset, 470, 250);
                    Excel.Chart _xlChartStdCurvePage = _xlChartStdCurve.Chart;
                    _xlChartStdCurvePage.ChartType = Excel.XlChartType.xlXYScatter;

                    _xlChartStdCurvePage.HasTitle = true;
                    _xlChartStdCurvePage.ChartTitle.Font.Size = 11;

                    string _chartTitleStdCurve = String.Format
                        ("Calibration curve {0} {1} {2} {3} ({4})",
                        _componentName,
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

                    _xlAxisX_Std.MinimumScale = 0; _xlAxisX_Std.MaximumScale = 0.12;
                    _xlAxisY_Std.MinimumScale = 0; //_xlAxisY_TP.MaximumScale = 15;

                    _xlChartStdCurvePage.Legend.IncludeInLayout = true;


                    Excel.SeriesCollection _xlSerColl_StdCurve = (Excel.SeriesCollection)_xlChartStdCurvePage.SeriesCollection(Type.Missing);

                    Excel.Series _xlSerie_Stds1 = (Excel.Series)_xlSerColl_StdCurve.NewSeries();

                    _xlSerie_Stds1.XValues = _xlRange_StdBracket_SpecifiedAmount.Cells;
                    _xlSerie_Stds1.Values = _xlRange_StdBracket_Ratio.Cells;

                    string _serieStdName1 = "";

                    string _axeStdXNumFormat = "0.00",
                           _axeStdYNumFormat = "0.00";

                    string _axeStdXCaption = "%, Specified amount",
                           _axeStdYCaption = "%, Area Ratio";


                    _serieStdName1 = "Media";

                    _xlAxisX_Std.TickLabels.NumberFormat = _axeStdXNumFormat; _xlAxisY_Std.TickLabels.NumberFormat = _axeStdYNumFormat;
                    _xlAxisX_Std.AxisTitle.Caption = _axeStdXCaption; _xlAxisY_Std.AxisTitle.Caption = _axeStdYCaption;

                    _xlSerie_Stds1.Name = _serieStdName1;


                    Excel.Trendlines _xlTrendlines_Stds1 = (Excel.Trendlines)_xlSerie_Stds1.Trendlines(Type.Missing);


                    Excel.XlTrendlineType _XlTrendlineType =
                        Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear;


                    _xlTrendlines_Stds1.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, true, true, Type.Missing);

                }



                //Unknowns
                _xlRange_Unknowns =
                    _wsComponent.get_Range
                    ("A" + (Title + _StdBracketCount + _BlankCount).ToString(), "V" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknowns.Name = _wsComponent.Name + "_Unknowns";

                _xlRange_Unknown_Area =
                    _wsComponent.get_Range
                    ("D" + (Title + _StdBracketCount + _BlankCount).ToString(), "D" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_Area.Name = _wsComponent.Name + "_Unknown_Area";

                _xlRange_Unknown_ISArea =
                    _wsComponent.get_Range
                    ("E" + (Title + _StdBracketCount + _BlankCount).ToString(), "E" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_ISArea.Name = _wsComponent.Name + "_Unknown_ISArea";

                _xlRange_Unknown_Ratio =
                   _wsComponent.get_Range
                   ("F" + (Title + _StdBracketCount + _BlankCount).ToString(), "F" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_Ratio.Name = _wsComponent.Name + "_Unknown_Ratio";

                string strCalcOrRAtio = "H";
                if (!_ifRel) { strCalcOrRAtio = "F"; }

                _xlRange_Unknown_CalcAmount =
                    _wsComponent.get_Range
                    (strCalcOrRAtio + (Title + _StdBracketCount + _BlankCount).ToString(), strCalcOrRAtio + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_CalcAmount.Name = _wsComponent.Name + "_Unknown_CalcAmount";

                _xlRange_Unknown_SampleID =
                    _wsComponent.get_Range
                    ("J" + (Title + _StdBracketCount + _BlankCount).ToString(), "J" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_SampleID.Name = _wsComponent.Name + "_Unknown_SampleID";

                _xlRange_Unknown_Subject =
                    _wsComponent.get_Range
                    ("K" + (Title + _StdBracketCount + _BlankCount).ToString(), "K" + (Title - 1 + _StdBracketCount + _UnknownCount + _BlankCount).ToString());
                _xlRange_Unknown_Subject.Name = _wsComponent.Name + "_Unknown_Subject";



                System.Array myvalues = (System.Array)_xlRange_Unknown_Subject.Cells.Value2;

                string[] subjectArray = ConvertToStringArray(myvalues);

                var subjectList = from _subject in subjectArray
                                  group _subject by _subject into _subjectGroupped
                                  orderby _subjectGroupped.Key
                                  select new { _subjectGroupped.Key, Count = _subjectGroupped.Count() };


                //*************************Format Ranges**************************
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;


                _xlRange_Unknowns.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                _xlRange_Unknowns.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                // _xlRange_Unknowns.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color. .Chocolate);

                //******Sort Unknowns
                _xlRange_Unknowns.Sort("Column K",
                    Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    "Column N", Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                    Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);

                //*************2011-07-19**************************************

                //_xlRange_Unknown_ISArea.FormatConditions.AddDatabar();

                Excel.ColorScale cfColorScale = (Excel.ColorScale)(_xlRange_Unknown_ISArea.FormatConditions.AddColorScale(3));

                // Set the minimum threshold to red (0x000000FF) and maximum threshold
                // to blue (0x00FF0000).
                cfColorScale.ColorScaleCriteria[1].FormatColor.Color = 7039480;  //0x000000FF;  
                cfColorScale.ColorScaleCriteria[2].FormatColor.Color = 8711167;  // 0x00FF0000;
                cfColorScale.ColorScaleCriteria[3].FormatColor.Color = 13011546; // 0x00FF0000;

                //*************************Charts*********************************
                //
                // Area Ratio
                Excel.ChartObject _xlChartTimeCorse = (Excel.ChartObject)_xlCharts.Add(
                                   0, _graphOffset, 750, 250);
                Excel.Chart _xlChartTimeCorsePage = _xlChartTimeCorse.Chart;

                _xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlXYScatter;

                _xlChartTimeCorsePage.HasTitle = true;
                _xlChartTimeCorsePage.ChartTitle.Font.Size = 11;

                _xlChartTimeCorsePage.ChartTitle.Caption = _componentName + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_TP = (Excel.Axes)_xlChartTimeCorsePage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                _xlAxisX_TP.HasMajorGridlines = true;

                _xlAxisX_TP.HasTitle = true;
                _xlAxisY_TP.HasTitle = true;

                _xlChartTimeCorsePage.Legend.Position = Excel.XlLegendPosition.xlLegendPositionRight;
                _xlChartTimeCorsePage.Legend.IncludeInLayout = true;

                _xlAxisX_TP.MinimumScale = 0; _xlAxisX_TP.MaximumScale = 49;
                _xlAxisY_TP.MinimumScale = 0;

                _xlAxisX_TP.MajorUnit = 6;

                Excel.SeriesCollection _xlSeriesColl_TimeCourse = (Excel.SeriesCollection)_xlChartTimeCorsePage.SeriesCollection(Type.Missing);


                int _rowIdx = 0, _subIdx = 0;

                foreach (var _subject in subjectList)
                {
                    _subIdx += 1;
                    Excel.Series _xlSeries_TTR = (Excel.Series)_xlSeriesColl_TimeCourse.NewSeries();
                    _xlSeries_TTR.Name = _subject.Key; /* + "_" +_componentName;*/

                    _xlSeries_TTR.XValues = _wsComponent.get_Range(_xlRange_Unknown_SampleID.Cells[_rowIdx + 1, 1],
                                                                   _xlRange_Unknown_SampleID.Cells[_rowIdx + _subject.Count, 1]);

                    _xlSeries_TTR.Values = _wsComponent.get_Range(_xlRange_Unknown_Ratio.Cells[_rowIdx + 1, 1],
                                                                  _xlRange_Unknown_Ratio.Cells[_rowIdx + _subject.Count, 1]);


                    Excel.Series _xlSeries_TTR_Norm = (Excel.Series)_xlSeriesColl_TimeCourse.NewSeries();
                    _xlSeries_TTR_Norm.Name = _subject.Key + /* "_" +_componentName +*/ "_Norm";

                    _xlSeries_TTR_Norm.XValues = _wsComponent.get_Range(_xlRange_Unknown_SampleID.Cells[_rowIdx + 1, 1],
                                                                        _xlRange_Unknown_SampleID.Cells[_rowIdx + _subject.Count, 1]);
                    _xlSeries_TTR_Norm.Values = _wsComponent.get_Range(_xlRange_Unknown_CalcAmount.Cells[_rowIdx + 1, 1],
                                                                       _xlRange_Unknown_CalcAmount.Cells[_rowIdx + _subject.Count, 1]);


                    _xlSeries_TTR.Format.Line.Visible = Office.MsoTriState.msoFalse;



                    _xlSeries_TTR_Norm.Format.Line.Visible = Office.MsoTriState.msoTrue;

                    switch (_subIdx)
                    {
                        case 1: _xlSeries_TTR_Norm.Format.Line.ForeColor.RGB = (int)Excel.XlRgbColor.rgbBlue;
                            _xlSeries_TTR_Norm.MarkerForegroundColor = (int)Excel.XlRgbColor.rgbBlue;
                            _xlSeries_TTR.MarkerForegroundColor = (int)Excel.XlRgbColor.rgbBlue;
                            _xlSeries_TTR_Norm.MarkerBackgroundColor = (int)Excel.XlRgbColor.rgbBlue;
                            _xlSeries_TTR.MarkerBackgroundColor = (int)Excel.XlRgbColor.rgbBlue;
                            _xlSeries_TTR.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleSquare;
                            _xlSeries_TTR_Norm.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleSquare;
                            break;
                        case 2: _xlSeries_TTR_Norm.Format.Line.ForeColor.RGB = (int)Excel.XlRgbColor.rgbRed;
                            _xlSeries_TTR_Norm.MarkerForegroundColor = (int)Excel.XlRgbColor.rgbRed;
                            _xlSeries_TTR.MarkerForegroundColor = (int)Excel.XlRgbColor.rgbRed;
                            _xlSeries_TTR_Norm.MarkerBackgroundColor = (int)Excel.XlRgbColor.rgbRed;
                            _xlSeries_TTR.MarkerBackgroundColor = (int)Excel.XlRgbColor.rgbRed;
                            _xlSeries_TTR.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleCircle;
                            _xlSeries_TTR_Norm.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleCircle;
                            break;
                        case 3: _xlSeries_TTR_Norm.Format.Line.ForeColor.RGB = (int)Excel.XlRgbColor.rgbGreen;
                            _xlSeries_TTR_Norm.MarkerForegroundColor = (int)Excel.XlRgbColor.rgbGreen;
                            _xlSeries_TTR.MarkerForegroundColor = (int)Excel.XlRgbColor.rgbGreen;
                            _xlSeries_TTR_Norm.MarkerBackgroundColor = (int)Excel.XlRgbColor.rgbGreen;
                            _xlSeries_TTR.MarkerBackgroundColor = (int)Excel.XlRgbColor.rgbGreen;
                            _xlSeries_TTR.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleTriangle;
                            _xlSeries_TTR_Norm.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleTriangle;
                            break;
                        default:
                            break;
                    }

                    _rowIdx += _subject.Count;
                }


                string _axeCurveXNumFormat = "0", _axeCurveYNumFormat = "0.000";

                string _axeCurveXCaption = "Time, h", _axeCurveYCaption = "TTR";


                if (!_ifRel)
                {
                    _axeCurveYCaption = "Area Ratio";
                    _axeCurveYNumFormat = "0.0";

                }

                _xlAxisX_TP.TickLabels.NumberFormat = _axeCurveXNumFormat; _xlAxisY_TP.TickLabels.NumberFormat = _axeCurveYNumFormat;

                _xlAxisX_TP.AxisTitle.Caption = _axeCurveXCaption; _xlAxisY_TP.AxisTitle.Caption = _axeCurveYCaption;

                /////////////////////////////////////////////////
                // Normalized

                Excel.ChartObject _xlChartTimeCorseNorm = (Excel.ChartObject)_xlCharts.Add(
                                   0, _graphOffset + _xlChartTimeCorse.Height, 750, 250);
                Excel.Chart _xlChartTimeCorseNormPage = _xlChartTimeCorseNorm.Chart;

                _xlChartTimeCorseNormPage.ChartType = Excel.XlChartType.xlXYScatter;

                _xlChartTimeCorseNormPage.HasTitle = true;
                _xlChartTimeCorseNormPage.ChartTitle.Font.Size = 11;

                _xlChartTimeCorseNormPage.ChartTitle.Caption = _componentName + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_TP_Norm = (Excel.Axes)_xlChartTimeCorseNormPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_TP_Norm = _xlAxes_TP_Norm.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_TP_Norm = _xlAxes_TP_Norm.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                _xlAxisX_TP_Norm.HasMajorGridlines = true;

                _xlAxisX_TP_Norm.HasTitle = true;
                _xlAxisY_TP_Norm.HasTitle = true;

                _xlAxisX_TP_Norm.TickLabels.NumberFormat = "0";
                _xlAxisY_TP_Norm.TickLabels.NumberFormat = "0.0E+0";

                _xlAxisX_TP_Norm.MinimumScale = 0; _xlAxisX_TP_Norm.MaximumScale = 49;
                _xlAxisY_TP_Norm.MinimumScale = 0;

                _xlAxisX_TP_Norm.MajorUnit = 6;

                _xlAxisX_TP_Norm.AxisTitle.Caption = _axeCurveXCaption;
                _xlAxisY_TP_Norm.AxisTitle.Caption = "IS Area";

                _xlChartTimeCorseNormPage.Legend.Position = Excel.XlLegendPosition.xlLegendPositionRight;
                _xlChartTimeCorseNormPage.Legend.IncludeInLayout = true;

                Excel.SeriesCollection _xlSeriesColl_TimeCourseNorm = (Excel.SeriesCollection)_xlChartTimeCorseNormPage.SeriesCollection(Type.Missing);

                _rowIdx = 0; _subIdx = 0;

                foreach (var _subject in subjectList)
                {
                    _subIdx += 1;

                    Excel.Series _xlSeries_ISArea = (Excel.Series)_xlSeriesColl_TimeCourseNorm.NewSeries();

                    _xlSeries_ISArea.Name = _subject.Key + "_IS Area";

                    _xlSeries_ISArea.XValues = _wsComponent.get_Range(_xlRange_Unknown_SampleID.Cells[_rowIdx + 1, 1],
                                                                   _xlRange_Unknown_SampleID.Cells[_rowIdx + _subject.Count, 1]);

                    _xlSeries_ISArea.Values = _wsComponent.get_Range(_xlRange_Unknown_ISArea.Cells[_rowIdx + 1, 1],
                                                                  _xlRange_Unknown_ISArea.Cells[_rowIdx + _subject.Count, 1]);


                    _xlSeries_ISArea.Format.Line.Visible = Office.MsoTriState.msoFalse;

                    switch (_subIdx)
                    {
                        case 1:
                            _xlSeries_ISArea.MarkerForegroundColor = (int)Excel.XlRgbColor.rgbBlue;
                            _xlSeries_ISArea.MarkerBackgroundColor = (int)Excel.XlRgbColor.rgbBlue;
                            _xlSeries_ISArea.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleSquare;

                            break;
                        case 2:
                            _xlSeries_ISArea.MarkerForegroundColor = (int)Excel.XlRgbColor.rgbRed;
                            _xlSeries_ISArea.MarkerBackgroundColor = (int)Excel.XlRgbColor.rgbRed;
                            _xlSeries_ISArea.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleCircle;
                            break;
                        case 3:
                            _xlSeries_ISArea.MarkerForegroundColor = (int)Excel.XlRgbColor.rgbGreen;
                            _xlSeries_ISArea.MarkerBackgroundColor = (int)Excel.XlRgbColor.rgbGreen;
                            _xlSeries_ISArea.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleTriangle;
                            break;
                        default:
                            break;
                    }

                    //_rowIdx += _subject.Count;


                    Excel.ChartObjects _xlChartsSummary = (Excel.ChartObjects)_wsSummary.ChartObjects(Type.Missing);
                    Excel.ChartObject _xlChartCurveSummary; Excel.ChartObject _xlChartLevelSummary;
                    Excel.Chart _xlChartCurveSummaryPage; Excel.Chart _xlChartLevelSummaryPage;

                    if (_xlChartsSummary.Count < 1)
                    {
                        _xlChartCurveSummary = (Excel.ChartObject)_xlChartsSummary.Add(
                                       10, 10, 850, 500);
                        _xlChartCurveSummaryPage = _xlChartCurveSummary.Chart;

                        _xlChartCurveSummaryPage.ChartType = Excel.XlChartType.xlXYScatter;
                        _xlChartCurveSummaryPage.HasTitle = true; _xlChartCurveSummaryPage.ChartTitle.Font.Size = 11;

                        Excel.Axes _xlAxes_CurveSummary = (Excel.Axes)_xlChartCurveSummaryPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                        Excel.Axis _xlAxisX_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                        Excel.Axis _xlAxisY_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                        _xlAxisX_CurveSummary.HasMajorGridlines = true; _xlAxisY_CurveSummary.HasMajorGridlines = true;
                        _xlAxisX_CurveSummary.Format.Line.Weight = 3.0F; _xlAxisY_CurveSummary.Format.Line.Weight = 3.0F;

                        _xlAxisX_CurveSummary.HasTitle = true; _xlAxisY_CurveSummary.HasTitle = true;

                        _xlAxisX_CurveSummary.AxisTitle.Caption = _axeCurveXCaption;
                        _xlAxisY_CurveSummary.AxisTitle.Caption = "TTR, %";

                        _xlAxisX_CurveSummary.MinimumScale = 0; _xlAxisX_CurveSummary.MaximumScale = 49;
                        _xlAxisX_CurveSummary.MajorUnit = 6;

                        _xlAxisX_CurveSummary.TickLabels.NumberFormat = "0"; _xlAxisY_CurveSummary.TickLabels.NumberFormat = "0.0%";

                        _xlChartCurveSummaryPage.ChartTitle.Caption = "TTR in ApoE peptides Normalized";

                        _xlChartCurveSummaryPage.ChartTitle.Caption += ExperimentInfo._CommonChartNameStr;

                        _xlChartCurveSummaryPage.Legend.Position = Excel.XlLegendPosition.xlLegendPositionRight;
                        _xlChartCurveSummaryPage.Legend.IncludeInLayout = true;

                        _xlChartLevelSummary = (Excel.ChartObject)_xlChartsSummary.Add(
                                       10, 510, 850, 500);
                        _xlChartLevelSummaryPage = _xlChartLevelSummary.Chart;

                        _xlChartLevelSummaryPage.ChartType = Excel.XlChartType.xlXYScatter;
                        _xlChartLevelSummaryPage.HasTitle = true; _xlChartLevelSummaryPage.ChartTitle.Font.Size = 11;

                        Excel.Axes _xlAxes_LevelSummary = (Excel.Axes)_xlChartLevelSummaryPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                        Excel.Axis _xlAxisX_LevelSummary = _xlAxes_LevelSummary.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                        Excel.Axis _xlAxisY_LevelSummary = _xlAxes_LevelSummary.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                        _xlAxisX_LevelSummary.HasMajorGridlines = true; _xlAxisY_LevelSummary.HasMajorGridlines = true;
                        _xlAxisX_LevelSummary.Format.Line.Weight = 3.0F; _xlAxisY_LevelSummary.Format.Line.Weight = 3.0F;

                        _xlAxisX_LevelSummary.HasTitle = true; _xlAxisY_LevelSummary.HasTitle = true;

                        _xlAxisX_LevelSummary.AxisTitle.Caption = _axeCurveXCaption;
                        _xlAxisY_LevelSummary.AxisTitle.Caption = "Ratio";

                        _xlAxisX_LevelSummary.MinimumScale = 0; _xlAxisX_LevelSummary.MaximumScale = 49;
                        _xlAxisX_LevelSummary.MajorUnit = 6;

                        _xlAxisX_LevelSummary.TickLabels.NumberFormat = "0"; _xlAxisY_LevelSummary.TickLabels.NumberFormat = "0.0";

                        _xlChartLevelSummaryPage.ChartTitle.Caption = "ApoE peptides levels";

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
                       _ifRel == true ?
                                               (Excel.SeriesCollection)_xlChartCurveSummaryPage.SeriesCollection(Type.Missing)
                                               :
                                               (Excel.SeriesCollection)_xlChartLevelSummaryPage.SeriesCollection(Type.Missing);

                    Excel.Series _xlSeries_timeCourse = (Excel.Series)_xlSeriesColl_Summary.NewSeries();
                    _xlSeries_timeCourse.XValues = _wsComponent.get_Range(_xlRange_Unknown_SampleID.Cells[_rowIdx + 1, 1],
                                                                   _xlRange_Unknown_SampleID.Cells[_rowIdx + _subject.Count, 1]);


                    //_xlRange_Unknown_SampleID;


                    _xlSeries_timeCourse.Values = _wsComponent.get_Range(_xlRange_Unknown_CalcAmount.Cells[_rowIdx + 1, 1],
                                                                  _xlRange_Unknown_CalcAmount.Cells[_rowIdx + _subject.Count, 1]);

                    // _xlRange_Unknown_CalcAmount;

                    _xlSeries_timeCourse.Name = _subject.Key + " " + (string)_xlRange_Compound.Cells.Value2;


                    _rowIdx += _subject.Count;

                }


                toolStripProgressBarMain.PerformStep();

            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = false;

            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();

            _XXQN.Close();

            object misValue = System.Reflection.Missing.Value;

            try
            {
                _XXQN.Close();

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
                    ((Excel._Worksheet)_wsSummary).Activate();

                }
                else
                {
                    _xlApp.Quit();
                }
            }

            if (checkBox_ShowDebug.Checked)
            {
                MessageBox.Show("ApoE REL()");
            }
        }

        //******************************************************* 2014-04-21 for ABS Tau CSF-curve + Syhtn
        private void ExportTAU_ABS()
        {

            List<string> _strExportSQL = new List<string>();
            /*05-07-2010*/
            string _strAnalyte = null;
            double _Area = 0, _ISTD_Area = 0, _Area_Ratio = 0, _CalcAmount = 0;
            decimal dcm_Area_Ratio = 0;
            string _strSpecAmount = null, _strRT = null, _strISTD_RT = null;
            string _response = null;

            /*bool IsPeakFound = false;*/

            /*05-07-2010*/
            if (textBox_XQNfile.Text == "")
            {
                MessageBox.Show("Select Input XQN file, please");
                return;
            }


            _XXQN = new XCALIBURFILESLib.XXQNClass();

            try
            {
                _XXQN.Open(textBox_XQNfile.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("File: " + textBox_XQNfile.Text + " is invalid", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //MessageBox.Show(_exception.Message, _exception.GetBaseException().Message);
                return;
            }
            finally
            {

            }

            XCALIBURFILESLib.IXHeader _IXHeader = (XCALIBURFILESLib.IXHeader)_XXQN.Header;

            ExperimentInfo._ProcessedBy = _IXHeader.ChangedId;

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = _IXHeader.ChangedLogon;
            }

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = "unknown";
            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = true;
            toolStripStatusLabelMain.Visible = true;



            _xlApp = new Excel.ApplicationClass();
            _xlApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;



            Excel.Workbook xlWB;

            try
            {

                xlWB =
                    _xlApp.Workbooks.Add(
                    //@"c:\Project\XQNReader\XQNReader\LOAD_TSQ_Template_Summary.xlsx"
                Application.StartupPath + Path.DirectorySeparatorChar +
                @"\Template\TAU_TSQ_ABS_Template_Summary.xlsx"
                );

            }

            catch (Exception)
            {
                MessageBox.Show("File: " + Application.StartupPath + Path.DirectorySeparatorChar +
                @"\Template\TAU_TSQ_ABS_Template_Summary.xlsx", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //MessageBox.Show(_exception.Message, _exception.GetBaseException().Message);
                return;
            }
            finally
            {

            }

            Excel.Worksheet _wsSummary = (Excel.Worksheet)xlWB.ActiveSheet;

            Excel.Worksheet _wsCurrent = null;


            byte Title = 5;
            byte _UnknownCount = 0, _StdBracketCount = 0, _TotalCount = 0;




            XCALIBURFILESLib.XBracketGroups _XBracketGroups
                = (XCALIBURFILESLib.XBracketGroups)_XXQN.BracketGroups;

            XCALIBURFILESLib.XBracketGroup _XBracketGroup
                = (XCALIBURFILESLib.XBracketGroup)_XBracketGroups.Item(1);

            XCALIBURFILESLib.XQuanResults _XQuanResults
                = (XCALIBURFILESLib.XQuanResults)_XBracketGroup.QuanResults;


            //********_XQuanResults*******************************
            foreach (XCALIBURFILESLib.XQuanResult _XQuanResult in _XQuanResults)
            {
                toolStripProgressBarMain.Maximum = _XQuanResults.Count;

                XCALIBURFILESLib.XSample _Sample
                    = (XCALIBURFILESLib.XSample)_XQuanResult.Sample;




                _TotalCount++;
                switch (_Sample.SampleType)
                {
                    case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                        _UnknownCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                        _StdBracketCount++;
                        break;
                    default:
                        break;
                }

                XCALIBURFILESLib.XQuanComponents _XQuanSelectedComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Selected; /*  Method; */
                //  = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.User;  
                ////////////////////////////////////////
                XCALIBURFILESLib.XQuanComponents _XQuanMethodComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Method; /*  Method; */


                ////////////////////////////////////////
                short _comNo = 1;
                foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponent in _XQuanSelectedComponents)
                {
                    XCALIBURFILESLib.XComponent _XComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanComponent.Component;

                    XCALIBURFILESLib.XQuanComponent _XQuanMethodComponent = (XCALIBURFILESLib.XQuanComponent)
                    _XQuanMethodComponents.Item(_comNo);
                    _comNo++;

                    XCALIBURFILESLib.XComponent _XMethodComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanMethodComponent.Component;

                    /*
                    XCALIBURFILESLib.XComponent _XComponentUser =
                        (XCALIBURFILESLib.XComponent)_XQuanComponent.Component;
                    */

                    //****************** 
                    XCALIBURFILESLib.IXDetection2 _IXDetection =
                        (XCALIBURFILESLib.IXDetection2)_XComponent.Detection;
                    //*****************

                    XCALIBURFILESLib.XQuanCalibration _XQuanCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XComponent.Calibration;

                    XCALIBURFILESLib.XQuanCalibration _XQuanMethodCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XMethodComponent.Calibration;

                    string _quanTypeKey = "";

                    // MessageBox.Show(ExperimentInfo._quanType.ToString());

                    if (ExperimentInfo._quanType == "RELATIVE")
                    {
                        _quanTypeKey = "N14:N15";
                    }
                    if (ExperimentInfo._quanType == "ABSOLUTE")
                    {
                        _quanTypeKey = "C13:C12";
                    }

                    //MessageBox.Show(_XQuanCalibration.ComponentType.ToString());
                    //MessageBox.Show(_XComponent.Name.Contains(_quanTypeKey).ToString());

                    if (_XQuanCalibration.ComponentType != XCALIBURFILESLib.XComponentType.XISTD
                        && !_XComponent.Name.Contains(_quanTypeKey))
                    {
                        _strAnalyte = _XComponent.Name.Replace("_", " "); //*******************

                        // ****** 2010-07-23
                        _strAnalyte = _strAnalyte.Replace("AB-Total", "AbetaTotal");
                        // ****** 2010-07-23


                        bool _SheetExist = false;
                        string _ComponentName = _XComponent.Name.Replace(":", "_");
                        _ComponentName = _ComponentName.Replace("-", "_");
                        //**********
                        _ComponentName = _ComponentName.Replace("(", "_");
                        _ComponentName = _ComponentName.Replace(")", "_");
                        //**********

                        foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
                        {

                            if (_wsComponent.Name == _ComponentName
                                )
                            {
                                _wsCurrent = _wsComponent;
                                _SheetExist = true;
                            }
                        }



                        if (!_SheetExist)
                        {
                            _wsCurrent = (Excel.Worksheet)xlWB.Sheets.Add
                          (Type.Missing, Type.Missing, Type.Missing,
                          Application.StartupPath + Path.DirectorySeparatorChar +
                            @"\Template\TAU_TSQ_ABS_Template.xlsx"
                          );

                            _wsCurrent.Name = _ComponentName;



                            _wsCurrent.Cells[2, 1] = _XComponent.Name.Replace("_", " ");

                            _wsCurrent.Cells[2, 3] =
                                CurveTypeToStr(_XQuanMethodCalibration.CurveType);
                            _wsCurrent.Cells[2, 4] =
                                weightingTypeToStr(_XQuanMethodCalibration.Weighting);
                            _wsCurrent.Cells[2, 6] =
                                OriginTypeToStr(_XQuanMethodCalibration.OriginType);

                            _wsCurrent.Cells[2, 8] = _XQuanMethodCalibration.FullEquation;

                        }



                        _wsCurrent.Cells[_Sample.RowNumber + Title, 1] = _Sample.FileName;
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 2] = SampleTypeToStr(_Sample.SampleType);
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 3] = IntegTypeToStr(_XQuanComponent.Selected_Type);


                        XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeak;


                        //************************************
                        /*
                        if (_Sample.RowNumber > 23 && _ComponentName.Contains("42") )
                        {
                            MessageBox.Show(_IXDetection.PeakWidthHeight.ToString()
                                + "  " + _IXDetection.SNThreshold.ToString());
                        }
                        */
                        //************************************

                        bool _IsPeakFoundANALYTE = false;
                        bool _IsPeakFoundISTD = false;


                        if (_XQuanComponent.DoesDetectedQuanPeakExist > 0)
                        {
                            _IsPeakFoundANALYTE = true;
                            _XDetectedQuanPeak =
                                (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponent.DetectedQuanPeak;

                            _wsCurrent.Cells[_Sample.RowNumber + Title, 4] = _XDetectedQuanPeak.Area;
                            _Area = _XDetectedQuanPeak.Area;

                            //_XDetectedQuanPeak.ISTDValid

                            _wsCurrent.Cells[_Sample.RowNumber + Title, 8] = _XDetectedQuanPeak.CalcAmount;
                            _CalcAmount = _XDetectedQuanPeak.CalcAmount;

                            if (_XDetectedQuanPeak.ISTDArea != 0)
                            {
                                _IsPeakFoundISTD = true;
                                _wsCurrent.Cells[_Sample.RowNumber + Title, 5] = _XDetectedQuanPeak.ISTDArea;
                                _wsCurrent.Cells[_Sample.RowNumber + Title, 6] =
                                _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;

                                _ISTD_Area = _XDetectedQuanPeak.ISTDArea;
                                _Area_Ratio = _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;
                                dcm_Area_Ratio = (decimal)_XDetectedQuanPeak.Area / (decimal)_XDetectedQuanPeak.ISTDArea;
                            }
                            else
                            {
                                _wsCurrent.Cells[_Sample.RowNumber + Title, 5] = "N/F";
                                //_wsCurrent.Cells[_Sample.RowNumber + Title, 6] = "-";
                            }



                            //string _response = "";
                            _response = "NULL"; _strRT = "NULL"; _strISTD_RT = "NULL";

                            if (_XDetectedQuanPeak.IsResponseOK > 0) { _response = "Ok"; }
                            if (_XDetectedQuanPeak.IsResponseHigh > 0) { _response = "High"; }
                            if (_XDetectedQuanPeak.IsResponseLow > 0) { _response = "Low"; }
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 9] = _response;
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 12] = _XDetectedQuanPeak.ApexRT;

                            _strRT = _XDetectedQuanPeak.ApexRT.ToString("F2");
                            //_RT = _XDetectedQuanPeak.ApexRT;
                        }
                        else
                        {
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 9] = "Not Found";

                        }



                        string _sample_ID = "N/A";
                        double _specAmount = 0;
                        switch (_Sample.SampleType)
                        {
                            case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                                _sample_ID = _Sample.Level;
                                XCALIBURFILESLib.XCalibrationCurveData _XCalibrationCurveData =
                                (XCALIBURFILESLib.XCalibrationCurveData)_XQuanComponent.CalibrationCurveData;
                                XCALIBURFILESLib.XReplicateRows _XReplicateRows =
                                (XCALIBURFILESLib.XReplicateRows)_XCalibrationCurveData.ReplicateRows;

                                foreach (XCALIBURFILESLib.XReplicateRow _XReplicateRow in _XReplicateRows)
                                {
                                    if (_XReplicateRow.LevelName == _Sample.Level)
                                    {
                                        _specAmount = _XReplicateRow.Amount;
                                        _strSpecAmount = "'" + _XReplicateRow.Amount.ToString("F5") + "'";
                                    }
                                }

                                break;
                            case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                                _sample_ID = _Sample.SampleId;
                                _specAmount = -1;
                                _strSpecAmount = "NULL";
                                break;
                            default:
                                break;
                        }


                        bool _IsSampleIDDefinded = false;
                        int _intTP = -1;

                        if (_sample_ID == "")
                        {
                            _sample_ID = _Sample.FileName.Substring(_Sample.FileName.LastIndexOf('_') + 1);
                        };

                        try
                        {
                            _intTP = Convert.ToInt16(_sample_ID);
                        }
                        catch (System.FormatException _FoEx)
                        {
                            MessageBox.Show(_FoEx.Message);
                        };

                        if (_intTP < 0)
                        {
                            _sample_ID = "";
                            _IsSampleIDDefinded = false;
                        }
                        else
                        {
                            _sample_ID = _intTP.ToString();
                            _IsSampleIDDefinded = true;
                        }


                        if (_specAmount != -1)
                        {
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 7] = _specAmount;
                        }
                        else
                        {
                            _wsCurrent.Cells[_Sample.RowNumber + Title, 7] = "";
                        }
                        _wsCurrent.Cells[_Sample.RowNumber + Title, 10] = _sample_ID;



                        foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponentISTD in _XQuanSelectedComponents)
                        {
                            XCALIBURFILESLib.XComponent _XISTD =
                            (XCALIBURFILESLib.XComponent)_XQuanComponentISTD.Component;

                            if (_XISTD.Name == _XQuanCalibration.ISTD)
                            {
                                if (_XQuanComponentISTD.DoesDetectedQuanPeakExist > 0)
                                {
                                    XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeakISTD =
                                        (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponentISTD.DetectedQuanPeak;
                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 13] = _XDetectedQuanPeakISTD.ApexRT;

                                    _strISTD_RT = _XDetectedQuanPeakISTD.ApexRT.ToString("F2");
                                }
                                else
                                {
                                    _wsCurrent.Cells[_Sample.RowNumber + Title, 13] = "N/F";
                                }

                            }

                        }

                        // if (GetSlopeOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType)) == -1)
                        // {
                        //     return;
                        // }

                        string _subjectNum = "";

                        if (checkBox_if_multi_subject.Checked && (int)_Sample.SampleType == 0)
                        {
                            _subjectNum = _Sample.FileName.Split('_')[1];
                        }
                        else
                        {
                            if (ExperimentInfo._Subject.Contains("-"))
                            {
                                _subjectNum = ExperimentInfo._Subject.Split('-')[1];
                            }
                            else _subjectNum = ExperimentInfo._Subject;

                        }


                        if (_IsPeakFoundANALYTE && _IsPeakFoundISTD && _IsSampleIDDefinded && checkBox_ExportIntoDB.Checked)
                        {
                            _strExportSQL.Add("INSERT INTO TMP_TSQ_IMPORT ( ANYLYTE_NAME, FILENAME_DESC, SAMPLE_TYPE_ID, INTEG_TYPE_ID, AREA, ISTD_AREA, AREA_RATIO, " +
                                                                          "SPEC_AMOUNT, CALC_AMOUNT, LEVEL_TP, PEAK_SATUS, RT, ISTD_RT, QUAN_TYPE_ID, FLUID_TYPE_ID, " +
                                                                          "SUBJECT_NUM, PROCESS_BY_ID, ASSAY_DATE, ASSAY_TYPE_ID, ANTIBODY_ID, ENZYME_ID, DONE_BY, " +
                                                                          "SAMPLE_PROCESS_BY, EXPER_FILE, CURVE, WEIGHTING, ORIGIN, EQUATION, R, R_SQR, NUM_SLOPE, " +
                                                                          "NUM_INTERSECT )");
                            _strExportSQL.Add(String.Format("VALUES ('{0}', '{1}', {2}, '{3}', {4:F0}, {5:F0}, {6:F5}, {7:F}, {8:F5}, {9}, '{10}', {11}, {12}, {13}, {14}, '{15}', '{16}', '{17}', {18}, {19}, {20}, {21}, {22}, '{23}'," +
                                " '{24}', '{25}', '{26}', '{27}', {28:F5}, {29:F5}, {30:F5}, {31:F5});",
                                   _strAnalyte,                                 // 0
                                   _Sample.FileName,                            // 1
                                   (int)_Sample.SampleType,                     // 2
                                   (int)_XQuanComponent.Selected_Type,          // 3
                                   _Area,                                       // 4
                                   _ISTD_Area,                                  // 5
                                   _Area_Ratio,                                 // 6
                                   _strSpecAmount,                              // 7 
                                   _CalcAmount,                                 // 8
                                   _sample_ID,                                  // 9
                                   _response,                                   // 10
                                   _strRT,                                      // 11 
                                   _strISTD_RT,                                 // 12
                                   QuanTypeToDBInt(ExperimentInfo._quanType),
                                   FluidTypeToDBInt(ExperimentInfo._matrix),    // 14
                                   _subjectNum,       // 15
                                   ExperimentInfo._ProcessedBy,                 // 16
                                   ExperimentInfo._Date,                        // 17
                                   2,                                           // 18      { ASSAY_TYPE_ID   2 - "IP LC/MS (C-term)"    
                                   IPToDBInt(ExperimentInfo._IP),               // 19      { 3, // A-BODY  3 - HJ5.x
                                   EnzymeToDBInt(ExperimentInfo._Enzyme),       // 20      { 2, //Enzyme 2 - LysN
                                   1,                                           // 21      { Done by 1 - ovodvo
                                   1,                                           // 22      { Sample process by 1 - ovodvo
                                   ExperimentInfo._expName,                     // 23

                                   CurveTypeToStr(_XQuanMethodCalibration.CurveType),
                                   weightingTypeToStr(_XQuanMethodCalibration.Weighting),
                                   OriginTypeToStr(_XQuanMethodCalibration.OriginType),
                                   GetEquaOfEquation(_XQuanMethodCalibration.FullEquation),
                                   GetROfEquation(_XQuanMethodCalibration.FullEquation),
                                   GetR_SQROfEquation(_XQuanMethodCalibration.FullEquation),
                                   GetSlopeOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType)),
                                   GetInterOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType))


                                   ));

                        }
                    }
                }

                toolStripProgressBarMain.PerformStep();
                toolStripStatusLabelMain.Text = "Processing ..." + _Sample.FileName;
                toolStripStatusLabelMain.PerformClick();

            }

            //_xlApp.Visible = true;
            toolStripProgressBarMain.Visible = false;
            toolStripProgressBarMain.Value = 0;
            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();
            toolStripProgressBarMain.Visible = true;

            toolStripProgressBarMain.Maximum = xlWB.Worksheets.Count - 1;


            string _AppPath = Application.StartupPath + Path.DirectorySeparatorChar;

            string[] _sql_P1 = File.ReadAllLines(_AppPath + "sql_TSQImport_P1.sql");
            string[] _sql_P2 = File.ReadAllLines(_AppPath + "sql_TSQImport_P2.sql");
            string[] _sql_P3 = File.ReadAllLines(_AppPath + "sql_TSQImport_P3.sql");

            _strExportSQL.Insert(0, " ");
            _strExportSQL.InsertRange(0, _sql_P1.ToList<string>());
            _strExportSQL.AddRange(_sql_P3.ToList<string>());
            //_strExportSQL.AddRange(_sql_P2.ToList<string>());


            File.WriteAllLines(Path.ChangeExtension(textBox_Excelfile.Text, "sql"), _strExportSQL.ToArray());

            foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
            {
                _wsComponent.Name = "_" + _wsComponent.Name;

                toolStripStatusLabelMain.Text = "Formatting ..." + _wsComponent.Name;
                toolStripStatusLabelMain.PerformClick();

                if (_wsComponent.Name.Contains("Summary")) { continue; }

                Excel.ChartObjects _xlCharts = (Excel.ChartObjects)_wsComponent.ChartObjects(Type.Missing);
                string _componentName = _wsComponent.get_Range("A2", "A2").Value2.ToString();

                Excel.Range _xlRange_Compound,
                            _xlRange_StdBracket_Area,
                            _xlRange_StdBracket_ISArea,
                            _xlRange_StdBracket_SpecifiedAmount,
                            _xlRange_StdBracket_Ratio,
                            _xlRange_Unknown_Area,
                            _xlRange_Unknown_ISArea,
                            _xlRange_Unknown_CalcAmount,
                            _xlRange_Unknown_SampleID,
                            _xlRange_STD,
                            _xlRange_Unknowns,
                            _xlRange_All;

                _xlRange_Compound = _wsComponent.get_Range("A2", "A2");

                //All
                _xlRange_All =
                    _wsComponent.get_Range
                    ("A" + Title.ToString(), "M" + (Title - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_All.Name = _wsComponent.Name + "_All";


                //Std
                if (_StdBracketCount > 0)
                {
                    _xlRange_STD =
                        _wsComponent.get_Range
                        ("A" + Title.ToString(), "M" + (Title - 1 + _StdBracketCount).ToString());
                    _xlRange_STD.Name = _wsComponent.Name + "_STD";


                    _xlRange_StdBracket_Ratio =
                        _wsComponent.get_Range
                        ("F" + Title.ToString(), "F" + (Title - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_Ratio.Name = _wsComponent.Name + "_StdBracket_Area";

                    _xlRange_StdBracket_Area =
                        _wsComponent.get_Range
                        ("D" + Title.ToString(), "D" + (Title - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_Area.Name = _wsComponent.Name + "_StdBracket_Area";

                    _xlRange_StdBracket_ISArea =
                        _wsComponent.get_Range
                        ("E" + Title.ToString(), "E" + (Title - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_ISArea.Name = _wsComponent.Name + "_StdBracket_ISArea";

                    _xlRange_StdBracket_SpecifiedAmount =
                       _wsComponent.get_Range
                       ("G" + Title.ToString(), "G" + (Title - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_SpecifiedAmount.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount";


                    _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                    _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                    _xlRange_STD.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);

                    //_xlRange_StdBracket_ISArea.FormatConditions.AddDatabar();


                    //*********Sort Range************
                    _xlRange_STD.Sort("Column J",
                        Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                        Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                        Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                        Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                        Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                        Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                        Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                        Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                        Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);




                    //****************
                    // StdCurve;
                    Excel.ChartObject _xlChartStdCurve = (Excel.ChartObject)_xlCharts.Add(
                                       490, 800, 470, 250);
                    Excel.Chart _xlChartStdCurvePage = _xlChartStdCurve.Chart;
                    _xlChartStdCurvePage.ChartType = Excel.XlChartType.xlXYScatter;

                    _xlChartStdCurvePage.HasTitle = true;
                    _xlChartStdCurvePage.ChartTitle.Font.Size = 11;

                    string _chartTitleStdCurve = String.Format
                        ("Calibration curve {0} {1} {2} {3} ({4})",
                        _componentName,
                        ExperimentInfo._IP,
                        ExperimentInfo._Enzyme,
                        ExperimentInfo._Instument,
                        ExperimentInfo._Date);
                    //_xlSeries_TTR.Name = _componentName;

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
                    _xlSerie_Stds.XValues = _xlRange_StdBracket_SpecifiedAmount.Cells;
                    _xlSerie_Stds.Values = _xlRange_StdBracket_Ratio.Cells;


                    string _serieStdName = "H:L Abeta";

                    string _axeStdXNumFormat = "0.00",
                           _axeStdYNumFormat = "0.00";

                    string _axeStdXCaption = "%, Predicted labeling Abeta",
                           _axeStdYCaption = "%, Measured labeling Abeta";

                    if (ExperimentInfo._quanType == "ABS")
                    {
                        _axeStdXNumFormat = "0.0";
                        _axeStdXNumFormat = "0.00";

                        _axeStdXCaption = "ng/mL, Specified amount";
                        _axeStdYCaption = "AreaRatio";

                        _serieStdName = "Analyte:ISTD Ratio";
                    }

                    _xlAxisX_Std.TickLabels.NumberFormat = _axeStdXNumFormat;
                    _xlAxisY_Std.TickLabels.NumberFormat = _axeStdYNumFormat;
                    _xlAxisX_Std.AxisTitle.Caption = _axeStdXCaption;
                    _xlAxisY_Std.AxisTitle.Caption = _axeStdYCaption;

                    _xlSerie_Stds.Name = _serieStdName;

                    Excel.Trendlines _xlTrendlines_Stds = (Excel.Trendlines)_xlSerie_Stds.Trendlines(Type.Missing);

                    Excel.XlTrendlineType _XlTrendlineType =
                        Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear;
                    string _CurveIndex = _wsComponent.get_Range("C2", "C2").Value2.ToString();

                    if (_CurveIndex == "Quadratic")
                    { _XlTrendlineType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlExponential; }


                    _xlTrendlines_Stds.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, true, true, Type.Missing);

                    //****************

                };


                //Unknowns
                _xlRange_Unknowns =
                    _wsComponent.get_Range
                    ("A" + (Title + _StdBracketCount).ToString(), "M" + (Title - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknowns.Name = _wsComponent.Name + "_Unknowns";

                _xlRange_Unknown_Area =
                    _wsComponent.get_Range
                    ("D" + (Title + _StdBracketCount).ToString(), "D" + (Title - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_Area.Name = _wsComponent.Name + "_Unknown_Area";

                _xlRange_Unknown_ISArea =
                    _wsComponent.get_Range
                    ("E" + (Title + _StdBracketCount).ToString(), "E" + (Title - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_ISArea.Name = _wsComponent.Name + "_Unknown_ISArea";

                string strCalcOrRAtio = "H";
                if (_StdBracketCount == 0) { strCalcOrRAtio = "F"; }
                _xlRange_Unknown_CalcAmount =
                    _wsComponent.get_Range
                    (strCalcOrRAtio + (Title + _StdBracketCount).ToString(), strCalcOrRAtio + (Title - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_CalcAmount.Name = _wsComponent.Name + "_Unknown_CalcAmount";

                _xlRange_Unknown_SampleID =
                    _wsComponent.get_Range
                    ("J" + (Title + _StdBracketCount).ToString(), "J" + (Title - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_SampleID.Name = _wsComponent.Name + "_Unknown_SampleID";







                //*************************Format Ranges**************************
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;

                //******Sort Unknowns
                _xlRange_Unknowns.Sort("Column J",
                    Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                    Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);

                //*************2011-07-19**************************************

                //_xlRange_Unknown_ISArea.FormatConditions.AddDatabar();

                Excel.ColorScale cfColorScale = (Excel.ColorScale)(_xlRange_Unknown_ISArea.FormatConditions.AddColorScale(3));

                // Set the minimum threshold to red (0x000000FF) and maximum threshold
                // to blue (0x00FF0000).
                cfColorScale.ColorScaleCriteria[1].FormatColor.Color = 7039480;  //0x000000FF;  
                cfColorScale.ColorScaleCriteria[2].FormatColor.Color = 8711167;  // 0x00FF0000;
                cfColorScale.ColorScaleCriteria[3].FormatColor.Color = 13011546; // 0x00FF0000;

                //*************************Charts*********************************
                //

                Excel.ChartObject _xlChartTimeCorse = (Excel.ChartObject)_xlCharts.Add(
                                   0, 800, 470, 250);
                Excel.Chart _xlChartTimeCorsePage = _xlChartTimeCorse.Chart;
                _xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlXYScatter;
                _xlChartTimeCorsePage.HasTitle = true;
                _xlChartTimeCorsePage.ChartTitle.Font.Size = 11;



                _xlChartTimeCorsePage.ChartTitle.Caption = _componentName + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_TP = (Excel.Axes)_xlChartTimeCorsePage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_TP.HasMajorGridlines = true;
                _xlAxisX_TP.HasTitle = true;
                _xlAxisY_TP.HasTitle = true;

                _xlChartTimeCorsePage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartTimeCorsePage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_TimeCourse = (Excel.SeriesCollection)_xlChartTimeCorsePage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_TTR = (Excel.Series)_xlSeriesColl_TimeCourse.NewSeries();
                _xlSeries_TTR.XValues = _xlRange_Unknown_SampleID.Cells;
                _xlSeries_TTR.Values = _xlRange_Unknown_CalcAmount;


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

                if (_StdBracketCount == 0)
                {
                    _axeCurveYCaption = "Area Ratio";
                }

                _xlAxisX_TP.TickLabels.NumberFormat = _axeCurveXNumFormat;
                _xlAxisY_TP.TickLabels.NumberFormat = _axeCurveYNumFormat;
                _xlAxisX_TP.AxisTitle.Caption = _axeCurveXCaption;
                _xlAxisY_TP.AxisTitle.Caption = _axeCurveYCaption;

                _xlSeries_TTR.Name = _serieCurveName;
                _xlSeries_TTR.Name = _componentName;








                Excel.ChartObjects _xlChartsSummary = (Excel.ChartObjects)_wsSummary.ChartObjects(Type.Missing);
                Excel.ChartObject _xlChartCurveSummary;
                Excel.Chart _xlChartCurveSummaryPage;

                if (_xlChartsSummary.Count < 1)
                {
                    _xlChartCurveSummary = (Excel.ChartObject)_xlChartsSummary.Add(
                                   10, 10, 850, 500);
                    _xlChartCurveSummaryPage = _xlChartCurveSummary.Chart;
                    _xlChartCurveSummaryPage.ChartType = Excel.XlChartType.xlXYScatter;
                    _xlChartCurveSummaryPage.HasTitle = true;
                    _xlChartCurveSummaryPage.ChartTitle.Font.Size = 11;



                    Excel.Axes _xlAxes_CurveSummary = (Excel.Axes)_xlChartCurveSummaryPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisX_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisY_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                    _xlAxisX_CurveSummary.HasMajorGridlines = true;
                    _xlAxisY_CurveSummary.HasMajorGridlines = true;

                    _xlAxisX_CurveSummary.Format.Line.Weight = 3.0F;
                    _xlAxisY_CurveSummary.Format.Line.Weight = 3.0F;

                    _xlAxisX_CurveSummary.HasTitle = true;
                    _xlAxisY_CurveSummary.HasTitle = true;

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
                _xlSeries_CurveTTR.XValues = _xlRange_Unknown_SampleID.Cells;
                _xlSeries_CurveTTR.Values = _xlRange_Unknown_CalcAmount;
                _xlSeries_CurveTTR.Name = (string)_xlRange_Compound.Cells.Value2;


                toolStripProgressBarMain.PerformStep();

            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = false;

            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();

            _XXQN.Close();

            object misValue = System.Reflection.Missing.Value;

            try
            {
                _XXQN.Close();

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
                    ((Excel._Worksheet)_wsSummary).Activate();

                }
                else
                {
                    _xlApp.Quit();
                }
            }
        }


        private void button_StartExport_Click(object sender, EventArgs e)
        {
            //ExportXcaliburUniTest(); //return;

            ExperimentInfo.reset();
            

            switch (comboBox_Project.Text)
            {
                case "LOAD100 CSF":
                    ExportLOAD();
                    break;

                case "ApoE ABS":
                    if (comboBox_QuantType.Text == "ABSOLUTE") { ExportAPOE_ABS(); }
                    break;

                case "ApoE REL":
                    ExportAPOE_REL();
                    break;
                case "ApoE REL KW":
                    
                    ExportAPOE_REL();
                    break;

                case "Tau":
                    if (comboBox_QuantType.Text == "ABSOLUTE") { ExportTAU_ABS(); }
                    break;

                case "Plasma Method Development":
                    ExportXcaliburUni();
                    break;

                case "LOAD100 Plasma: 9-hr Infusion":
                    ExportXcaliburLOAD_Plasma();
                    break;

                case "Zenith LOAD Plasma: IV-bolus (A)":
                    //MessageBox.Show("No method yet");
                    ExportXcaliburLOAD_Plasma();
                    break;

                case "LOAD60 Plasma: IV-bolus (D)":
                    //MessageBox.Show("No method yet");
                    ExportXcaliburLOAD_Plasma();
                    break;

                case "Plasma Abeta Validation":
                    //MessageBox.Show("No method yet");
                    ExportXcaliburLOAD_Plasma();
                    break;

                default:

                    ExportXcalibur_ForNico();
                    //MessageBox.Show("No procedure for export found", " Export", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    break;
            }

        }

        private void button_StartExport_Batch_Click(object sender, EventArgs e)
        {
            foreach (string _selectedFile in _FileList)
            {
                textBox_XQNfile.Text = _selectedFile;
                textBox_Excelfile.Text = Path.ChangeExtension(textBox_XQNfile.Text, "xlsx");
                getExperimentInfoType(Path.GetFileNameWithoutExtension(textBox_XQNfile.Text));

                this.button_StartExport_Click(this, null);

            }

        }

        private void checkBox_fileName_CheckedChanged(object sender, EventArgs e)
        {

            if (checkBox_fileName.Checked)
            {
                //groupBox_ExperInfo.Visible = false;
                button_StartExport_Batch.Enabled = true;
            }
            else
            {
                //groupBox_ExperInfo.Visible = true;
                button_StartExport_Batch.Enabled = false;
            }

        }

        private void fmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            RegistryKey _key;
            _key = Registry.CurrentUser.CreateSubKey("SOFTWARE\\XQNReader");
            try
            {
                //**********
                _key.SetValue(textBox_date_format.Name, textBox_date_format.Text);
                //********

                foreach (ComboBox _com in this.groupBox_ExperInfo.Controls.OfType<ComboBox>())
                {
                    _key.SetValue(_com.Name, _com.Text);
                }


                foreach (NumericUpDown _numer in this.groupBox_ExperInfo.Controls.OfType<NumericUpDown>())
                {
                    _key.SetValue(_numer.Name, _numer.Value);
                }
                foreach (CheckBox _chk in this.Controls.OfType<CheckBox>())
                {
                    _key.SetValue(_chk.Name, _chk.Checked);
                }

                foreach (CheckBox _chk in this.groupBox_ExperInfo.Controls.OfType<CheckBox>())
                {
                    _key.SetValue(_chk.Name, _chk.Checked);
                }
                //_key.SetValue("last_folder", _workFolder);


            }
            catch { }
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
            _key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\XQNReader");

            if (_key == null) { return; };


            //*************

            textBox_date_format.Text = _key.GetValue(textBox_date_format.Name).ToString();

            //*************



 
            try
            {

                foreach (ComboBox _combo in this.groupBox_ExperInfo.Controls.OfType<ComboBox>())
                {
                    _combo.Text = _key.GetValue(_combo.Name).ToString();
                }


                foreach (NumericUpDown _numer in this.groupBox_ExperInfo.Controls.OfType<NumericUpDown>())
                {
                    _numer.Value = Convert.ToInt32(_key.GetValue(_numer.Name));
                }
                //_workFolder = _key.GetValue("last_folder").ToString();

                foreach (CheckBox _chk in this.Controls.OfType<CheckBox>())
                {
                    _chk.Checked = Convert.ToBoolean(_key.GetValue(_chk.Name));
                }

                foreach (CheckBox _chk in this.groupBox_ExperInfo.Controls.OfType<CheckBox>())
                {
                    _chk.Checked = Convert.ToBoolean(_key.GetValue(_chk.Name));
                }


            }
            catch (Exception _exep)
            {
                MessageBox.Show("cannot read the registry - " + _exep.Message);
            }

            //****************
            _CustomDateFormat = textBox_date_format.Text;
        }

/*
        private void comboBox_Fluid_TextChanged(object sender, EventArgs e)
        {
            Control _com = (Control)sender;
            if (_com.Text != "")
            {
                this.getExperimentInfoType(null, false);
            }
            //MessageBox.Show( _com.Name );
        }
*/
        private void comboBox_Enzyme_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox_if_multi_subject_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_if_multi_subject.Checked)
            {
                this.textBox_subject.Enabled = false;
            }
            else
            {
                this.textBox_subject.Enabled = true;
            }
        }

        private void fmMain_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'dsBatemanLabDB.ASSAY_DESIGN' table. You can move, or remove it, as needed.
            this.aSSAY_DESIGNTableAdapter.Fill(this.dsBatemanLabDB.ASSAY_DESIGN);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.TIME_POINT' table. You can move, or remove it, as needed.
            this.tIME_POINTTableAdapter.Fill(this.dsBatemanLabDB.TIME_POINT);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.PROJECT' table. You can move, or remove it, as needed.
            this.pROJECTTableAdapter.Fill(this.dsBatemanLabDB.PROJECT);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.LAB_MEMBERS' table. You can move, or remove it, as needed.
            this.lAB_MEMBERSTableAdapter.Fill(this.dsBatemanLabDB.LAB_MEMBERS);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.FLUID_TYPE' table. You can move, or remove it, as needed.
            this.fLUID_TYPETableAdapter.Fill(this.dsBatemanLabDB.FLUID_TYPE);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.ENZYME' table. You can move, or remove it, as needed.
            this.eNZYMETableAdapter.Fill(this.dsBatemanLabDB.ENZYME);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.QUANT_TYPE' table. You can move, or remove it, as needed.
            this.qUANT_TYPETableAdapter.Fill(this.dsBatemanLabDB.QUANT_TYPE);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.EQUIPMENT' table. You can move, or remove it, as needed.
            this.eQUIPMENTTableAdapter.Fill(this.dsBatemanLabDB.EQUIPMENT);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.ANTIBODY' table. You can move, or remove it, as needed.
            this.aNTIBODYTableAdapter.Fill(this.dsBatemanLabDB.ANTIBODY);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.STUDY' table. You can move, or remove it, as needed.
            this.sTUDYTableAdapter.Fill(this.dsBatemanLabDB.STUDY);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(comboBox_Abody.SelectedValue.ToString());
            MessageBox.Show(ExperimentInfo._Date);

            MessageBox.Show(checkBox_if_multi_subject.FindForm().Name);
            MessageBox.Show(checkBox_ExportIntoDB.FindForm().Name);
        }

        private void dateTimePicker_AssayDate_ValueChanged(object sender, EventArgs e)
        {
            if (ExperimentInfo != null)
            {
                ExperimentInfo._Date = ((DateTimePicker)sender).Value.ToShortDateString();
            }

        }

        private void comboBox_QuantitatedBy_SelectionChangeCommitted(object sender, EventArgs e)
        {
            ExperimentInfo._SampleQuantitatedBy_ID = Convert.ToByte(comboBox_QuantitatedBy.SelectedValue);

        }

        private void comboBox_DoneBy_SelectionChangeCommitted(object sender, EventArgs e)
        {
            ExperimentInfo._SampleRanBy_ID = Convert.ToByte(comboBox_DoneBy.SelectedValue);
        }

        private void comboBox_SampleProcessBy_SelectionChangeCommitted(object sender, EventArgs e)
        {
            ExperimentInfo._SamplePreparedBy_ID = Convert.ToByte(comboBox_SampleProcessBy.SelectedValue);
        }

        private void comboBox_Instrument_SelectionChangeCommitted(object sender, EventArgs e)
        {
            ExperimentInfo._Instrument_ID = Convert.ToByte(comboBox_Instrument.SelectedValue);
        }


        private void comboBox_Fluid_SelectionChangeCommitted(object sender, EventArgs e)
        {
            ExperimentInfo._Matrix_ID = Convert.ToByte(comboBox_Fluid.SelectedValue);
        }


        //*******************************2016-05-22
        private void ExportXcaliburUni()
        {
            string expFunctionName = "ExportXcaliburUni()";
            string excelAnalyteTemplate = Application.StartupPath + Path.DirectorySeparatorChar + @"\Templates\QuanBrow_Universal_Analyte_Template.xlsx";
            string excelSummaryTemplate = Application.StartupPath + Path.DirectorySeparatorChar + @"\Templates\QuanBrow_Universal_Summary_Template.xlsx";

            bool ifDumpScanData = checkBox_ExportScanInfo.Checked;

            List<string> _strExportSQL = new List<string>();

            string _strAnalyte = null;
            double _Area = 0, _ISTD_Area = 0, _AreaRatio = 0, _CalcAmount = 0;
            decimal dcm_AreaRatio = 0;
            string _strSpecAmount = null, _strRT = null, _strISTD_RT = null;
            string _response = null;

            if (textBox_XQNfile.Text == "")
            {
                MessageBox.Show("Select Input XQN file, please");
                return;
            }

            _XXQN = new XCALIBURFILESLib.XXQNClass();

           // _XRaw = new XCALIBURFILESLib.IXParentScan();

            try
            {
                _XXQN.Open(textBox_XQNfile.Text);
            }
            catch (Exception _exception)
            {
                MessageBox.Show("File: " + textBox_XQNfile.Text + " is invalid XQN-file", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                if (checkBox_ShowDebug.Checked)
                {
                    MessageBox.Show(_exception.Message);
                }
                return;
            }
            finally
            {

            }


            XCALIBURFILESLib.IXHeader _IXHeader = (XCALIBURFILESLib.IXHeader)_XXQN.Header; //MessageBox.Show(_XXQN.FileName);

            ExperimentInfo._ProcessedBy = _IXHeader.ChangedId;

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = _IXHeader.ChangedLogon;
            }

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = "unknown";
            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = true;
            toolStripStatusLabelMain.Visible = true;

            _xlApp = new Excel.ApplicationClass();
            _xlApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;

            Excel.Workbook xlWB;


                      

            try
            {

                xlWB =
                    _xlApp.Workbooks.Add( excelSummaryTemplate  );

            }

            catch (Exception)
            {
                MessageBox.Show("File: " + excelSummaryTemplate + " causing Anchor exeption", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {

            }

            
            //xlWB.BuiltinDocumentProperties()
            

        //    xlWB.CustomDocumentProperties = new Office.DocumentProperties;

        //    Office.DocumentProperties _excelDocProperties = (Office.DocumentProperties)xlWB.CustomDocumentProperties;


       //     _excelDocProperties.Add("Test", false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, "Vitalik");
            
            //_excelDocProperties[1].Value = "lala";
            
            //(Office.DocumentProperties) xlWB.BuiltinDocumentProperties;


            Excel.Worksheet _wsSummary = (Excel.Worksheet)xlWB.ActiveSheet;

            Excel.Worksheet _wsCurrent = null;

            byte SkipFirstN = ExperimentInfo.TitleHeight; /*Title height of template in rows*/

            byte _UnknownCount = 0, _StdBracketCount = 0, _BlankCount = 0, _QCCount = 0, _TotalCount = 0;


            XCALIBURFILESLib.XBracketGroups _XBracketGroups
                = (XCALIBURFILESLib.XBracketGroups)_XXQN.BracketGroups;

            XCALIBURFILESLib.XBracketGroup _XBracketGroup
                = (XCALIBURFILESLib.XBracketGroup)_XBracketGroups.Item(1);

            XCALIBURFILESLib.XQuanResults _XQuanResults
                = (XCALIBURFILESLib.XQuanResults)_XBracketGroup.QuanResults;

            bool ifheaderForDump = true;

            int _exporSampleN = 0;
            int _curSampleN = 0;
            //********_XQuanResults*******************************
            foreach (XCALIBURFILESLib.XQuanResult _XQuanResult in _XQuanResults)
            {
                toolStripProgressBarMain.Maximum = _XQuanResults.Count;

                XCALIBURFILESLib.XSample _Sample
                    = (XCALIBURFILESLib.XSample)_XQuanResult.Sample; 

                
                /*Raw.ThermoRawFileReaderClass.GetScanNum(_Sample.Path + _Sample.FileName);*/    // MessageBox.Show( _Sample.Path + _Sample.FileName);

                //MessageBox.Show( Path.GetDirectoryName(textBox_XQNfile.Text) + /*_Sample.Path*/ _Sample.FileName + ".raw");

                string _extractFrom = Path.GetDirectoryName(textBox_XQNfile.Text) +"\\" + _Sample.FileName + ".raw";


                if (_extractFrom.Contains("Golden")) { _curSampleN++; /*MessageBox.Show(_curSampleN.ToString()); */}

                //MessageBox.Show(_extractFrom);

                int numSpectra = Raw.ThermoRawFileReaderClass.GetScanNum(_extractFrom);


                ExperimentInfo._totSampleCnt++;

                _TotalCount++;
                switch (_Sample.SampleType)
                {
                    case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                        _UnknownCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                        _StdBracketCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleBlank:
                        _BlankCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleQC:
                        _QCCount++;
                        break;

                    default:
                        MessageBox.Show("Unusual sample type " + _Sample.SampleType.ToString());
                        break;
                }


                XCALIBURFILESLib.XQuanComponents _XQuanSelectedComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Selected;

                XCALIBURFILESLib.XQuanComponents _XQuanMethodComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Method;


                short _comNo = 1;
                foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponent in _XQuanSelectedComponents)
                {
                    XCALIBURFILESLib.XComponent _XComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanComponent.Component;

                    XCALIBURFILESLib.XQuanComponent _XQuanMethodComponent = (XCALIBURFILESLib.XQuanComponent)
                    _XQuanMethodComponents.Item(_comNo);
                    _comNo++;

                    XCALIBURFILESLib.XComponent _XMethodComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanMethodComponent.Component;

                    XCALIBURFILESLib.IXDetection2 _IXDetection =
                        (XCALIBURFILESLib.IXDetection2)_XComponent.Detection;

                    XCALIBURFILESLib.XQuanCalibration _XQuanCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XComponent.Calibration;

                    XCALIBURFILESLib.XQuanCalibration _XQuanMethodCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XMethodComponent.Calibration;

                    string _quanTypeKey = "All";

                    if (ExperimentInfo._quanType == "RELATIVE")
                    {
                        _quanTypeKey = "C12N14";
                    }
                    if (ExperimentInfo._quanType == "ABSOLUTE")
                    {
                        _quanTypeKey = "C13N14";
                    }



                    string _ComponentName = _XComponent.Name.Replace(":", "_");
                    _ComponentName = _ComponentName.Replace("-", "_");

                    _ComponentName = _ComponentName.Replace("(", "_");
                    _ComponentName = _ComponentName.Replace(")", "_");

                    double _PeakLeftRT = 0.0, _PeakRightRT = 0.0;

                    XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeak;

                    
                    if ( _XQuanCalibration.ComponentType != XCALIBURFILESLib.XComponentType.XISTD &&
                        !_XComponent.Name.Contains(_quanTypeKey))
                    {
                        _strAnalyte = _XComponent.Name.Replace("_", " ");

                        _strAnalyte = _strAnalyte.Replace("AB-Total", "AbetaTotal");

                        bool _SheetExist = false;
                       


                        foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
                        {

                            if (_wsComponent.Name == _ComponentName
                                )
                            {
                                _wsCurrent = _wsComponent;
                                _SheetExist = true;
                            }
                        }



                        if (!_SheetExist)
                        {
                            _wsCurrent = (Excel.Worksheet)xlWB.Sheets.Add(
                                Type.Missing, Type.Missing, Type.Missing,
                                excelAnalyteTemplate
                          );

                            _wsCurrent.Name = _ComponentName;

                            string stampFormat = "Generated by XQN Reader {0} from {1}; " +
                             "-Function {2}; " +
                             "-Machine: {3} ran by {4}@{5}; " +
                             "-Timestamp: {6} @ {7}; " +
                             "-Source: {8}; " +
                             "-Analyte template {9}; " +
                             "-Summary template {10} ";

                            string stamp = String.Format(stampFormat, 
                                          Assembly.GetExecutingAssembly().GetName().Version.ToString(), Convert.ToDateTime(Properties.Resources.BuildDate).ToShortDateString(),
                                          expFunctionName,
                                          Environment.MachineName, Environment.UserName, Environment.UserDomainName,
                                          DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString(),
                                          textBox_XQNfile.Text,
                                          Path.GetFileName(excelAnalyteTemplate), 
                                          Path.GetFileName(excelSummaryTemplate)
                                          );

                            Excel.Comment GenerateStamp = (_wsCurrent.Cells[3, 1] as Excel.Range).AddComment(stamp);

                            _wsCurrent.Cells[2, 1] = _XComponent.Name.Replace("_", " ");

                            _wsCurrent.Cells[2, 3] =
                                CurveTypeToStr(_XQuanMethodCalibration.CurveType);
                            _wsCurrent.Cells[2, 4] =
                                weightingTypeToStr(_XQuanMethodCalibration.Weighting);
                            _wsCurrent.Cells[2, 6] =
                                OriginTypeToStr(_XQuanMethodCalibration.OriginType);

                            _wsCurrent.Cells[2, 8] = _XQuanMethodCalibration.FullEquation;

                            _wsCurrent.Cells[4, 11] = "Scans";
                            _wsCurrent.Cells[4, 14] = "LeftRT";
                            _wsCurrent.Cells[4, 15] = "RightRT";

                            //_wsCurrent.Cells[4, 17] = "SignalToNoise";
                            //_wsCurrent.Cells[4, 18] = "ConfidenceLevel";

                            _wsCurrent.Cells[4, 16] = "# Spec.";

                        }


                        int _currRowN = _Sample.RowNumber + SkipFirstN;

                        _wsCurrent.Cells[_currRowN, 1] = _Sample.FileName; 
                        _wsCurrent.Cells[_currRowN, 2] = SampleTypeToStr(_Sample.SampleType);
                        _wsCurrent.Cells[_currRowN, 3] = IntegTypeToStr(_XQuanComponent.Selected_Type);


                        


                        _XDetectedQuanPeak = null;

                        bool _IsPeakFoundANALYTE = false;
                        bool _IsPeakFoundISTD = false;

                        //double _PeakLeftRT = 0.0, _PeakRightRT = 0.0; 

                        if (_XQuanComponent.DoesDetectedQuanPeakExist > 0)
                        {
                            _IsPeakFoundANALYTE = true;
                            _XDetectedQuanPeak =
                                (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponent.DetectedQuanPeak;
//****************************
                            
                            _wsCurrent.Cells[_currRowN, 11] = _XDetectedQuanPeak.Scans;
                            _wsCurrent.Cells[_currRowN, 14] = _XDetectedQuanPeak.LeftRT;
                            //Excel.Range _rt = _wsCurrent.get_Range(_wsCurrent.Cells[_currRowN, 14], _wsCurrent.Cells[_currRowN, 15]).NumberFormat

                            
                            _wsCurrent.Cells[_currRowN, 15] = _XDetectedQuanPeak.RightRT;

                            (_wsCurrent.Cells[_currRowN, 14] as Excel.Range).NumberFormat = "0.00";
                            (_wsCurrent.Cells[_currRowN, 15] as Excel.Range).NumberFormat = "0.00";

                           // _wsCurrent.Cells[_currRowN, 17] = _XDetectedQuanPeak.SignalToNoise;
                           // _wsCurrent.Cells[_currRowN, 18] = _XDetectedQuanPeak.ConfidenceLevel;
                            _wsCurrent.Cells[_currRowN, 16] = numSpectra;
//****************************

                            _wsCurrent.Cells[_currRowN, 4] = _XDetectedQuanPeak.Area;
                            _Area = _XDetectedQuanPeak.Area; 


                            _wsCurrent.Cells[_currRowN, 8] = _XDetectedQuanPeak.CalcAmount;
                            _CalcAmount = _XDetectedQuanPeak.CalcAmount;

                            /*
                            if (_currRowN < 700 && !_ComponentName.Contains("Noise")) 
                            {
                                Raw.ThermoRawFileReaderClass.DumpScanHeaderDataForCompound
                                    (_extractFrom, Path.ChangeExtension(textBox_Excelfile.Text, "csv"),
                                    ifheaderForDump, _ComponentName, _XDetectedQuanPeak.LeftRT, _XDetectedQuanPeak.RightRT, 0.5);
                            }*/

                            _PeakLeftRT = _XDetectedQuanPeak.LeftRT;
                            _PeakRightRT = _XDetectedQuanPeak.RightRT;

                           

                            if (_XDetectedQuanPeak.ISTDArea != 0)
                            {
                                _IsPeakFoundISTD = true;
                                _wsCurrent.Cells[_currRowN, 5] = _XDetectedQuanPeak.ISTDArea;
                                _wsCurrent.Cells[_currRowN, 6] =
                                _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;

                                _ISTD_Area = _XDetectedQuanPeak.ISTDArea;
                                _AreaRatio = _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;
                                dcm_AreaRatio = (decimal)_XDetectedQuanPeak.Area / (decimal)_XDetectedQuanPeak.ISTDArea;
                            }
                            else
                            {
                                _wsCurrent.Cells[_currRowN, 5] = "N/F";

                            }


                            _response = "NULL"; _strRT = "NULL"; _strISTD_RT = "NULL";

                            if (_XDetectedQuanPeak.IsResponseOK > 0) { _response = "Ok"; }
                            if (_XDetectedQuanPeak.IsResponseHigh > 0) { _response = "High"; }
                            if (_XDetectedQuanPeak.IsResponseLow > 0) { _response = "Low"; }
                            _wsCurrent.Cells[_currRowN, 9] = _response;
                            _wsCurrent.Cells[_currRowN, 12] = _XDetectedQuanPeak.ApexRT;

                            _strRT = _XDetectedQuanPeak.ApexRT.ToString("F2");
                        }
                        else
                        {
                            _wsCurrent.Cells[_currRowN, 9] = "Not Found";

                        }


                        if (/*_curSampleN < 4 && _exporSampleN < 4 &&*/ !_ComponentName.Contains("Noise") && _Sample.FileName.Contains("Golden-Plasma") && ifDumpScanData) /*Remove after debug*/
                        {
                           // MessageBox.Show("ifheader (exc) - " + _ComponentName + " " + ifheaderForDump.ToString());

                            MessageBox.Show("allo");

                            int exported = Raw.ThermoRawFileReaderClass.DumpScanHeaderDataForCompound
                                (_extractFrom, Path.ChangeExtension(textBox_Excelfile.Text, "csv"),
                                ifheaderForDump, _ComponentName, _PeakLeftRT, _PeakRightRT, 0.5);

                            if (exported > 0) { ifheaderForDump = false; }
                            _exporSampleN++;
                        }

                       


                        string _sample_ID = "N/A";
                        double _specAmount = 0;
                        switch (_Sample.SampleType)
                        {
                            case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                                _sample_ID = _Sample.Level;
                                XCALIBURFILESLib.XCalibrationCurveData _XCalibrationCurveData =
                                (XCALIBURFILESLib.XCalibrationCurveData)_XQuanComponent.CalibrationCurveData;
                                XCALIBURFILESLib.XReplicateRows _XReplicateRows =
                                (XCALIBURFILESLib.XReplicateRows)_XCalibrationCurveData.ReplicateRows;

                                foreach (XCALIBURFILESLib.XReplicateRow _XReplicateRow in _XReplicateRows)
                                {
                                    if (_XReplicateRow.LevelName == _Sample.Level)
                                    {
                                        _specAmount = _XReplicateRow.Amount;
                                        _strSpecAmount = "'" + _XReplicateRow.Amount.ToString("F5") + "'";
                                    }
                                }

                                break;
                            case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                                _sample_ID = _Sample.SampleId;
                                _specAmount = -1;
                                _strSpecAmount = "NULL";
                                break;

                            case XCALIBURFILESLib.XSampleTypes.XSampleBlank:
                                _sample_ID = "BLANK";
                                break;

                            default:
                                break;
                        }


                        bool _IsSampleIDDefinded = false;
                        int _intTP = -1;

                        if (_sample_ID == "")
                        {
                            _sample_ID = _Sample.FileName.Substring(_Sample.FileName.LastIndexOf('_') + 1);
                        };

                        try
                        {
                            if (_sample_ID != "BLANK") 
                            {
                                _intTP = Convert.ToInt16(_sample_ID);
                            }
                        }
                        catch (System.FormatException)
                        {
                          //  MessageBox.Show(String.Format("Cannot convert '{0}' into Integer", _sample_ID));
                        };

                        if (_intTP < 0)
                        {
                            _sample_ID = "";
                            _IsSampleIDDefinded = false;
                        }
                        else
                        {
                            _sample_ID = _intTP.ToString();
                            _IsSampleIDDefinded = true;
                        }


                        if (_specAmount != -1)
                        {
                            _wsCurrent.Cells[_currRowN, 7] = _specAmount;
                        }
                        else
                        {
                            _wsCurrent.Cells[_currRowN, 7] = "";
                        }
                        _wsCurrent.Cells[_currRowN, 10] = _sample_ID;



                        foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponentISTD in _XQuanSelectedComponents)
                        {
                            XCALIBURFILESLib.XComponent _XISTD =
                            (XCALIBURFILESLib.XComponent)_XQuanComponentISTD.Component;

                            if (_XISTD.Name == _XQuanCalibration.ISTD)
                            {
                                if (_XQuanComponentISTD.DoesDetectedQuanPeakExist > 0)
                                {
                                    XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeakISTD =
                                        (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponentISTD.DetectedQuanPeak;
                                    _wsCurrent.Cells[_currRowN, 13] = _XDetectedQuanPeakISTD.ApexRT;

                                    _wsCurrent.Cells[_currRowN, 5] = _XDetectedQuanPeakISTD.Area;

                                    _strISTD_RT = _XDetectedQuanPeakISTD.ApexRT.ToString("F2");
                                }
                                else
                                {
                                    _wsCurrent.Cells[_currRowN, 13] = "N/F";

                                }

                            }

                        }

                        // if (GetSlopeOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType)) == -1)
                        // {
                        //     return;
                        // }

                        string _subjectNum = "";

                        if (checkBox_if_multi_subject.Checked && (int)_Sample.SampleType == 0)
                        {
                            _subjectNum = _Sample.FileName.Split('_')[1];
                        }
                        else
                        {
                            if (ExperimentInfo._Subject.Contains("-"))
                            {
                                _subjectNum = ExperimentInfo._Subject.Split('-')[1];
                            }
                            else _subjectNum = ExperimentInfo._Subject;

                        }


                        if (_IsPeakFoundANALYTE && _IsPeakFoundISTD && _IsSampleIDDefinded && checkBox_ExportIntoDB.Checked)
                        {

                            _strExportSQL.Add("INSERT INTO TMP_TSQ_IMPORT ( ANYLYTE_NAME, FILENAME_DESC, SAMPLE_TYPE_ID, INTEG_TYPE_ID, AREA, ISTD_AREA, AREA_RATIO, " +
                                                                          "SPEC_AMOUNT, CALC_AMOUNT, LEVEL_TP, PEAK_SATUS, RT, ISTD_RT, QUAN_TYPE_ID, FLUID_TYPE_ID, " +
                                                                          "SUBJECT_NUM, PROCESS_BY_ID, ASSAY_DATE, ASSAY_TYPE_ID, ANTIBODY_ID, ENZYME_ID, DONE_BY, " +
                                                                          "SAMPLE_PROCESS_BY, EXPER_FILE, CURVE, WEIGHTING, ORIGIN, EQUATION, R, R_SQR, NUM_SLOPE, " +
                                                                          "NUM_INTERSECT )");
                            _strExportSQL.Add(String.Format("VALUES ('{0}', '{1}', {2}, '{3}', {4:F0}, {5:F0}, {6:F5}, {7:F}, {8:F5}, {9}, '{10}', {11}, {12}, {13}, {14}, '{15}', {16}, '{17}', {18}, {19}, {20}, {21}, {22}, '{23}'," +
                                " '{24}', '{25}', '{26}', '{27}', {28:F5}, {29:F5}, {30:F5}, {31:F5});",
                                   _strAnalyte,                                 // 0
                                   _Sample.FileName,                            // 1
                                   (int)_Sample.SampleType,                     // 2
                                   (int)_XQuanComponent.Selected_Type,          // 3
                                   _Area,                                       // 4
                                   _ISTD_Area,                                  // 5
                                   _AreaRatio,                                  // 6
                                   _strSpecAmount,                              // 7 
                                   _CalcAmount,                                 // 8
                                   _sample_ID,                                  // 9
                                   _response,                                   // 10
                                   _strRT,                                      // 11 
                                   _strISTD_RT,                                 // 12
                                   QuanTypeToDBInt(ExperimentInfo._quanType),
                                   FluidTypeToDBInt(ExperimentInfo._matrix),    // 14
                                   _subjectNum,       // 15
                                   ExperimentInfo._SampleQuantitatedBy_ID,      // 16
                                   ExperimentInfo._Date,                        // 17
                                   2,                                           // 18      { ASSAY_TYPE_ID   2 - "IP LC/MS (C-term)"    
                                   IPToDBInt(ExperimentInfo._IP),               // 19      { 3, // A-BODY  3 - HJ5.x
                                   EnzymeToDBInt(ExperimentInfo._Enzyme),       // 20      { 2, //Enzyme 2 - LysN
                                   ExperimentInfo._SampleRanBy_ID,              // 21      { Ran by 1 - //ovodvo
                                   ExperimentInfo._SamplePreparedBy_ID,         // 22      { Sample prepared by //1 - ovodvo
                                   ExperimentInfo._expName,                     // 23

                                   CurveTypeToStr(_XQuanMethodCalibration.CurveType),
                                   weightingTypeToStr(_XQuanMethodCalibration.Weighting),
                                   OriginTypeToStr(_XQuanMethodCalibration.OriginType),
                                   GetEquaOfEquation(_XQuanMethodCalibration.FullEquation),
                                   GetROfEquation(_XQuanMethodCalibration.FullEquation),
                                   GetR_SQROfEquation(_XQuanMethodCalibration.FullEquation),
                                   GetSlopeOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType)),
                                   GetInterOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType))


                                   ));

                        }
                    }
                    else
                    {
                        if (_XQuanComponent.DoesDetectedQuanPeakExist > 0)
                        {
                            _XDetectedQuanPeak =
                               (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponent.DetectedQuanPeak;
                            _PeakLeftRT = _XDetectedQuanPeak.LeftRT;
                            _PeakRightRT = _XDetectedQuanPeak.RightRT;
                        }
                        
                        //MessageBox.Show("Ha");

                        if (/*_curSampleN < 4 && _exporSampleN < 4 &&*/ !_ComponentName.Contains("Noise") && _Sample.FileName.Contains("Golden-Plasma") && ifDumpScanData ) /*Remove after debug*/
                        {
                            //MessageBox.Show("ifheader (int) - " + _ComponentName + " " + ifheaderForDump.ToString());
                            int exported = Raw.ThermoRawFileReaderClass.DumpScanHeaderDataForCompound
                                (_extractFrom, Path.ChangeExtension(textBox_Excelfile.Text, "csv"),
                                ifheaderForDump, _ComponentName, _PeakLeftRT, _PeakRightRT, 0.5);

                            if (exported > 0) { ifheaderForDump = false; }
                            _exporSampleN++;
                           
                        }
                    }
                }

                

                toolStripProgressBarMain.PerformStep();
                toolStripStatusLabelMain.Text = "Processing ..." + _Sample.FileName;
                toolStripStatusLabelMain.PerformClick();

            }

            //_xlApp.Visible = true;
            toolStripProgressBarMain.Visible = false;
            toolStripProgressBarMain.Value = 0;
            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();
            toolStripProgressBarMain.Visible = true;

            toolStripProgressBarMain.Maximum = xlWB.Worksheets.Count - 1;


            string _AppPath = Application.StartupPath + Path.DirectorySeparatorChar;

            string[] _sql_P1 = File.ReadAllLines(_AppPath + "sql_TSQImport_P1.sql");
            string[] _sql_P2 = File.ReadAllLines(_AppPath + "sql_TSQImport_P2.sql");
            string[] _sql_P3 = File.ReadAllLines(_AppPath + "sql_TSQImport_P3.sql");

            _strExportSQL.Insert(0, " ");

            _strExportSQL.InsertRange(0, _sql_P1.ToList<string>());


            _strExportSQL.Insert(0, "/*"); //Assembly.GetExecutingAssembly().

            _strExportSQL.Insert(1, String.Format("Generated by XQN Reader {0} from {1}", Assembly.GetExecutingAssembly().GetName().Version.ToString(), Convert.ToDateTime(Properties.Resources.BuildDate).ToShortDateString()));

            _strExportSQL.Insert(2, String.Format("       Machine: {0} ran by {1}@{2}", Environment.MachineName, Environment.UserName, Environment.UserDomainName));
            _strExportSQL.Insert(3, String.Format("     Timestamp: {0} @ {1}", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString()));
            _strExportSQL.Insert(4, String.Format("        Source: {0} ", textBox_XQNfile.Text));

            _strExportSQL.Insert(5, "*/");


            _strExportSQL.Add(" ");

            _strExportSQL.AddRange(_sql_P3.ToList<string>());


            File.WriteAllLines(Path.ChangeExtension(textBox_Excelfile.Text, "sql"), _strExportSQL.ToArray());

            foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
            {
                string lastColLetter = "P";

                int len = _wsComponent.Name.Length;

                if (len > 30 )
                {
                    int _numToReplace = len - 30;
                    
                    _wsComponent.Name = _wsComponent.Name.Replace(_wsComponent.Name.Substring(len / 2, _numToReplace + 1), ".");
                }

                _wsComponent.Name = "_" + _wsComponent.Name;

                toolStripStatusLabelMain.Text = "Formatting ..." + _wsComponent.Name;
                toolStripStatusLabelMain.PerformClick();

                if (_wsComponent.Name.Contains("Summary")) { continue; }

                Excel.ChartObjects _xlCharts = (Excel.ChartObjects)_wsComponent.ChartObjects(Type.Missing);
                string _componentName = _wsComponent.get_Range("A2", "A2").Value2.ToString();

                Excel.Range _xlRange_Compound,
                            _xlRange_StdBracket_Area,
                            _xlRange_StdBracket_ISArea,
                            _xlRange_StdBracket_SpecifiedAmount,
                            _xlRange_StdBracket_Ratio,
                            _xlRange_Unknown_Area,
                            _xlRange_Unknown_ISArea,
                            _xlRange_Unknown_CalcAmount,
                            _xlRange_Unknown_SampleID,
                            _xlRange_STD,
                            _xlRange_Unknowns,
                            _xlRange_All,
                            _xlRange_AllbutStd;

                _xlRange_Compound = _wsComponent.get_Range("A2", "A2");

                //All
                _xlRange_All =
                    _wsComponent.get_Range
                    ("A" + SkipFirstN.ToString(), lastColLetter + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_All.Name = _wsComponent.Name + "_All";

                _xlRange_AllbutStd =
                     _wsComponent.get_Range
                    ("A" + (SkipFirstN + _StdBracketCount).ToString(), lastColLetter + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount + _QCCount + _BlankCount).ToString());
                _xlRange_AllbutStd.Name = _wsComponent.Name + "_AllbutStd";

//-----------------------

                string cellAddress = "A" + (ExperimentInfo.TitleHeight + ExperimentInfo._totSampleCnt + 2).ToString();

                double TimeCoursePlotOffsetY = (double)_wsCurrent.get_Range(cellAddress, cellAddress).Top;
                
                double TimeCoursePlotWidth = 470, TimeCoursePlotHight = 250;

                double StdsPlotOffsetY = TimeCoursePlotOffsetY + TimeCoursePlotHight;
                

                double StdsPlotWidth = 470, StdsPlotHight = 250;

                if (ExperimentInfo._StdBracketCount == 0) { TimeCoursePlotWidth = (double)_xlRange_All.Width; }

                double StdsPlotOffsetX = (double)_xlRange_All.Width - StdsPlotWidth;




                //Std
                if (_StdBracketCount > 0)
                {
                    _xlRange_STD =
                        _wsComponent.get_Range
                        ("A" + SkipFirstN.ToString(), lastColLetter + (SkipFirstN - 1 + _StdBracketCount).ToString());
                    _xlRange_STD.Name = _wsComponent.Name + "_STD";


                    _xlRange_StdBracket_Ratio =
                        _wsComponent.get_Range
                        ("F" + SkipFirstN.ToString(), "F" + (SkipFirstN - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_Ratio.Name = _wsComponent.Name + "_StdBracket_Area";

                    _xlRange_StdBracket_Area =
                        _wsComponent.get_Range
                        ("D" + SkipFirstN.ToString(), "D" + (SkipFirstN - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_Area.Name = _wsComponent.Name + "_StdBracket_Area";

                    _xlRange_StdBracket_ISArea =
                        _wsComponent.get_Range
                        ("E" + SkipFirstN.ToString(), "E" + (SkipFirstN - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_ISArea.Name = _wsComponent.Name + "_StdBracket_ISArea";

                    _xlRange_StdBracket_SpecifiedAmount =
                       _wsComponent.get_Range
                       ("G" + SkipFirstN.ToString(), "G" + (SkipFirstN - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_SpecifiedAmount.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount";


                    _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                    _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                    _xlRange_STD.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);

                    //_xlRange_StdBracket_ISArea.FormatConditions.AddDatabar();


                    //*********Sort Range************
                    
                      _xlRange_STD.Sort("Column J",
                          Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                          Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                          Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                          Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                          Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                          Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                          Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                          Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                          Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal); 


                    //****************
                    // StdCurve;
                    Excel.ChartObject _xlChartStdCurve = (Excel.ChartObject)_xlCharts.Add(
                                       StdsPlotOffsetX, StdsPlotOffsetY, StdsPlotWidth, StdsPlotHight);
                    Excel.Chart _xlChartStdCurvePage = _xlChartStdCurve.Chart;
                    _xlChartStdCurvePage.ChartType = Excel.XlChartType.xlXYScatter;

                    _xlChartStdCurvePage.HasTitle = true;
                    _xlChartStdCurvePage.ChartTitle.Font.Size = 11;

                    string _chartTitleStdCurve = String.Format
                        ("Calibration curve {0} {1} {2} {3} ({4})",
                        _componentName,
                        ExperimentInfo._IP,
                        ExperimentInfo._Enzyme,
                        ExperimentInfo._Instument,
                        ExperimentInfo._Date);
                    //_xlSeries_TTR.Name = _componentName;

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
                    _xlSerie_Stds.XValues = _xlRange_StdBracket_SpecifiedAmount.Cells;
                    _xlSerie_Stds.Values = _xlRange_StdBracket_Ratio.Cells;


                    string _serieStdName = "H:L Abeta";

                    string _axeStdXNumFormat = "0.00",
                           _axeStdYNumFormat = "0.00";

                    string _axeStdXCaption = "%, Predicted labeling Abeta",
                           _axeStdYCaption = "%, Measured labeling Abeta";

                    if (ExperimentInfo._quanType == "ABS")
                    {
                        _axeStdXNumFormat = "0.0";
                        _axeStdXNumFormat = "0.00";

                        _axeStdXCaption = "ng/mL, Specified amount";
                        _axeStdYCaption = "AreaRatio";

                        _serieStdName = "Analyte:ISTD Ratio";
                    }

                    _xlAxisX_Std.TickLabels.NumberFormat = _axeStdXNumFormat;
                    _xlAxisY_Std.TickLabels.NumberFormat = _axeStdYNumFormat;
                    _xlAxisX_Std.AxisTitle.Caption = _axeStdXCaption;
                    _xlAxisY_Std.AxisTitle.Caption = _axeStdYCaption;

                    _xlSerie_Stds.Name = _serieStdName;

                    Excel.Trendlines _xlTrendlines_Stds = (Excel.Trendlines)_xlSerie_Stds.Trendlines(Type.Missing);

                    Excel.XlTrendlineType _XlTrendlineType =
                        Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear;
                    string _CurveIndex = _wsComponent.get_Range("C2", "C2").Value2.ToString();

                    if (_CurveIndex == "Quadratic")
                    { _XlTrendlineType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlExponential; }


                    _xlTrendlines_Stds.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, true, true, Type.Missing);

                    //****************

                };


                //Unknowns
                _xlRange_Unknowns =
                    _wsComponent.get_Range
                    ("A" + (SkipFirstN + _StdBracketCount).ToString(), lastColLetter + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknowns.Name = _wsComponent.Name + "_Unknowns";

                _xlRange_Unknown_Area =
                    _wsComponent.get_Range
                    ("D" + (SkipFirstN + _StdBracketCount).ToString(), "D" + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_Area.Name = _wsComponent.Name + "_Unknown_Area";

                _xlRange_Unknown_ISArea =
                    _wsComponent.get_Range
                    ("E" + (SkipFirstN + _StdBracketCount).ToString(), "E" + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_ISArea.Name = _wsComponent.Name + "_Unknown_ISArea";

                string strCalcOrRAtio = "H";
                if (_StdBracketCount == 0) { strCalcOrRAtio = "F"; }
                _xlRange_Unknown_CalcAmount =
                    _wsComponent.get_Range
                    (strCalcOrRAtio + (SkipFirstN + _StdBracketCount).ToString(), strCalcOrRAtio + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_CalcAmount.Name = _wsComponent.Name + "_Unknown_CalcAmount";

                _xlRange_Unknown_SampleID =
                    _wsComponent.get_Range
                    ("J" + (SkipFirstN + _StdBracketCount).ToString(), "J" + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_Unknown_SampleID.Name = _wsComponent.Name + "_Unknown_SampleID";


                //*************************Format Ranges**************************
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;


              //  MessageBox.Show(string.Format("Std {0}, Unknown {1}, QC {2}, Blank {3}", _StdBracketCount, _UnknownCount, _QCCount, _BlankCount));

                //******Sort Unknowns

                _xlRange_AllbutStd.Sort("Column B",
                    Microsoft.Office.Interop.Excel.XlSortOrder.xlDescending,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                    Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal); 


                _xlRange_Unknowns.Sort("Column J",
                    Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                    Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal); 

                Excel.ColorScale cfColorScale_forISArea = (Excel.ColorScale)(_xlRange_Unknown_ISArea.FormatConditions.AddColorScale(3));

                cfColorScale_forISArea.ColorScaleCriteria[1].FormatColor.Color = 7039480;  //0x000000FF;  
                cfColorScale_forISArea.ColorScaleCriteria[2].FormatColor.Color = 8711167;  // 0x00FF0000;
                cfColorScale_forISArea.ColorScaleCriteria[3].FormatColor.Color = 13011546; // 0x00FF0000;

                Excel.ColorScale cfColorScale_forArea = (Excel.ColorScale)(_xlRange_Unknown_Area.FormatConditions.AddColorScale(3));

                cfColorScale_forArea = cfColorScale_forISArea;

     
                //*************************Charts*****************************************************************************************************

                //******* Unknowns************
 
                Excel.ChartObject _xlChartTimeCorse = (Excel.ChartObject)_xlCharts.Add(
                                   0, TimeCoursePlotOffsetY, TimeCoursePlotWidth, TimeCoursePlotHight);
                Excel.Chart _xlChartTimeCorsePage = _xlChartTimeCorse.Chart;
                _xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlXYScatter;
                _xlChartTimeCorsePage.HasTitle = true; _xlChartTimeCorsePage.ChartTitle.Font.Size = 11;

                _xlChartTimeCorsePage.ChartTitle.Caption = _componentName + ExperimentInfo._CommonChartNameStr;

                Excel.Axes _xlAxes_TP = (Excel.Axes)_xlChartTimeCorsePage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisX_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                Excel.Axis _xlAxisY_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                _xlAxisX_TP.HasMajorGridlines = true; _xlAxisX_TP.HasTitle = true; _xlAxisY_TP.HasTitle = true;

                _xlChartTimeCorsePage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                _xlChartTimeCorsePage.Legend.IncludeInLayout = false;

                Excel.SeriesCollection _xlSeriesColl_TimeCourse = (Excel.SeriesCollection)_xlChartTimeCorsePage.SeriesCollection(Type.Missing);

                Excel.Series _xlSeries_TTR = (Excel.Series)_xlSeriesColl_TimeCourse.NewSeries();
                _xlSeries_TTR.XValues = _xlRange_Unknown_SampleID.Cells; _xlSeries_TTR.Values = _xlRange_Unknown_CalcAmount;


                string _serieCurveName = "H:L Abeta";

                string _axeCurveXNumFormat = "0",
                       _axeCurveYNumFormat = "0.0%";

                string _axeCurveXCaption = "h, Time",
                       _axeCurveYCaption = "H:L ratio";

                if (ExperimentInfo._quanType == "ABSOLUTE")
                {
                    _axeCurveYNumFormat = "0.0";
                    _axeCurveYCaption = "ng/mL, Concentration";
                    _serieCurveName = "Abeta level";
                }

                if (_StdBracketCount == 0) { _axeCurveYCaption = "Area Ratio"; }

                _xlAxisX_TP.TickLabels.NumberFormat = _axeCurveXNumFormat; _xlAxisY_TP.TickLabels.NumberFormat = _axeCurveYNumFormat;
                _xlAxisX_TP.AxisTitle.Caption = _axeCurveXCaption; _xlAxisY_TP.AxisTitle.Caption = _axeCurveYCaption;

                _xlSeries_TTR.Name = _serieCurveName; _xlSeries_TTR.Name = _componentName;
                //**************************** 






                //**********************************************************************************************************************************************
                Excel.ChartObjects _xlChartsSummary = (Excel.ChartObjects)_wsSummary.ChartObjects(Type.Missing);
                Excel.ChartObject _xlChartCurveSummary;
                Excel.Chart _xlChartCurveSummaryPage;

                if (_xlChartsSummary.Count < 1)
                {
                    _xlChartCurveSummary = (Excel.ChartObject)_xlChartsSummary.Add(
                                   10, 10, 850, 500);
                    _xlChartCurveSummaryPage = _xlChartCurveSummary.Chart;
                    _xlChartCurveSummaryPage.ChartType = Excel.XlChartType.xlXYScatter;
                    _xlChartCurveSummaryPage.HasTitle = true;
                    _xlChartCurveSummaryPage.ChartTitle.Font.Size = 11;



                    Excel.Axes _xlAxes_CurveSummary = (Excel.Axes)_xlChartCurveSummaryPage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisX_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisY_CurveSummary = _xlAxes_CurveSummary.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                    _xlAxisX_CurveSummary.HasMajorGridlines = true; _xlAxisY_CurveSummary.HasMajorGridlines = true;

                    _xlAxisX_CurveSummary.Format.Line.Weight = 3.0F; _xlAxisY_CurveSummary.Format.Line.Weight = 3.0F;

                    _xlAxisX_CurveSummary.HasTitle = true; _xlAxisY_CurveSummary.HasTitle = true;

                    _xlAxisX_CurveSummary.AxisTitle.Caption = _axeCurveXCaption; _xlAxisY_CurveSummary.AxisTitle.Caption = _axeCurveYCaption;

                    _xlAxisX_CurveSummary.TickLabels.NumberFormat = _axeCurveXNumFormat; _xlAxisY_CurveSummary.TickLabels.NumberFormat = _axeCurveYNumFormat;

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
                _xlSeries_CurveTTR.XValues = _xlRange_Unknown_SampleID.Cells;
                _xlSeries_CurveTTR.Values = _xlRange_Unknown_CalcAmount;
                _xlSeries_CurveTTR.Name = (string)_xlRange_Compound.Cells.Value2;


                toolStripProgressBarMain.PerformStep();

            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = false;

            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();

            _XXQN.Close();

            object misValue = System.Reflection.Missing.Value;

            try
            {
                _XXQN.Close();

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
                    SortExclelWorkSheets(xlWB.Worksheets);

                    _xlApp.Visible = true;
                    ((Excel._Worksheet)_wsSummary).Activate();

                }
                else
                {
                    _xlApp.Quit();
                }
            }
            if (this.checkBox_ShowDebug.Checked)
            {
                MessageBox.Show("ExportXcaliburUniTEST()");
            }
        }

        [HandleProcessCorruptedStateExceptions]

        [SecurityCritical]

        private void ExportXcaliburLOAD_Plasma()
        {
            string expFunctionName = "ExportXcaliburLOAD_Plasma()";
            string excelAnalyteTemplate = Application.StartupPath + Path.DirectorySeparatorChar + @"\Templates\QuanBrow_Universal_Analyte_Template.xlsx";
            string excelSummaryTemplate = Application.StartupPath + Path.DirectorySeparatorChar + @"\Templates\QuanBrow_Universal_Summary_Template.xlsx";
            

            List<string> listAnalytesForDB = new List<string>()
                {
                     "Aß38_C13N14", "AßMD_C13N14", "Aß40_C13N14", "Aß42_C13N14", "Aß42_C13N14_M0",
                     "Aß38_C12N14", "AßMD_C12N14", "Aß40_C12N14", "Aß42_C12N14"
                };

            List<string> listAnalytesForDB_REL = new List<string>()
                {
                    "Aß38_C13N14", "AßMD_C13N14", "Aß40_C13N14", "Aß42_C13N14", "Aß42_C13N14_M0"                    
                };

            List<string> listAnalytesForDB_ABS = new List<string>()
                {
                    "Aß38_C12N14", "AßMD_C12N14", "Aß40_C12N14", "Aß42_C12N14"
                };

            bool ifDumpScanData = checkBox_ExportScanInfo.Checked;

            List<string> _strExportSQL = new List<string>();
            List<string> _strExcludedSTD = new List<string>();


            if (textBox_XQNfile.Text == "")
            {
                MessageBox.Show("Select Input XQN file, please");
                return;
            }

            _XXQN = new XCALIBURFILESLib.XXQNClass(); 
            
            try
            {
                _XXQN.Open(textBox_XQNfile.Text);
            }
            catch (Exception _exception)
            {
                MessageBox.Show("File: " + textBox_XQNfile.Text + " is invalid XQN-file", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                if (checkBox_ShowDebug.Checked)
                {
                    MessageBox.Show(_exception.Message);
                }
                return;
            }
            finally
            {

            }

         

            XCALIBURFILESLib.IXHeader _IXHeader = (XCALIBURFILESLib.IXHeader)_XXQN.Header; 

            ExperimentInfo._ProcessedBy = _IXHeader.ChangedId;

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = _IXHeader.ChangedLogon;
            }

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = "unknown";
            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = true;
            toolStripStatusLabelMain.Visible = true;

            _xlApp = new Excel.ApplicationClass();
            _xlApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;

            Excel.Workbook xlWB;

            try
            {
                xlWB =
                    _xlApp.Workbooks.Add(excelSummaryTemplate);
            }

            catch (Exception)
            {
                MessageBox.Show("File: " + excelSummaryTemplate + " causing Anchor exeption", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {

            }

            Excel.Worksheet _wsSummary = (Excel.Worksheet)xlWB.ActiveSheet;

            Excel.Worksheet _wsCurrent = null;

            byte SkipFirstN = ExperimentInfo.TitleHeight;

            byte _UnknownCount = 0, _StdBracketCount = 0, _BlankCount = 0, _QCCount = 0, _TotalCount = 0;

            XCALIBURFILESLib.XBracketGroups _XBracketGroups
                = (XCALIBURFILESLib.XBracketGroups)_XXQN.BracketGroups;

            XCALIBURFILESLib.XBracketGroup _XBracketGroup
                = (XCALIBURFILESLib.XBracketGroup)_XBracketGroups.Item(1);

            XCALIBURFILESLib.XQuanResults _XQuanResults
                = (XCALIBURFILESLib.XQuanResults)_XBracketGroup.QuanResults;

            bool ifheaderForDump = true;

            int _exporSampleN = 0;
            int _curSampleN = 0;

            foreach (XCALIBURFILESLib.IXQuanResult _XQuanResult in _XQuanResults)
            {
                toolStripProgressBarMain.Maximum = _XQuanResults.Count;

                XCALIBURFILESLib.XSample _Sample
                    = (XCALIBURFILESLib.XSample)_XQuanResult.Sample;

                string _extractFrom = Path.GetDirectoryName(textBox_XQNfile.Text) + "\\" + _Sample.FileName + ".raw";

                if (_extractFrom.Contains("Golden")) { _curSampleN++; }

                ExperimentInfo._totSampleCnt++;

                _TotalCount++;
                switch (_Sample.SampleType)
                {
                    case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                        _UnknownCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                        _StdBracketCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleBlank:
                        _BlankCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleQC:
                        _QCCount++;
                        break;

                    default:
                        MessageBox.Show("Unusual sample type " + _Sample.SampleType.ToString());
                        break;
                }

                XCALIBURFILESLib.XQuanComponents _XQuanSelectedComponents = null;
                //XCALIBURFILESLib.IXQuanComponents _XQuanSelectedComponents = null;

                try
                {
                    if (_XQuanResult == null) { MessageBox.Show("_XQuanResult is empty"); }

                    

                    _XQuanSelectedComponents
                       = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Selected;
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(_Sample.FileName);
                   // _XQuanSelectedComponents
                   //   = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Method;
                }


                XCALIBURFILESLib.XQuanComponents _XQuanMethodComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Method;
                   


                short _comNo = 1;
                foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponent in _XQuanSelectedComponents)
                {
                    XCALIBURFILESLib.XComponent _XComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanComponent.Component;

                    XCALIBURFILESLib.XQuanComponent _XQuanMethodComponent = (XCALIBURFILESLib.XQuanComponent)
                    _XQuanMethodComponents.Item(_comNo);
                    _comNo++;

                    XCALIBURFILESLib.XComponent _XMethodComponent =
                        (XCALIBURFILESLib.XComponent)_XQuanMethodComponent.Component;

                    XCALIBURFILESLib.IXDetection2 _IXDetection =
                        (XCALIBURFILESLib.IXDetection2)_XComponent.Detection;

                    XCALIBURFILESLib.XQuanCalibration _XQuanCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XComponent.Calibration;

                    XCALIBURFILESLib.XQuanCalibration _XQuanMethodCalibration =
                            (XCALIBURFILESLib.XQuanCalibration)_XMethodComponent.Calibration;

                    


                    string _quanTypeKey = "All";

                    if (ExperimentInfo._quanType == "RELATIVE")
                    {
                        _quanTypeKey = "N14:N15";
                    }
                    if (ExperimentInfo._quanType == "ABSOLUTE")
                    {
                        _quanTypeKey = "C13:C12";
                    }

                    bool _IsAnalyteForDB = listAnalytesForDB.Exists(x => x == _XComponent.Name); //if (!_IsAnalyteForDB) { continue; }

                    bool _IsAnalyteForDB_ABS = listAnalytesForDB_ABS.Exists(x => x == _XComponent.Name);

                    bool _IsAnalyteForDB_REL = listAnalytesForDB_REL.Exists(x => x == _XComponent.Name);

                    string _ComponentName = _XComponent.Name.Replace(":", "_");
                           _ComponentName = _ComponentName.Replace("-", "_");
                           _ComponentName = _ComponentName.Replace("(", "_");
                           _ComponentName = _ComponentName.Replace(")", "_");


                   string _strAnalyte = null;


                   string _strSpecAmount = null;

                   double _Area = 0, _ISTD_Area = 0, _AreaRatio = 0, _CalcAmount = 0;

                   int _Scans = 0;
                   
                   double _PeakRT = 0.0, _PeakLeftRT = 0.0, _PeakRightRT = 0.0;
                   double _ISTDPeakRT = 0.0, _ISTDPeakLeftRT = 0.0, _ISTDPeakRightRT = 0.0;
                   double _IT_C13_AVG = 0.0, _IT_C12_AVG = 0.0, _IT_N15_AVG = 0.0;
                   double _SN_C13_CNT = 0.0, _SN_C12_CNT = 0.0, _SN_N15_CNT = 0.0, _SN_TOT_CNT = 0.0; 

                   string _peakStatus = null;


                    XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeak;

                    if (_XQuanCalibration.ComponentType != XCALIBURFILESLib.XComponentType.XISTD &&
                        !_XComponent.Name.Contains(_quanTypeKey))
                    {
                        _strAnalyte = _XComponent.Name.Replace("_", " ");
                        _strAnalyte = _strAnalyte.Replace("AB-Total", "AbetaTotal");
                        bool _SheetExist = false;

                        foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
                        {
                            if (_wsComponent.Name == _ComponentName
                                )
                            {
                                _wsCurrent = _wsComponent;
                                _SheetExist = true;
                            }
                        }

                        if (!_SheetExist)
                        {
                            _wsCurrent = (Excel.Worksheet)xlWB.Sheets.Add( 
                                Type.Missing, Type.Missing, Type.Missing,
                                excelAnalyteTemplate
                          );

                            _wsCurrent.Name = _ComponentName;

                            string stampFormat = "Generated by XQN Reader {0} from {1}; " +
                             "-Function {2}; " +
                             "-Machine: {3} ran by {4}@{5}; " +
                             "-Timestamp: {6} @ {7}; " +
                             "-Source: {8}; " +
                             "-Analyte template {9}; " +
                             "-Summary template {10} ";

                            string stamp = String.Format(stampFormat,
                                          Assembly.GetExecutingAssembly().GetName().Version.ToString(), Convert.ToDateTime(Properties.Resources.BuildDate).ToShortDateString(),
                                          expFunctionName,
                                          Environment.MachineName, Environment.UserName, Environment.UserDomainName,
                                          DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString(),
                                          textBox_XQNfile.Text,
                                          Path.GetFileName(excelAnalyteTemplate),
                                          Path.GetFileName(excelSummaryTemplate)
                                          );

                            Excel.Comment GenerateStamp = (_wsCurrent.Cells[3, 1] as Excel.Range).AddComment(stamp);

                            _wsCurrent.Cells[2, 1] = _XComponent.Name.Replace("_", " ");
                            _wsCurrent.Cells[2, 3] =
                                CurveTypeToStr(_XQuanMethodCalibration.CurveType);
                            _wsCurrent.Cells[2, 4] =
                                weightingTypeToStr(_XQuanMethodCalibration.Weighting);
                            _wsCurrent.Cells[2, 6] =
                                OriginTypeToStr(_XQuanMethodCalibration.OriginType);
                            _wsCurrent.Cells[2, 8] = _XQuanMethodCalibration.FullEquation;
                        }

                        int _currRowN = _Sample.RowNumber + SkipFirstN;

                        _wsCurrent.Cells[_currRowN, 1] = _Sample.FileName;
                        _wsCurrent.Cells[_currRowN, 2] = SampleTypeToStr(_Sample.SampleType);
                        _wsCurrent.Cells[_currRowN, 3] = IntegTypeToStr(_XQuanComponent.Selected_Type);
                        (_wsCurrent.Cells[_currRowN, 27] as Excel.Range).FormulaR1C1 = "=RIGHT(RC[-26], 1)";

                        double _sampleVolume = 0.000;

                        _wsCurrent.Cells[_currRowN, 28] = _sampleVolume = _Sample.SampleVol;

                        _XDetectedQuanPeak = null;

                        bool _IsPeakFoundANALYTE = false;
                        bool _IsPeakFoundISTD = false;

                        //MessageBox.Show(_XQuanComponent.Selected_Type.ToString() + "  " + _Sample.FileName + _XComponent.Name);

                        if (  _XQuanComponent.DoesDetectedQuanPeakExist > 0)
                        {
                            _IsPeakFoundANALYTE = true;
                            _XDetectedQuanPeak =
                                (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponent.DetectedQuanPeak;

                            _wsCurrent.Cells[_currRowN, 4] =  _Area = _XDetectedQuanPeak.Area;
                            _wsCurrent.Cells[_currRowN, 8] =  _CalcAmount = _XDetectedQuanPeak.CalcAmount;

                            _wsCurrent.Cells[_currRowN, 11] = _Scans = _XDetectedQuanPeak.Scans;

                            _wsCurrent.Cells[_currRowN, 13] =  _PeakRT = _XDetectedQuanPeak.ApexRT;
                            _wsCurrent.Cells[_currRowN, 15] =  _PeakLeftRT = _XDetectedQuanPeak.LeftRT;
                            _wsCurrent.Cells[_currRowN, 16] =  _PeakRightRT = _XDetectedQuanPeak.RightRT;

                            _peakStatus = "NULL";
                            if (_XDetectedQuanPeak.IsResponseOK > 0) { _peakStatus = "Ok"; }
                            if (_XDetectedQuanPeak.IsResponseHigh > 0) { _peakStatus = "High"; }
                            if (_XDetectedQuanPeak.IsResponseLow > 0) { _peakStatus = "Low"; }       _wsCurrent.Cells[_currRowN, 9] = _peakStatus;

                        }
                        else
                        {
                            _wsCurrent.Cells[_currRowN, 9] = "N/F";
                        }

                        foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponentISTD in _XQuanSelectedComponents)
                        {
                            XCALIBURFILESLib.XComponent _XISTD =
                            (XCALIBURFILESLib.XComponent)_XQuanComponentISTD.Component;

                            if (_XISTD.Name == _XQuanCalibration.ISTD)
                            {
                                if (_XQuanComponentISTD.DoesDetectedQuanPeakExist > 0)
                                {
                                    _IsPeakFoundISTD = true;
                                    XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeakISTD =
                                        (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponentISTD.DetectedQuanPeak;

                                    _wsCurrent.Cells[_currRowN,  5] =  _ISTD_Area  = _XDetectedQuanPeakISTD.Area;
                                    _wsCurrent.Cells[_currRowN, 14] =  _ISTDPeakRT = _XDetectedQuanPeakISTD.ApexRT;
                                    _wsCurrent.Cells[_currRowN, 17] =  _ISTDPeakLeftRT   = _XDetectedQuanPeakISTD.LeftRT;
                                    _wsCurrent.Cells[_currRowN, 18] =  _ISTDPeakRightRT  = _XDetectedQuanPeakISTD.RightRT;
                                    
                                    Raw.ThermoRawFileReaderClass.PeptidePeaksExtra _peakExtra = new Raw.ThermoRawFileReaderClass.PeptidePeaksExtra();
                                    try
                                    {
                                        _peakExtra =
                                            Raw.ThermoRawFileReaderClass.GetPeptideExtrasForCompound(_extractFrom, _ComponentName, _XDetectedQuanPeakISTD.LeftRT, _XDetectedQuanPeakISTD.RightRT);
                                    }
                                    catch (Exception _ex)
                                    {
                                        if (checkBox_ShowError.Checked)
                                        {

                                            MessageBox.Show("Cannot get extras", "GetPeptideExtrasForCompound()", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            MessageBox.Show(_ex.Message);
                                            MessageBox.Show(_ex.Source);
                                            MessageBox.Show(_Sample.FileName + _ComponentName);
                                        }
                                    }
                                    _wsCurrent.Cells[_currRowN, 19] = _IT_C13_AVG = _peakExtra._AvgC13_InjT; _wsCurrent.Cells[_currRowN, 22] = _SN_C13_CNT = _peakExtra._NumSpecC13_InjT;
                                    _wsCurrent.Cells[_currRowN, 20] = _IT_C12_AVG = _peakExtra._AvgC12_InjT; _wsCurrent.Cells[_currRowN, 23] = _SN_C12_CNT = _peakExtra._NumSpecC13_InjT;
                                    _wsCurrent.Cells[_currRowN, 21] = _IT_N15_AVG = _peakExtra._AvgN15_InjT; _wsCurrent.Cells[_currRowN, 24] = _SN_N15_CNT = _peakExtra._NumSpecC13_InjT;
                                    _wsCurrent.Cells[_currRowN, 25] = _SN_TOT_CNT = _peakExtra._TotScans;
                                }
                                else
                                {
                                    //_wsCurrent.Cells[_currRowN, 14] = "N/F";
                                }
                            }
                        }
                             
                        if ( _IsPeakFoundANALYTE && _IsPeakFoundISTD)
                        {
                            _wsCurrent.Cells[_currRowN, 6] = _AreaRatio = _XDetectedQuanPeak.Area / _XDetectedQuanPeak.ISTDArea;
                        }
                        else
                        {
                            //_wsCurrent.Cells[_currRowN, 6] = "N/F";
                        }
 

                        if (!_ComponentName.Contains("Noise") && _Sample.FileName.Contains("Golden-Plasma") && ifDumpScanData) 
                       {
                            int exported = Raw.ThermoRawFileReaderClass.DumpScanHeaderDataForCompound
                                (_extractFrom, Path.ChangeExtension(textBox_Excelfile.Text, "csv"),
                                ifheaderForDump, _ComponentName, _PeakLeftRT, _PeakRightRT, 0.5);

                            if (exported > 0) { ifheaderForDump = false; }
                            _exporSampleN++;
                        }

                        string _sample_ID = "N/A";
                        double _specAmount = 0; float _intTPorLevel = -1;
                        bool _IsSampleIDDefinded = false;
                        string _subjectNum = ""; int _exclude = 0;
                        int _fluid_type_id = -1;


                   /*     MessageBox.Show(
                        String.Format("FileName {0}; _Sample.SampleId {1}, _Sample.SampleName {2}, _Sample.SampleType {3}; _Sample.UserText[0] {4}; _Sample.Comment {5} _Sample.CalFile {6} _Sample_InjVol {7}",
                            _Sample.FileName,
                            _Sample.SampleId,
                            _Sample.SampleName,
                            _Sample.SampleType.ToString(),
                            _Sample.get_UserText(1),
                            _Sample.Comment,//, _Sample.
                            _Sample.CalFile, _Sample.InjVol
                            )); */
                        
                        switch (_Sample.SampleType)
                        {
                            case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                                _sample_ID = _Sample.Level;
                                XCALIBURFILESLib.XCalibrationCurveData _XCalibrationCurveData =
                                (XCALIBURFILESLib.XCalibrationCurveData)_XQuanComponent.CalibrationCurveData;
                                
                                XCALIBURFILESLib.XReplicateRows _XReplicateRows =
                                (XCALIBURFILESLib.XReplicateRows)_XCalibrationCurveData.ReplicateRows;

                                foreach (XCALIBURFILESLib.XReplicateRow _XReplicateRow in _XReplicateRows)
                                {
                                    if ((_XReplicateRow.LevelName == _Sample.Level) && (Path.GetFileNameWithoutExtension(_XReplicateRow.ResultFileName) == _Sample.FileName))  
                                    {
                                        _specAmount = _XReplicateRow.Amount;
                                        _wsCurrent.Cells[_currRowN, 26] = _exclude = _XReplicateRow.ExcludeFromCalibration;
                                        _strSpecAmount = "'" + _XReplicateRow.Amount.ToString("F5") + "'";

                                        if (_exclude == 1)
                                        {
                                            _strExcludedSTD.Add(String.Format("{0}; {1}; {2}; {3};",
                                                                              ExperimentInfo._expName,
                                                                              _Sample.FileName,
                                                                              _strAnalyte,
                                                                              _Sample.Level )
                                                               );
                                        }
                                    }
                                    //MessageBox.Show("Rep " + Path.GetFileName(_XReplicateRow.ResultFileName));
                                    //MessageBox.Show("Sam " + _Sample.FileName);
                                }
                                _wsCurrent.Cells[_currRowN, 10] = _Sample.Level;
                                try
                                {
                                    _intTPorLevel = float.Parse(_Sample.Level);
                                }
                                catch (Exception _ex)
                                {
                                    MessageBox.Show(_ex.Message, "float.Parse(_Sample.Level):");
                                }
                               
                                _IsSampleIDDefinded = true;
                                break;

                            case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                                _wsCurrent.Cells[_currRowN, 10] = _sample_ID = _Sample.SampleId;
                                _specAmount = -1; _strSpecAmount = "NULL";


                               // MessageBox.Show(_sample_ID);

                                try
                                                                     
                                    {
                                        if (comboBox_ClinicalStudy.SelectedValue.ToString() == "8")
                                        {
                                            _intTPorLevel = GetTPfromDB("0", "Unknown");
                                        }
                                        else
                                        {
                                            _intTPorLevel = GetTPfromDB(_sample_ID, "Unknown");
                                        }
                                    }
                                catch (Exception _ex)
                                    {
                                        MessageBox.Show(_ex.Message, "GetTPfromDB():");
                                    }

                                if (_intTPorLevel < 0)
                                {
                                    _IsSampleIDDefinded = false;
                                }
                                else
                                {
                                    _IsSampleIDDefinded = true;
                                }

                                if (checkBox_if_multi_subject.Checked && (int)_Sample.SampleType == 0)
                                {
                                    _subjectNum = _Sample.FileName.Split('_')[1];

                                    if (_subjectNum.Contains("-"))
                                    {
                                        _subjectNum = _subjectNum.Split('-')[1];
                                    }

                                }
                                else
                                {
                                    if (ExperimentInfo._Subject.Contains("-"))
                                    {
                                        _subjectNum = ExperimentInfo._Subject.Split('-')[1];
                                    }
                                    else _subjectNum = ExperimentInfo._Subject;
                                }

                                _fluid_type_id = ExperimentInfo._Matrix_ID;

                                if (checkBox_if_parse_fluid.Checked && (int)_Sample.SampleType == 0)
                                {
                                    
                                    switch (_Sample.FileName.Split('_')[2])
                                    {
                                        case "Plasma": _fluid_type_id = 2;
                                            break;
                                        case "Plasma-H": _fluid_type_id = 5;
                                            break;
                                        default:
                                            _fluid_type_id = ExperimentInfo._Matrix_ID;
                                            break;
                                    }
                                }
                              

                                break;

                            case XCALIBURFILESLib.XSampleTypes.XSampleBlank:
                                _sample_ID = "BLANK";
                                break;

                            case XCALIBURFILESLib.XSampleTypes.XSampleQC:
                                _wsCurrent.Cells[_currRowN, 10] = _Sample.Level;
                                _specAmount = -1; _strSpecAmount = "NULL";
                                //_intTPorLevel = float.Parse(_Sample.Level);  _Sample.
                                _IsSampleIDDefinded = true;
                                _sample_ID = "QC";
                                break;

                            default:
                                break;
                        }
                     
                        if (_specAmount != -1)  { _wsCurrent.Cells[_currRowN, 7] = _specAmount; }  else { _wsCurrent.Cells[_currRowN, 7] = ""; }

                        bool _IsForDB =   _Sample.FileName.Contains("LOAD") 
                                       || _Sample.FileName.Contains("media")
                                       || _Sample.FileName.Contains("GPlasma")
                                       || _Sample.FileName.Contains("ABS-STD")
                                       || _Sample.FileName.Contains("ADRC");
 
                        // for inclution of LOAD plasma & media in ther DB only

                        if (_IsPeakFoundANALYTE && _IsPeakFoundISTD && _IsSampleIDDefinded && checkBox_ExportIntoDB.Checked && _IsForDB && _IsAnalyteForDB)
                        {
                            string CURVE = null, WEIGHTING = null, ORIGIN = null, EQUATION = null;
                            double R = 0, R_SQR = 0, NUM_SLOPE = 0, NUM_INTERSECT = 0;
 
 
                            if (_IsAnalyteForDB_REL)
                            {
                                CURVE = CurveTypeToStr(_XQuanMethodCalibration.CurveType);
                                WEIGHTING = weightingTypeToStr(_XQuanMethodCalibration.Weighting);
                                ORIGIN = OriginTypeToStr(_XQuanMethodCalibration.OriginType);
                                EQUATION = GetEquaOfEquation(_XQuanMethodCalibration.FullEquation);
                                
                                R = GetROfEquation(_XQuanMethodCalibration.FullEquation); R_SQR = GetR_SQROfEquation(_XQuanMethodCalibration.FullEquation);

                                if (_XQuanMethodCalibration.FullEquation != null && _XQuanMethodCalibration.FullEquation != "")
                                {
                                    NUM_SLOPE = GetSlopeOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType));
                                    NUM_INTERSECT = GetInterOfEquation(_XQuanMethodCalibration.FullEquation, OriginTypeToStr(_XQuanMethodCalibration.OriginType));
                                }
 
                            } 
                            else { }

                            _strExportSQL.Add("INSERT INTO TMP_MSQBROW_IMPORT ( ANALYTE_NAME, FILENAME_DESC, SAMPLE_TYPE_ID, INTEG_TYPE_ID, AREA, ISTD_AREA, AREA_RATIO, " +
                                                                          "SPEC_AMOUNT, CALC_AMOUNT, LEVEL_TP, PEAK_SATUS, RT, ISTD_RT, QUAN_TYPE_ID, FLUID_TYPE_ID, " +
                                                                          "SUBJECT_NUM, PROCESS_BY_ID, ASSAY_DATE, ASSAY_TYPE_ID, ANTIBODY_ID, ENZYME_ID, DONE_BY, " +
                                                                          "SAMPLE_PROCESS_BY, EXPER_FILE, CURVE, WEIGHTING, ORIGIN, EQUATION, R, R_SQR, NUM_SLOPE, " +
                                                                          "NUM_INTERSECT, PROJECT_ID, ASSAY_DESIGN_ID, " +
                                                                          "RT_LEFT, RT_RIGHT, ISTD_RT_LEFT, ISTD_RT_RIGHT, " +
                                                                          "PEAK_SCANS, " +
                                                                          "IT_C13_AVG, IT_C12_AVG, IT_N15_AVG, SN_C13_CNT, SN_C12_CNT, SN_N15_CNT, SN_TOT_CNT, " +
                                                                          "EXCLUDE, SAMPLE_VOLUME, EQUIPMENT_ID, " +
                                                                          "ASSAY_SAMPLE_COMMENT, ASSAY_CALFILE, ASSAY_CAL_EXPER_FILE, ASSAY_SAMPLE_INJVOL )");
                            _strExportSQL.Add(String.Format("VALUES ('{0}', '{1}', {2}, '{3}', {4:F0}, {5:F0}, {6:F5}, {7:F}, {8:F5}, {9}, '{10}', {11}, {12}, {13}, {14}, '{15}', {16}, '{17}', {18}, {19}, {20}, {21}, {22}, '{23}'," +
                                " '{24}', '{25}', '{26}', '{27}', {28:F5}, {29:F5}, {30:F5}, {31:F5}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}," +
                                " '{49}', '{50}', '{51}', {52} );",
                                   _strAnalyte,                                 // 0
                                   _Sample.FileName,                            // 1
                                   (int)_Sample.SampleType,                     // 2
                                   (int)_XQuanComponent.Selected_Type,          // 3
                                   _Area,                                       // 4
                                   _ISTD_Area,                                  // 5
                                   _AreaRatio,                                  // 6
                                   _strSpecAmount,                              // 7 
                                   _CalcAmount,                                 // 8
                                   _intTPorLevel,                               // 9
                                   _peakStatus,                                 // 10
                                   Math.Round(_PeakRT, 2),                      // 11 
                                   Math.Round(_ISTDPeakRT, 2),                  // 12
                                   QuanTypeToDBInt(ExperimentInfo._quanType),
                                   _fluid_type_id,                              // 14
                                   _subjectNum,                                 // 15
                                   ExperimentInfo._SampleQuantitatedBy_ID,      // 16
                                   ExperimentInfo._Date,                        // 17
                                   12,                                          // 18     { ASSAY_TYPE_ID   12 - "IP LC/MS (C-term)"    
                                   IPToDBInt(ExperimentInfo._IP),               // 19      { 3, // A-BODY  3 - HJ5.x
                                   EnzymeToDBInt(ExperimentInfo._Enzyme),       // 20      { 2, //Enzyme 2 - LysN
                                   ExperimentInfo._SampleRanBy_ID,              // 21      { Ran by 1 - //ovodvo
                                   ExperimentInfo._SamplePreparedBy_ID,         // 22      { Sample prepared by //1 - ovodvo
                                   ExperimentInfo._expName,                     // 23

                                   CURVE, WEIGHTING, ORIGIN, EQUATION,
                                   R, R_SQR, NUM_SLOPE, NUM_INTERSECT,
                                   comboBox_Project.SelectedValue,              // plasma Abeta
                                   comboBox_AssayDesign.SelectedValue,                                           // assay design
                                   Math.Round(_PeakLeftRT, 2),  Math.Round(_PeakRightRT, 2),  Math.Round(_ISTDPeakLeftRT, 2),  Math.Round(_ISTDPeakRightRT, 2),
                                   _Scans,
                                    Math.Round(_IT_C13_AVG, 1),  Math.Round(_IT_C12_AVG, 1),   Math.Round(_IT_N15_AVG, 1), _SN_C13_CNT, _SN_C12_CNT, _SN_N15_CNT, _SN_TOT_CNT,
                                    _exclude == 0 ? "NULL" : _exclude.ToString(),
                                  _sampleVolume,
                                  ExperimentInfo._Instrument_ID,
                                  _Sample.Comment, _Sample.CalFile, textBox_Calib_assay.Text, _Sample.InjVol
                                   ));

                        }
                    }
                    else
                    {
                        if (_XQuanComponent.DoesDetectedQuanPeakExist > 0)
                        {
                            _XDetectedQuanPeak =
                               (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponent.DetectedQuanPeak;
                            _PeakLeftRT = _XDetectedQuanPeak.LeftRT;
                            _PeakRightRT = _XDetectedQuanPeak.RightRT;
                        }

                        if (!_ComponentName.Contains("Noise") && _Sample.FileName.Contains("Golden-Plasma") && ifDumpScanData)
                        {
                            int exported = Raw.ThermoRawFileReaderClass.DumpScanHeaderDataForCompound
                                (_extractFrom, Path.ChangeExtension(textBox_Excelfile.Text, "csv"),
                                ifheaderForDump, _ComponentName, _PeakLeftRT, _PeakRightRT, 0.5);

                            if (exported > 0) { ifheaderForDump = false; }
                            _exporSampleN++;

                        }

                    }
                }



                toolStripProgressBarMain.PerformStep();
                toolStripStatusLabelMain.Text = "Processing ..." + _Sample.FileName;
                toolStripStatusLabelMain.PerformClick();

            }
            //************************************************************************************************

            ExperimentInfo._UnknownCount = _UnknownCount;
            ExperimentInfo._StdBracketCount = _StdBracketCount;
 
            //_xlApp.Visible = true;
            toolStripProgressBarMain.Visible = false;
            toolStripProgressBarMain.Value = 0;
            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();
            toolStripProgressBarMain.Visible = true;

            toolStripProgressBarMain.Maximum = xlWB.Worksheets.Count - 1;


            string _AppPath = Application.StartupPath + Path.DirectorySeparatorChar;

            string[] _sql_P1 = File.ReadAllLines(_AppPath + "sql_TSQImport_P1.sql");
            string[] _sql_P2 = File.ReadAllLines(_AppPath + "sql_TSQImport_P2.sql");
            string[] _sql_P3 = File.ReadAllLines(_AppPath + "sql_TSQImport_P3.sql");

            _strExportSQL.Insert(0, " ");
            _strExportSQL.InsertRange(0, _sql_P1.ToList<string>());
            _strExportSQL.Insert(0, "/*"); 
            _strExportSQL.Insert(1, String.Format("Generated by XQN Reader {0} from {1}", Assembly.GetExecutingAssembly().GetName().Version.ToString(), Convert.ToDateTime(Properties.Resources.BuildDate).ToShortDateString()));
            _strExportSQL.Insert(2, String.Format("       Machine: {0} ran by {1}@{2}", Environment.MachineName, Environment.UserName, Environment.UserDomainName));
            _strExportSQL.Insert(3, String.Format("     Timestamp: {0} @ {1}", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString()));
            _strExportSQL.Insert(4, String.Format("        Source: {0} ", textBox_XQNfile.Text));
            _strExportSQL.Insert(5, String.Format("      Function: {0} ", expFunctionName));
            _strExportSQL.Insert(6, "*/");
            _strExportSQL.Add(" ");
            _strExportSQL.AddRange(_sql_P3.ToList<string>());

            if (checkBox_ExportIntoDB.Checked)
            {
                File.WriteAllLines(Path.ChangeExtension(textBox_Excelfile.Text, "sql"), _strExportSQL.ToArray());
            }

            if (checkBox_ExportOutliers.Checked)
            {
                File.AppendAllLines(Path.ChangeExtension(textBox_Excelfile.Text, "out"), _strExcludedSTD.ToArray());
            }

            foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
            {
                string lastColLetter = "AB"; int lastColIndex = 25;

                bool _IsAnalyteForDB = listAnalytesForDB.Exists(x => x == _wsComponent.Name);
                bool _IsAnalyte_ABS = listAnalytesForDB_ABS.Exists(x => x == _wsComponent.Name);
                bool _IsAnalyte_REL = listAnalytesForDB_REL.Exists(x => x == _wsComponent.Name);
               

                _wsComponent.Name = "_" + _wsComponent.Name;

                toolStripStatusLabelMain.Text = "Formatting ..." + _wsComponent.Name;
                toolStripStatusLabelMain.PerformClick();

                if (_wsComponent.Name.Contains("Summary")) { continue; }

                Excel.ChartObjects _xlCharts = (Excel.ChartObjects)_wsComponent.ChartObjects(Type.Missing);
                string _componentName = _wsComponent.get_Range("A2", "A2").Value2.ToString();

                Excel.Range _xlRange_Compound,
                            _xlRange_StdBracket_Area,
                            _xlRange_StdBracket_ISArea,
                            _xlRange_StdBracket_SpecifiedAmount,
                            _xlRange_StdBracket_Ratio,
                            _xlRange_Unknown_Area = null,
                            _xlRange_Unknown_ISArea = null,
                            _xlRange_Unknown_AreaRatio = null,
                            _xlRange_Unknown_CalcAmount = null,
                            _xlRange_Unknown_SampleID = null,
                            _xlRange_STD,
                            _xlRange_Unknowns = null,
                            _xlRange_All,
                            _xlRange_AllbutStd,
                            _xlRange_Table;
                _xlRange_Compound = _wsComponent.get_Range("A2", "A2");

                //All
                _xlRange_All =
                    _wsComponent.get_Range
                    ("A" + SkipFirstN.ToString(), lastColLetter + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                _xlRange_All.Name = _wsComponent.Name + "_All";

                _xlRange_AllbutStd =
                     _wsComponent.get_Range
                    ("A" + (SkipFirstN + _StdBracketCount).ToString(), lastColLetter + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount + _QCCount + _BlankCount).ToString());
                _xlRange_AllbutStd.Name = _wsComponent.Name + "_AllbutStd";

                _xlRange_Table =
                    _wsComponent.get_Range
                    (_wsComponent.Cells[SkipFirstN - 1, 1], _wsComponent.Cells[(SkipFirstN - 1 + _StdBracketCount + _UnknownCount + _QCCount + _BlankCount), lastColIndex + 3]);
                _xlRange_Table.Name = _wsComponent.Name + "_Table";


                string cellAddress = "A" + (ExperimentInfo.TitleHeight + ExperimentInfo._totSampleCnt + 2).ToString();

                double TimeCoursePlotOffsetY = (double)_wsCurrent.get_Range(cellAddress, cellAddress).Top;

                double TimeCoursePlotWidth = 470 * 2, TimeCoursePlotHight = 250;

                double StdsPlotOffsetY = TimeCoursePlotOffsetY + TimeCoursePlotHight;

                


                double StdsPlotWidth = 470, StdsPlotHight = 250;

                if (ExperimentInfo._StdBracketCount == 0) { TimeCoursePlotWidth = (double)_xlRange_All.Width; }

                double StdsPlotOffsetX = 0; // (double)_xlRange_All.Width - StdsPlotWidth;

                if (ExperimentInfo._UnknownCount == 0) { StdsPlotOffsetY = TimeCoursePlotOffsetY; StdsPlotOffsetX = 0; }

                //Std
                if (_StdBracketCount > 0 && _IsAnalyte_REL )
                {
                    _xlRange_STD =
                        _wsComponent.get_Range
                        ("A" + SkipFirstN.ToString(), lastColLetter + (SkipFirstN - 1 + _StdBracketCount).ToString());
                    _xlRange_STD.Name = _wsComponent.Name + "_STD";


                    _xlRange_StdBracket_Ratio =
                        _wsComponent.get_Range
                        ("F" + SkipFirstN.ToString(), "F" + (SkipFirstN - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_Ratio.Name = _wsComponent.Name + "_StdBracket_Area";

                    _xlRange_StdBracket_Area =
                        _wsComponent.get_Range
                        ("D" + SkipFirstN.ToString(), "D" + (SkipFirstN - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_Area.Name = _wsComponent.Name + "_StdBracket_Area";

                    _xlRange_StdBracket_ISArea =
                        _wsComponent.get_Range
                        ("E" + SkipFirstN.ToString(), "E" + (SkipFirstN - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_ISArea.Name = _wsComponent.Name + "_StdBracket_ISArea";

                    _xlRange_StdBracket_SpecifiedAmount =
                       _wsComponent.get_Range
                       ("G" + SkipFirstN.ToString(), "G" + (SkipFirstN - 1 + _StdBracketCount).ToString());
                    _xlRange_StdBracket_SpecifiedAmount.Name = _wsComponent.Name + "_StdBracket_SpecifiedAmount";


                    _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                    _xlRange_STD.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;
                    _xlRange_STD.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);

                    //*********Sort Range************

                    _xlRange_STD.Sort("Column AA",
                                        Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                                        "Column G", Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                                        Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                                        Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                                        Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                                        Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                                        Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                                        Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                                        Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);
                    
                    // StdCurve;
                    Excel.ChartObject _xlChartStdCurve = (Excel.ChartObject)_xlCharts.Add(
                                       StdsPlotOffsetX, StdsPlotOffsetY, StdsPlotWidth, StdsPlotHight);
                    Excel.Chart _xlChartStdCurvePage = _xlChartStdCurve.Chart;
                    _xlChartStdCurvePage.ChartType = Excel.XlChartType.xlXYScatter;

                    _xlChartStdCurvePage.HasTitle = true;
                    _xlChartStdCurvePage.ChartTitle.Font.Size = 11;

                    string _chartTitleStdCurve = String.Format
                        ("Calibration curve {0} {1} {2} {3} ({4})",
                        _componentName,
                        ExperimentInfo._IP,
                        ExperimentInfo._Enzyme,
                        ExperimentInfo._Instument,
                        ExperimentInfo._Date);
                    //_xlSeries_TTR.Name = _componentName;

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
                    _xlSerie_Stds.XValues = _xlRange_StdBracket_SpecifiedAmount.Cells;
                    _xlSerie_Stds.Values = _xlRange_StdBracket_Ratio.Cells;


                    string _serieStdName = "H:L Abeta";

                    string _axeStdXNumFormat = "0.00",
                           _axeStdYNumFormat = "0.00";

                    string _axeStdXCaption = "%, Predicted labeling Abeta",
                           _axeStdYCaption = "%, Measured labeling Abeta";

                    if (ExperimentInfo._quanType == "ABS")
                    {
                        _axeStdXNumFormat = "0.0";
                        _axeStdXNumFormat = "0.00";

                        _axeStdXCaption = "ng/mL, Specified amount";
                        _axeStdYCaption = "AreaRatio";

                        _serieStdName = "Analyte:ISTD Ratio";
                    }

                    _xlAxisX_Std.TickLabels.NumberFormat = _axeStdXNumFormat;
                    _xlAxisY_Std.TickLabels.NumberFormat = _axeStdYNumFormat;
                    _xlAxisX_Std.AxisTitle.Caption = _axeStdXCaption;
                    _xlAxisY_Std.AxisTitle.Caption = _axeStdYCaption;

                    _xlSerie_Stds.Name = _serieStdName;

                    Excel.Trendlines _xlTrendlines_Stds = (Excel.Trendlines)_xlSerie_Stds.Trendlines(Type.Missing);

                    Excel.XlTrendlineType _XlTrendlineType =
                        Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear;
                    string _CurveIndex = _wsComponent.get_Range("C2", "C2").Value2.ToString();

                    if (_CurveIndex == "Quadratic")
                    { _XlTrendlineType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlPolynomial; }


                    _xlTrendlines_Stds.Add(_XlTrendlineType, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, true, true, Type.Missing);

                    //****************

                };


                //Unknowns
                if (_UnknownCount > 0)
                {

                    _xlRange_Unknowns =
                        _wsComponent.get_Range
                        ("A" + (SkipFirstN + _StdBracketCount).ToString(), lastColLetter + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                    _xlRange_Unknowns.Name = _wsComponent.Name + "_Unknowns";

                    _xlRange_Unknown_Area =
                        _wsComponent.get_Range
                        ("D" + (SkipFirstN + _StdBracketCount).ToString(), "D" + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                    _xlRange_Unknown_Area.Name = _wsComponent.Name + "_Unknown_Area";

                    _xlRange_Unknown_ISArea =
                        _wsComponent.get_Range
                        ("E" + (SkipFirstN + _StdBracketCount).ToString(), "E" + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                    _xlRange_Unknown_ISArea.Name = _wsComponent.Name + "_Unknown_ISArea";

                    string strCalcOrRAtio = "H";
                    if (_StdBracketCount == 0  && _IsAnalyte_ABS ) { strCalcOrRAtio = "F"; }
                    _xlRange_Unknown_CalcAmount =
                        _wsComponent.get_Range
                        (strCalcOrRAtio + (SkipFirstN + _StdBracketCount).ToString(), strCalcOrRAtio + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                    _xlRange_Unknown_CalcAmount.Name = _wsComponent.Name + "_Unknown_CalcAmount";

                     _xlRange_Unknown_AreaRatio =
                        _wsComponent.get_Range
                        ("F" + (SkipFirstN + _StdBracketCount).ToString(), "F" + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                    _xlRange_Unknown_AreaRatio.Name = _wsComponent.Name + "_Unknown_AreaRatio";

                    _xlRange_Unknown_SampleID =
                        _wsComponent.get_Range
                        ("J" + (SkipFirstN + _StdBracketCount).ToString(), "J" + (SkipFirstN - 1 + _StdBracketCount + _UnknownCount).ToString());
                    _xlRange_Unknown_SampleID.Name = _wsComponent.Name + "_Unknown_SampleID";
                }


                //*************************Format Ranges**************************
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).Weight = 3.75;
                _xlRange_All.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = 3.75;


                //******Sort Unknowns

                _xlRange_AllbutStd.Sort("Column B",
                    Microsoft.Office.Interop.Excel.XlSortOrder.xlDescending,
                    Type.Missing, Type.Missing,
                    Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                    Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                    Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                    Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                    Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);

                if (_UnknownCount > 0)
                {
                    _xlRange_Unknowns.Sort("Column AA",
                        Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                        "Column J", Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                        Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending,
                        Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                        Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns,
                        Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin,
                        Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                        Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal,
                        Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);


                    Excel.ColorScale cfColorScale_forISArea = (Excel.ColorScale)(_xlRange_Unknown_ISArea.FormatConditions.AddColorScale(3));

                    cfColorScale_forISArea.ColorScaleCriteria[1].FormatColor.Color = 7039480;  //0x000000FF;  
                    cfColorScale_forISArea.ColorScaleCriteria[2].FormatColor.Color = 8711167;  // 0x00FF0000;
                    cfColorScale_forISArea.ColorScaleCriteria[3].FormatColor.Color = 13011546; // 0x00FF0000;

                    Excel.ColorScale cfColorScale_forArea = (Excel.ColorScale)(_xlRange_Unknown_Area.FormatConditions.AddColorScale(3));

                    cfColorScale_forArea = cfColorScale_forISArea;




                    //*************************Charts*****************************************************************************************************
                    //******* Unknowns************

                    string _serieCurveName = null; string _strYlabelREL = "TTR, %"; string _strYlabelABS = "Area Ratio"; 

                    string _axeCurveXNumFormat = null, _axeCurveYNumFormat = null;

                    string _axeCurveXCaption = null,  _axeCurveYCaption = null;

                    Excel.Range _UnknownLalues = null;

                    if (_IsAnalyte_REL)
                    {
                        _serieCurveName = "H:L Abeta";

                        _axeCurveXNumFormat = "0"; _axeCurveYNumFormat = "0.0%";

                        _axeCurveXCaption = "h, Time"; _axeCurveYCaption = _strYlabelREL;

                        _UnknownLalues = _xlRange_Unknown_CalcAmount;
                    }
                    if (_IsAnalyte_ABS)
                    {
                        _serieCurveName = "Abeta level";

                        _axeCurveYNumFormat = "0.0";

                        _axeCurveXCaption = "h, Time"; _axeCurveYCaption = _strYlabelABS;
                        
                        _UnknownLalues = _xlRange_Unknown_AreaRatio;
                    }

                    Excel.ChartObject _xlChartTimeCorse = (Excel.ChartObject)_xlCharts.Add(
                                       0, TimeCoursePlotOffsetY, TimeCoursePlotWidth, TimeCoursePlotHight);
                    Excel.Chart _xlChartTimeCorsePage = _xlChartTimeCorse.Chart;
                    _xlChartTimeCorsePage.ChartType = Excel.XlChartType.xlXYScatter;
                    _xlChartTimeCorsePage.HasTitle = true; _xlChartTimeCorsePage.ChartTitle.Font.Size = 11;

                    _xlChartTimeCorsePage.ChartTitle.Caption = _componentName + ExperimentInfo._CommonChartNameStr;

                    Excel.Axes _xlAxes_TP = (Excel.Axes)_xlChartTimeCorsePage.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisX_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Excel.Axis _xlAxisY_TP = _xlAxes_TP.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                    _xlAxisX_TP.HasMajorGridlines = true; _xlAxisX_TP.HasTitle = true; _xlAxisY_TP.HasTitle = true;

                    _xlChartTimeCorsePage.Legend.Position = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner;
                    _xlChartTimeCorsePage.Legend.IncludeInLayout = false;

                    Excel.SeriesCollection _xlSeriesColl_TimeCourse = (Excel.SeriesCollection)_xlChartTimeCorsePage.SeriesCollection(Type.Missing);

                    Excel.Series _xlSeries_TTR = (Excel.Series)_xlSeriesColl_TimeCourse.NewSeries();
                    _xlSeries_TTR.XValues = _xlRange_Unknown_SampleID.Cells; _xlSeries_TTR.Values = _UnknownLalues;


                    if (_StdBracketCount == 0  && _IsAnalyte_ABS) { _axeCurveYCaption = "Area Ratio"; }

                    _xlAxisX_TP.TickLabels.NumberFormat = _axeCurveXNumFormat; _xlAxisY_TP.TickLabels.NumberFormat = _axeCurveYNumFormat;
                    _xlAxisX_TP.AxisTitle.Caption = _axeCurveXCaption; _xlAxisY_TP.AxisTitle.Caption = _axeCurveYCaption;

                    _xlSeries_TTR.Name = _serieCurveName; _xlSeries_TTR.Name = _componentName;




                    Excel.ChartObjects _xlChartsSummary = (Excel.ChartObjects)_wsSummary.ChartObjects(Type.Missing);
                    Excel.ChartObject _xlChartCurveSummaryREL; Excel.ChartObject _xlChartCurveSummaryABS; Excel.ChartObject _xlChartCurveSummary = null;
                    Excel.Chart _xlChartCurveSummaryPageREL; Excel.Chart _xlChartCurveSummaryPageABS; Excel.Chart _xlChartCurveSummaryPage = null;

                    if (_xlChartsSummary.Count < 1)
                    {
                        _xlChartCurveSummaryREL = (Excel.ChartObject)_xlChartsSummary.Add(
                                       10, 10, 850, 500);
                        _xlChartCurveSummaryABS = (Excel.ChartObject)_xlChartsSummary.Add(
                                       10, 510, 850, 500);


                        _xlChartCurveSummaryPageREL = _xlChartCurveSummaryREL.Chart; _xlChartCurveSummaryPageREL.ChartType = Excel.XlChartType.xlXYScatter;
                        _xlChartCurveSummaryPageREL.HasTitle = true; _xlChartCurveSummaryPageREL.ChartTitle.Font.Size = 11;

                        _xlChartCurveSummaryPageABS = _xlChartCurveSummaryABS.Chart; _xlChartCurveSummaryPageABS.ChartType = Excel.XlChartType.xlXYScatter;
                        _xlChartCurveSummaryPageABS.HasTitle = true; _xlChartCurveSummaryPageABS.ChartTitle.Font.Size = 11;

                        Excel.Axes _xlAxes_CurveSummaryREL = (Excel.Axes)_xlChartCurveSummaryPageREL.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);
                        Excel.Axes _xlAxes_CurveSummaryABS = (Excel.Axes)_xlChartCurveSummaryPageABS.Axes(Type.Missing, Excel.XlAxisGroup.xlPrimary);

                        Excel.Axis _xlAxisX_CurveSummaryREL = _xlAxes_CurveSummaryREL.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                        Excel.Axis _xlAxisY_CurveSummaryREL = _xlAxes_CurveSummaryREL.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                        Excel.Axis _xlAxisX_CurveSummaryABS = _xlAxes_CurveSummaryABS.Item(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                        Excel.Axis _xlAxisY_CurveSummaryABS = _xlAxes_CurveSummaryABS.Item(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                        _xlAxisX_CurveSummaryREL.HasMajorGridlines = true; _xlAxisY_CurveSummaryREL.HasMajorGridlines = true;

                        _xlAxisX_CurveSummaryREL.Format.Line.Weight = 3.0F; _xlAxisY_CurveSummaryREL.Format.Line.Weight = 3.0F;

                        _xlAxisX_CurveSummaryREL.HasTitle = true; _xlAxisY_CurveSummaryREL.HasTitle = true;

                        _xlAxisX_CurveSummaryREL.AxisTitle.Caption = _axeCurveXCaption; _xlAxisY_CurveSummaryREL.AxisTitle.Caption = _strYlabelREL;

                        _xlAxisX_CurveSummaryREL.TickLabels.NumberFormat = _axeCurveXNumFormat; _xlAxisY_CurveSummaryREL.TickLabels.NumberFormat = _axeCurveYNumFormat;


                        _xlAxisX_CurveSummaryABS.HasMajorGridlines = true; _xlAxisY_CurveSummaryABS.HasMajorGridlines = true;

                        _xlAxisX_CurveSummaryABS.Format.Line.Weight = 3.0F; _xlAxisY_CurveSummaryABS.Format.Line.Weight = 3.0F;

                        _xlAxisX_CurveSummaryABS.HasTitle = true; _xlAxisY_CurveSummaryABS.HasTitle = true;

                        _xlAxisX_CurveSummaryABS.AxisTitle.Caption = _axeCurveXCaption; _xlAxisY_CurveSummaryABS.AxisTitle.Caption = _strYlabelABS;

                        _xlAxisX_CurveSummaryABS.TickLabels.NumberFormat = _axeCurveXNumFormat; _xlAxisY_CurveSummaryABS.TickLabels.NumberFormat = "0.0";


                        _xlChartCurveSummaryPageREL.ChartTitle.Caption = "TTR, % ";
                        _xlChartCurveSummaryPageABS.ChartTitle.Caption = "Area Ratio ";



                        _xlChartCurveSummaryPageREL.ChartTitle.Caption += ExperimentInfo._CommonChartNameStr;
                        _xlChartCurveSummaryPageREL.Legend.Position = Excel.XlLegendPosition.xlLegendPositionRight;
                        _xlChartCurveSummaryPageREL.Legend.IncludeInLayout = true;

                        _xlChartCurveSummaryPageABS.ChartTitle.Caption += ExperimentInfo._CommonChartNameStr;
                        _xlChartCurveSummaryPageABS.Legend.Position = Excel.XlLegendPosition.xlLegendPositionRight;
                        _xlChartCurveSummaryPageABS.Legend.IncludeInLayout = true;

                      


                    }


                   

                    if (_IsAnalyte_REL)
                    {
                        _xlChartCurveSummary = (Excel.ChartObject)_xlChartsSummary.Item(1);
                        _UnknownLalues = _xlRange_Unknown_CalcAmount;
                    }
                    if (_IsAnalyte_ABS)
                    {
                        _xlChartCurveSummary = (Excel.ChartObject)_xlChartsSummary.Item(2);
                        _UnknownLalues = _xlRange_Unknown_AreaRatio;
                    }

                    _xlChartCurveSummaryPage = _xlChartCurveSummary.Chart;


                    Excel.SeriesCollection _xlSeriesColl_CurveSummary = (Excel.SeriesCollection)_xlChartCurveSummaryPage.SeriesCollection(Type.Missing);

                    Excel.Series _xlSeries_CurveTTR = (Excel.Series)_xlSeriesColl_CurveSummary.NewSeries();
                    _xlSeries_CurveTTR.XValues = _xlRange_Unknown_SampleID.Cells;
                    _xlSeries_CurveTTR.Values = _UnknownLalues;
                    _xlSeries_CurveTTR.Name = (string)_xlRange_Compound.Cells.Value2;

                }     

                toolStripProgressBarMain.PerformStep();

            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = false;

            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();

            _XXQN.Close();

            object misValue = System.Reflection.Missing.Value;

            try
            {
                _XXQN.Close();

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
                    SortExclelWorkSheets(xlWB.Worksheets);

                    _xlApp.Visible = true;
                    ((Excel._Worksheet)_wsSummary).Activate();

                }
                else
                {
                    _xlApp.Quit();
                }
            }
            if (this.checkBox_ShowDebug.Checked)
            {
                MessageBox.Show( expFunctionName);
            }
        }



        //*******************************2017-03-07
        private void ExportXcalibur_ForNico()
        {
            string expFunctionName = "ExportXcalibur_ForNico()";
            string excelAnalyteTemplate = Application.StartupPath + Path.DirectorySeparatorChar + @"\Templates\QuanBrow_Universal_Analyte_Template.xlsx";
            string excelSummaryTemplate = Application.StartupPath + Path.DirectorySeparatorChar + @"\Templates\QuanBrow_Universal_Summary_Template.xlsx";

   
   
            string _strAnalyte = null;
            double _Area = 0, _ISTD_Area = 0, _AreaRatio = 0, _CalcAmount = 0;
            decimal dcm_AreaRatio = 0;
            string _strSpecAmount = null, _strRT = null, _strISTD_RT = null;
            string _response = null;

            if (textBox_XQNfile.Text == "")
            {
                MessageBox.Show("Select Input XQN file, please");
                return;
            }

            _XXQN = new XCALIBURFILESLib.XXQNClass();

            try
            {
                _XXQN.Open(textBox_XQNfile.Text);
            }
            catch (Exception _exception)
            {
                MessageBox.Show("File: " + textBox_XQNfile.Text + " is invalid XQN-file", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                if (checkBox_ShowDebug.Checked)
                {
                    MessageBox.Show(_exception.Message);
                }
                return;
            }
            finally
            {

            }


            XCALIBURFILESLib.IXHeader _IXHeader = (XCALIBURFILESLib.IXHeader)_XXQN.Header; //MessageBox.Show(_XXQN.FileName);

            ExperimentInfo._ProcessedBy = _IXHeader.ChangedId;

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = _IXHeader.ChangedLogon;
            }

            if (ExperimentInfo._ProcessedBy == "")
            {
                ExperimentInfo._ProcessedBy = "unknown";
            }

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = true;
            toolStripStatusLabelMain.Visible = true;

            _xlApp = new Excel.ApplicationClass();
            _xlApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;

            Excel.Workbook xlWB;




            try
            {

                xlWB =
                    _xlApp.Workbooks.Add( /*Excel.XlWBATemplate.xlWBATWorksheet*/);

            }

            catch (Exception)
            {
                MessageBox.Show("File: " + excelSummaryTemplate + " causing Anchor exeption", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {

            }

      

            //Excel.Worksheet _wsSummary = (Excel.Worksheet)xlWB.ActiveSheet;

            Excel.Worksheet _wsCurrent = (Excel.Worksheet)xlWB.ActiveSheet; //null;

            byte SkipFirstN = ExperimentInfo.TitleHeight; /*Title height of template in rows*/

            byte _UnknownCount = 0, _StdBracketCount = 0, _BlankCount = 0, _QCCount = 0, _TotalCount = 0;


            XCALIBURFILESLib.XBracketGroups _XBracketGroups
                = (XCALIBURFILESLib.XBracketGroups)_XXQN.BracketGroups;

            XCALIBURFILESLib.XBracketGroup _XBracketGroup
                = (XCALIBURFILESLib.XBracketGroup)_XBracketGroups.Item(1);

            XCALIBURFILESLib.XQuanResults _XQuanResults
                = (XCALIBURFILESLib.XQuanResults)_XBracketGroup.QuanResults;



            int _curSampleN = 0;
            //********_XQuanResults*******************************
            foreach (XCALIBURFILESLib.IXQuanResult /*.XQuanResult*/ _XQuanResult in _XQuanResults)
            {
                toolStripProgressBarMain.Maximum = _XQuanResults.Count;

                XCALIBURFILESLib.XSample _Sample
                    = (XCALIBURFILESLib.XSample)_XQuanResult.Sample;


                string _extractFrom = Path.GetDirectoryName(textBox_XQNfile.Text) + "\\" + _Sample.FileName + ".raw";


                ExperimentInfo._totSampleCnt++;

                _TotalCount++;
                switch (_Sample.SampleType)
                {
                    case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                        _UnknownCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                        _StdBracketCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleBlank:
                        _BlankCount++;
                        break;
                    case XCALIBURFILESLib.XSampleTypes.XSampleQC:
                        _QCCount++;
                        break;

                    default:
                        MessageBox.Show("Unusual sample type " + _Sample.SampleType.ToString());
                        break;
                }


                XCALIBURFILESLib.XQuanComponents _XQuanSelectedComponents = null;


                try
                {
                    _XQuanSelectedComponents
                       = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Selected;

                }
                catch (Exception ex)
                { MessageBox.Show(_Sample.FileName); }


                XCALIBURFILESLib.XQuanComponents _XQuanMethodComponents
                    = (XCALIBURFILESLib.XQuanComponents)_XQuanResult.Method;


                short _comNo = 1;



                //_XQuanSelectedComponents.

                var _varlistOfCompounds = from XCALIBURFILESLib.XQuanComponent _com in _XQuanSelectedComponents.OfType<XCALIBURFILESLib.XQuanComponent>() // .AsQueryable()
                                                select  
                                                
                                                
                                                ((XCALIBURFILESLib.XComponent) _com.Component).Name; //.ToString().ToList();



               // foreach (string s in _varlistOfCompounds)
               // {
 
               // }


                List<string> _listOfCompounds = _varlistOfCompounds.ToList<string>() ;   //.ToList<string>();



                foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponent in _XQuanSelectedComponents)
                {

                    try
                    {

                        XCALIBURFILESLib.XComponent _XComponent =
                            (XCALIBURFILESLib.XComponent)_XQuanComponent.Component;

                        XCALIBURFILESLib.XQuanComponent _XQuanMethodComponent = (XCALIBURFILESLib.XQuanComponent)
                        _XQuanMethodComponents.Item(_comNo);
                        _comNo++;

                        XCALIBURFILESLib.XComponent _XMethodComponent =
                            (XCALIBURFILESLib.XComponent)_XQuanMethodComponent.Component;

                        XCALIBURFILESLib.IXDetection2 _IXDetection =
                            (XCALIBURFILESLib.IXDetection2)_XComponent.Detection;

                        XCALIBURFILESLib.XQuanCalibration _XQuanCalibration =
                                (XCALIBURFILESLib.XQuanCalibration)_XComponent.Calibration;

                        XCALIBURFILESLib.XQuanCalibration _XQuanMethodCalibration =
                                (XCALIBURFILESLib.XQuanCalibration)_XMethodComponent.Calibration;




                        string _ComponentName = _XComponent.Name;

                        double _PeakLeftRT = 0.0, _PeakRightRT = 0.0;

                        XCALIBURFILESLib.XDetectedQuanPeak _XDetectedQuanPeak;

                        int _CompIndex = _listOfCompounds.IndexOf(_ComponentName);


                        _strAnalyte = _XComponent.Name.Replace("_", " ");


                        int len = (Path.GetFileNameWithoutExtension(_extractFrom)).Length;
                        _wsCurrent.Name = Path.GetFileNameWithoutExtension(_extractFrom).Substring(1, 4/*, len > 31 ? 31 : len*/);

                       /* string stampFormat = "Generated by XQN Reader {0} from {1}; " +
                         "-Function {2}; " +
                         "-Machine: {3} ran by {4}@{5}; " +
                         "-Timestamp: {6} @ {7}; " +
                         "-Source: {8}; " +
                         "-Analyte template {9}; " +
                         "-Summary template {10} ";*/

                       /* string stamp = String.Format(stampFormat,
                                      Assembly.GetExecutingAssembly().GetName().Version.ToString(), Convert.ToDateTime(Properties.Resources.BuildDate).ToShortDateString(),
                                      expFunctionName,
                                      Environment.MachineName, Environment.UserName, Environment.UserDomainName,
                                      DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString(),
                                      textBox_XQNfile.Text,
                                      Path.GetFileName(excelAnalyteTemplate),
                                      Path.GetFileName(excelSummaryTemplate)
                                      );*/

                        //   Excel.Comment GenerateStamp = (_wsCurrent.Cells[3, 1] as Excel.Range).AddComment(stamp);





                        int _currRowN = _Sample.RowNumber + SkipFirstN;


                        if (_Sample.RowNumber == 1)
                        {
                            //foreach (string _comName in _listOfCompounds)
                            //{
                            _wsCurrent.Cells[SkipFirstN - 1, _CompIndex + 2] = _ComponentName;
                            //}
                        }


                        if (_CompIndex == 1)
                        {
                            _wsCurrent.Cells[_currRowN, 1] = _Sample.FileName;
                        }




                        _XDetectedQuanPeak = null;


                        if (_XQuanComponent.DoesDetectedQuanPeakExist > 0)
                        {

                            _XDetectedQuanPeak =
                                (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponent.DetectedQuanPeak;


                            _wsCurrent.Cells[_currRowN, _CompIndex + 2] = _XDetectedQuanPeak.Area;

                            /*     if (_XDetectedQuanPeak.ISTDArea != 0)
                                 {
                                 }
                                 else
                                 {
                                 }*/


                            _response = "NULL"; _strRT = "NULL"; _strISTD_RT = "NULL";


                        }
                        else
                        {

                        }


                        /*     switch (_Sample.SampleType)
                             {
                                 case XCALIBURFILESLib.XSampleTypes.XSampleStdBracket:
                                     break;
                                 case XCALIBURFILESLib.XSampleTypes.XSampleUnknown:
                                     break;

                                 case XCALIBURFILESLib.XSampleTypes.XSampleBlank:
                                     break;

                                 case XCALIBURFILESLib.XSampleTypes.XSampleQC:
                                     break;

                                 default:
                                     break;
                             } 


                             foreach (XCALIBURFILESLib.XQuanComponent _XQuanComponentISTD in _XQuanSelectedComponents)
                             {
                                 XCALIBURFILESLib.XComponent _XISTD =
                                 (XCALIBURFILESLib.XComponent)_XQuanComponentISTD.Component;

                                 if (_XISTD.Name == _XQuanCalibration.ISTD)
                                 {
                                     if (_XQuanComponentISTD.DoesDetectedQuanPeakExist > 0)
                                     {
     
                                     }
                                     else
                                     {
                                     }

                                 }

                             } 

                    
                             if (_XQuanComponent.DoesDetectedQuanPeakExist > 0)
                             {
                                 _XDetectedQuanPeak =
                                    (XCALIBURFILESLib.XDetectedQuanPeak)_XQuanComponent.DetectedQuanPeak;
                                 _PeakLeftRT = _XDetectedQuanPeak.LeftRT;
                                 _PeakRightRT = _XDetectedQuanPeak.RightRT;
                             }*/

                        //   MessageBox.Show(_Sample.FileName, _ComponentName); 

                    }
                    catch (COMException _exception)
                    {}
                }



                toolStripProgressBarMain.PerformStep();
                toolStripStatusLabelMain.Text = "Processing ..." + _Sample.FileName;
                toolStripStatusLabelMain.PerformClick();

            }



            //_xlApp.Visible = true;
            toolStripProgressBarMain.Visible = false;
            toolStripProgressBarMain.Value = 0;
            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();
            toolStripProgressBarMain.Visible = true;

            toolStripProgressBarMain.Maximum = xlWB.Worksheets.Count - 1;


            string _AppPath = Application.StartupPath + Path.DirectorySeparatorChar;

 
            foreach (Excel.Worksheet _wsComponent in xlWB.Worksheets)
            { };
 
            

            toolStripProgressBarMain.Value = 0;
            toolStripProgressBarMain.Visible = false;

            toolStripStatusLabelMain.Text = "Done";
            toolStripStatusLabelMain.PerformClick();

            _XXQN.Close();

            object misValue = System.Reflection.Missing.Value;

            try
            {
                _XXQN.Close();

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
                    SortExclelWorkSheets(xlWB.Worksheets);

                    _xlApp.Visible = true;

                }
                else
                {
                    _xlApp.Quit();
                }
            }
            if (this.checkBox_ShowDebug.Checked)
            {
                MessageBox.Show(expFunctionName);
            }
        }



        //*******************************2016-05-22

        private void SortExclelWorkSheets( Excel.Sheets worksheets)
        {
            List<string> listSortOrder1 = new List<string>()
                {
                    "_Aß38_Oxy_N14_N15", "_AßMD_N14_N15", "_Aß40_Oxy_N14_N15", "_Aß42_Oxy_N14_N15",
                    "_Aß38_Oxy_C13_C12", "_AßMD_C13_C12", "_Aß40_Oxy_C13_C12", "_Aß42_Oxy_C13_C12",
                    "_Summary"
                };

            listSortOrder1.Reverse();

            foreach (string _wsName in listSortOrder1)
            {
                Excel.Worksheet _wsCurrent = null;
                bool _found = false;
                foreach (Excel.Worksheet _ws in worksheets)
                {
                    if (_ws.Name == _wsName)
                    {
                        _wsCurrent = _ws;
                        _found = true;
                    }
                }

                if (_found)
                {
                    _wsCurrent.Move(worksheets[1]);
                }
                
            }

        }


        private void extractMethodFromRAW(object sender, EventArgs e)
        {
            OpenFileDialog _openDlg = new OpenFileDialog();
            _openDlg.Filter = "RAW|*.RAW";
            _openDlg.Multiselect = true;

            if (_openDlg.ShowDialog() == DialogResult.OK)
            {
              //  int exported = Raw.ThermoRawFileReaderClass.DumpInstrumentMethod(_openDlg.FileName);

                Raw.ThermoRawFileReaderClass.DumpScanHeaderData(_openDlg.FileName, Path.GetFileNameWithoutExtension(_openDlg.FileName) + ".dmD", true);

            }

        }

        private void toolStripMenuItem_Change_log_Click(object sender, EventArgs e)
        {
            string _AppPath = Application.StartupPath + Path.DirectorySeparatorChar;

            Process.Start(_AppPath + "ChangeLog.pdf");
        }

        private void label13_Click(object sender, EventArgs e)
        {
            OpenFileDialog _openDlg = new OpenFileDialog();
            _openDlg.Filter = "XQN|*.XQN";
            _openDlg.Multiselect = false;

            if (_openDlg.ShowDialog() == DialogResult.OK)
            {
                textBox_Calib_assay.Text = Path.GetFileNameWithoutExtension(_openDlg.FileName);
            }

        }

    }

    

}
