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
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GCMSReader
{
    public partial class fmMain : Form
    {

        private string _sqlFileName;

        private byte _fluid_id;

        private List<string> _exportList;

        private Excel.Application _xlApp;

        public fmMain()
        {
            InitializeComponent();
        }

        private Int32 Get_TP(double _TP)
        {
            Int32 _TimePointID = -1;

            try
            {
                var query_TP_ID = tIME_POINTTableAdapter.GetData().Select("TIME_POINT_VALUE_NUM = "
                                                                      + _TP.ToString()
                                                                      + "AND STUDY_ID = "
                                                                      + comboBox_Study.SelectedValue);

                _TimePointID = Convert.ToInt32(query_TP_ID[0]["TIME_POINT_ID"]);
           
            }
            catch (Exception _err)
            {
                MessageBox.Show(_err.Message + ": " + _TP.ToString(), "Can't convert to Integer");
            }

            return _TimePointID;
        }

        private Int32 Get_TP1(string _TP)
        {
            Int32 _TimePointID = -1;

            try
            {
                var query_TP_ID = tIME_POINTTableAdapter.GetData().Select("TIME_POINT_LABEL LIKE '" + _TP.ToString() + "%' "
                                                                      + "AND STUDY_ID = "
                                                                      + comboBox_Study.SelectedValue);

                _TimePointID = Convert.ToInt32(query_TP_ID[0]["TIME_POINT_ID"]);

            }
            catch (Exception _err)
            {
                MessageBox.Show(_err.Message + ": " + _TP.ToString(), "Can't convert to Integer");
            }

            return _TimePointID;
        }

        private Int32 Get_TP_TAU(string _TP)  // GC Yarosheski / Tau
        {
            Int32 _TimePointID = -1;

            try
            {
                var query_TP_ID = tIME_POINTTableAdapter.GetData().Select("TIME_POINT_VALUE_NUM = " + _TP + " "
                                                                      + "AND STUDY_ID = "
                                                                      + comboBox_Study.SelectedValue);

                _TimePointID = Convert.ToInt32(query_TP_ID[0]["TIME_POINT_ID"]);

            }
            catch (Exception _err)
            {
                MessageBox.Show(_err.Message + ": " + _TP.ToString(), "Can't convert to Integer");
            }

            return _TimePointID;
        }

        private void button_StartExport_Click(object sender, EventArgs e)
        {
            if (textBox_GCMSfile.Text == "") { return; } ;

                         
                _xlApp = new Excel.ApplicationClass();
                _xlApp.ReferenceStyle = Excel.XlReferenceStyle.xlA1;

                Excel.Workbook _xlWBook =
                  _xlApp.Workbooks.Open
                  (textBox_GCMSfile.Text, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);


               // if (radioButton_hCSF.Checked) { _fluid_id = 1; }                    /*hCSF*/
               // if (radioButton_hPlasma.Checked) { _fluid_id = 2; }                 /*hPlasma*/
               // if (radioButton_rmCSF.Checked) { _fluid_id = 3; }                   /*rmCSF*/
               // if (radioButton_rmPlasma.Checked) { _fluid_id = 4; }                /*rmPlasma*/

                _fluid_id = Convert.ToByte(comboBox_FluidType.SelectedValue);

                string _AppPath = Application.StartupPath + Path.DirectorySeparatorChar;
                string[] _sql_Part1 = null;
                string[] _sql_Part3 = null; 


                if (radioButton_YarasheskiLab.Checked)
                  {
                    YarosheskiExport(ref _xlWBook, _fluid_id);
                    _sql_Part1 = File.ReadAllLines(_AppPath + "sql_GCMSImport_Yar_P1.sql");
                    _sql_Part3 = File.ReadAllLines(_AppPath + "sql_GCMSImport_Yar_P3.sql");
                  }

                if (radioButton_YarasheskiLab_TAU.Checked)
                  {
                    YarosheskiExportTAU(ref _xlWBook, _fluid_id);
                    _sql_Part1 = File.ReadAllLines(_AppPath + "sql_GCMSImport_YarTAU_P1.sql");
                    _sql_Part3 = File.ReadAllLines(_AppPath + "sql_GCMSImport_YarTAU_P3.sql");
                  }

            if (radioButton_YarasheskiLab_BACE.Checked)
                {
                    YarosheskiExportBASE(ref _xlWBook, _fluid_id);
                    _sql_Part1 = File.ReadAllLines(_AppPath + "sql_GCMSImport_YarBACE_P1.sql");
                    _sql_Part3 = File.ReadAllLines(_AppPath + "sql_GCMSImport_YarBACE_P3.sql");
                }
                if (radioButton_PattersonLab.Checked)
                  {
                    PattersonExport(ref _xlWBook, _fluid_id);
                    _sql_Part1 = File.ReadAllLines(_AppPath + "sql_GCMSImport_Pat_P1.sql");
                    _sql_Part3 = File.ReadAllLines(_AppPath + "sql_GCMSImport_Pat_P3.sql");
                  }
                
            
            _exportList.Insert(0, " ");
            _exportList.InsertRange(0, _sql_Part1.ToList<string>());

            //_exportList.Insert(0, String.Format("/* Generated by GCMSReader on {0} @ {1} */", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString()));

            _exportList.Insert(0, "/*");

            _exportList.Insert(1, String.Format("Generated by GCMC Reader {0} from {1}", Assembly.GetExecutingAssembly().GetName().Version.ToString(), Convert.ToDateTime(Properties.Resources.BuildDate).ToShortDateString()));
            _exportList.Insert(2, String.Format("       Machine: {0} ran by {1}@{2}", Environment.MachineName, Environment.UserName, Environment.UserDomainName));
            _exportList.Insert(3, String.Format("     Timestamp: {0} @ {1}", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString()));
            _exportList.Insert(4, String.Format("        Source: {0} ", textBox_GCMSfile.Text));

            _exportList.Insert(5, "*/");
            
            _exportList.AddRange(_sql_Part3.ToList<string>());

            //_exportList.AddRange(_sql_Part1.ToList<string>());

            _exportList.Add("COMMIT WORK;");

            File.WriteAllLines(_sqlFileName, _exportList.ToArray() );

            MessageBox.Show( this, "Done", "GCMS Reader", MessageBoxButtons.OK, MessageBoxIcon.Information );
            
            
            _xlApp.Quit();

            releaseObject(_xlWBook);
            releaseObject(_xlApp);
            

        }

        private void YarosheskiExportBASE(ref Excel.Workbook _xlWB, byte _fluid_id)
        {
            _exportList = new List<string>();

            foreach (Excel.Worksheet _xlWS in _xlWB.Worksheets)
            {
                double _mp0 = 0;

                for (int i = 18; i < 62; i++)
                {
                    Int32 _dateOffset = 9;
                    if ((_xlWS.Cells[1, 9] as Excel.Range).Value2 == null) { _dateOffset = 12; }
                    string _assay_date = DateTime.FromOADate((double)(_xlWS.Cells[1, _dateOffset] as Excel.Range).Value2).Date.ToString("yyyy-MM-dd");
                    string _subject = _xlWS.Name.Replace("p", "");


                    if ((_xlWS.Cells[i, 1] as Excel.Range).Value2 != null)
                    {
                        double _timepoint = 0;
                        double _349;
                        double _355;
                        double _355_349;
                        double _ttr;
                        double _mp;
                        string _sample_name;
                        try
                        {
                            _timepoint = (double)(_xlWS.Cells[i, 1] as Excel.Range).Value2;
                            _349 = (double)(_xlWS.Cells[i, 2] as Excel.Range).Value2;
                            _355 = (double)(_xlWS.Cells[i, 3] as Excel.Range).Value2;
                            _355_349 = (double)(_xlWS.Cells[i, 4] as Excel.Range).Value2;
                            _ttr = (double)(_xlWS.Cells[i, 5] as Excel.Range).Value2;
                            _mp = _355 / (_355 + _349);
                            _sample_name = (string)(_xlWS.Cells[i, 6] as Excel.Range).Value2;
                        }
                        catch
                        {
                            MessageBox.Show(String.Format
                                ("Kelly Moor says: There is a problem with subject {0}, hr {1}", _subject, _timepoint),
                                 "Missing value");
                            continue;
                        }

                        if (_timepoint == -1.0) { _mp0 = _mp; };
                        double _mpe = (_mp - _mp0) * 100;

                        string _insert = "INSERT INTO TMP_GC_MS_IMPORT " +
                        "(SUBJECT_NUM, FLUID_TYPE_ID, ASSAY_DATE, TP_ID, NUM_349, NUM_355, MP, MPE, NUM355_349, TTR, SAMPLE_NAME) ";
                        string _values = String.Format("VALUES ('{0}', {1}, '{2}', {3}, {4}, {5}, {6}, {7}, {8}, {9}, '{10}');",
                               _subject,
                               _fluid_id,
                               _assay_date,
                               _timepoint,
                               _349,
                               _355,
                               _mp, _mpe,
                               _355_349,
                               _ttr,
                               _sample_name);

                        _exportList.Add(_insert);
                        _exportList.Add(_values);

                    }
                }
            }

        }


        private void YarosheskiExport(ref Excel.Workbook _xlWB, byte _fluid_id )
        {
            _exportList = new List<string>();

            foreach (Excel.Worksheet _xlWS in _xlWB.Worksheets)
            {
                double _mp0 = 0;

                Excel.Range _xlRange_AssayDate = _xlWS.get_Range(textBox_AssayDateAddr.Text, textBox_AssayDateAddr.Text);
                Excel.Range _xlRange_Subject = _xlWS.get_Range(textBox_SubjectAddr.Text, textBox_SubjectAddr.Text);
                Excel.Range _xlRange_FirstDataCell = _xlWS.get_Range(textBox_FirstDataCell.Text, textBox_FirstDataCell.Text);
                Excel.Range _xlRange_LastDataCell = _xlWS.get_Range(textBox_LastDataCell.Text, textBox_LastDataCell.Text);


                Int32 _dateOffset = _xlRange_AssayDate.Column;
                Int32 _start_datarow_inx = _xlRange_FirstDataCell.Row;
                Int32 _last_datarow_inx = _xlRange_LastDataCell.Row + 1;


                for (int i = _start_datarow_inx; i < _last_datarow_inx; i++)
                {

                    string _assay_date = DateTime.FromOADate((double)(_xlWS.Cells[_xlRange_AssayDate.Row, _xlRange_AssayDate.Column] as Excel.Range).Value2).Date.ToString("yyyy-MM-dd");
                    
                    //string _subject = _xlWS.Name.Replace("p", "");

                    string _subject = (string)(_xlWS.Cells[_xlRange_Subject.Row , _xlRange_Subject.Column ] as Excel.Range).Value2;

                    _subject = _subject.Replace("Subject #: ", "").Replace("Control-", "");
                    
                    //_subject = _subject.Split(' ')[0];



                    if ((_xlWS.Cells[i, 1] as Excel.Range).Value2 != null)
                    {
                        double _timepoint = 0;
                        double _349;
                        double _355;
                        double _355_349;
                        double _ttr;
                        double _mp;
                        try
                        {
                            _timepoint   = (double)(_xlWS.Cells[i, 1] as Excel.Range).Value2;
                            _349         = (double)(_xlWS.Cells[i, 2] as Excel.Range).Value2;
                            _355         = (double)(_xlWS.Cells[i, 3] as Excel.Range).Value2;
                            _355_349     = (double)(_xlWS.Cells[i, 4] as Excel.Range).Value2;
                            _ttr         = (double)(_xlWS.Cells[i, 5] as Excel.Range).Value2;
                            _mp          = _355 / (_355 + _349);
                        }
                        catch 
                        {
                            MessageBox.Show(String.Format
                                ("Yaroslav Matviyiv says: There is a problem with subject {0}, hr {1}", _subject, _timepoint),
                                 "Missing value");
                            continue;
                        }

                        if (_timepoint == 0.0) { _mp0 = _mp; };
                        double _mpe = (_mp - _mp0) * 100;

                        

                        string _insert = "INSERT INTO TMP_GC_MS_IMPORT " +
                        "(SUBJECT_NUM, FLUID_TYPE_ID, ASSAY_DATE, TP_ID, NUM_349, NUM_355, MP, MPE, NUM355_349, TTR) ";
                        string _values = String.Format("VALUES ('{0}', {1}, '{2}', {3}, {4}, {5}, {6}, {7}, {8}, {9});",
                               _subject,
                               _fluid_id, 
                               _assay_date,

                               //_timepoint,

                               Get_TP(_timepoint),

                               _349,
                               _355,
                               _mp, _mpe,
                               _355_349,
                               _ttr);

                        _exportList.Add(_insert);
                        _exportList.Add(_values);

                    }
                }
            }

        }

        private void YarosheskiExportTAU(ref Excel.Workbook _xlWB, byte _fluid_id) // new format after 2018-08-29
        {
            _exportList = new List<string>();

            foreach (Excel.Worksheet _xlWS in _xlWB.Worksheets)
            {
                double _mp0 = 0;

                Excel.Range _xlRange_AssayDate = _xlWS.get_Range(textBox_AssayDateAddr.Text, textBox_AssayDateAddr.Text);
                Excel.Range _xlRange_Subject = _xlWS.get_Range(textBox_SubjectAddr.Text, textBox_SubjectAddr.Text);
                Excel.Range _xlRange_FirstDataCell = _xlWS.get_Range(textBox_FirstDataCell.Text, textBox_FirstDataCell.Text);
                Excel.Range _xlRange_LastDataCell = _xlWS.get_Range(textBox_LastDataCell.Text, textBox_LastDataCell.Text);


                Int32 _dateOffset = _xlRange_AssayDate.Column;
                Int32 _start_datarow_inx = _xlRange_FirstDataCell.Row;
                Int32 _last_datarow_inx = _xlRange_LastDataCell.Row + 1;

                int dataRowCount = 0;

                for (int i = _start_datarow_inx; i < _last_datarow_inx; i++)
                {
                    dataRowCount++;

                    var sAssayDate = (_xlWS.Cells[_xlRange_AssayDate.Row, _xlRange_AssayDate.Column] as Excel.Range).Value.ToString();
                    
                    var _assayDate = DateTime.Parse(sAssayDate);

                    string _assay_date = _assayDate.Date.ToString("yyyy-MM-dd");
                    
                    string _subject = (string)(_xlWS.Cells[_xlRange_Subject.Row, _xlRange_Subject.Column] as Excel.Range).Value2;

                    _subject = _subject.Replace("Subject #: ", "").Replace("Control-", "");

                    if ((_xlWS.Cells[i, 1] as Excel.Range).Value2 != null)
                    {
                        double _timepoint = 0;
                        double _349;
                        double _355;
                        double _355_349;
                        double _ttr;
                        double _mp;
                        double _calcTTR;
                        try
                        {
                            _timepoint = (double)(_xlWS.Cells[i, 1] as Excel.Range).Value2;
                            _349 = (double)(_xlWS.Cells[i, 2] as Excel.Range).Value2;
                            _355 = (double)(_xlWS.Cells[i, 3] as Excel.Range).Value2;
                            _355_349 = (double)(_xlWS.Cells[i, 4] as Excel.Range).Value2;
                            _ttr = (double)(_xlWS.Cells[i, 5] as Excel.Range).Value2;
                            _calcTTR = (double)(_xlWS.Cells[i, 6] as Excel.Range).Value2;
                            _mp = _355 / (_355 + _349);
                        }
                        catch
                        {
                            MessageBox.Show(String.Format
                                ("Vitaliy says: There is a problem with subject {0}, hr {1}", _subject, _timepoint),
                                 "Missing value");
                            continue;
                        }

                        if (_timepoint == 0.0) { _mp0 = _mp; };
                        double _mpe = (_mp - _mp0) * 100;



                        string _insert = "INSERT INTO TMP_GC_MS_IMPORT " +
                        "(SUBJECT_NUM, FLUID_TYPE_ID, ASSAY_DATE, TP_ID, NUM_349, NUM_355, MP, MPE, NUM355_349, TTR, CALC_TTR) ";
                        string _values = String.Format("VALUES ('{0}', {1}, '{2}', {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10});",
                               _subject,
                               _fluid_id,
                               _assay_date,
                               
                               Get_TP_TAU(dataRowCount.ToString()),

                               _349,
                               _355,
                               _mp, _mpe,
                               _355_349,
                               _ttr,
                               _calcTTR);

                        _exportList.Add(_insert);
                        _exportList.Add(_values);

                    }
                }
            }

        }

        private void PattersonExport(ref Excel.Workbook _xlWB, byte _fluid_id)
        {
            _exportList = new List<string>();

            foreach (Excel.Worksheet _xlWS in _xlWB.Worksheets)
            {
                double _mp0 = 0;

                Excel.Range _xlRange_AssayDate = _xlWS.get_Range(textBox_AssayDateAddr.Text, textBox_AssayDateAddr.Text);

                for (int i = 3; i < 50; i++)
                {


                    string _assay_date = null;

                    try
                    {
                        _assay_date = DateTime.FromOADate(Convert.ToDouble((_xlWS.Cells[_xlRange_AssayDate.Row, _xlRange_AssayDate.Column] as Excel.Range).Value2)).Date.ToString("MM/dd/yyyy");
                    }

                    catch
                    {
                        MessageBox.Show("Issue in converting date in row (i): " + i.ToString() + "  on worksheet: " + _xlWS.Name);
                    }

                   /* string _assay_date = DateTime.FromOADate((double)(_xlWS.Cells[5, 3] as Excel.Range).Value2).Date.ToString("yyyy-MM-dd");*/
                    string _subject = _xlWS.Name.Replace("p", "").Replace("_", "");

                   // if (i == 3) { (_xlWS.Cells[i, 1] as Excel.Range).Value2 = 0; }

                    if (     (   (_xlWS.Cells[i, 4] as Excel.Range).Value2 != null    ) &&
                             (   (_xlWS.Cells[i, 1] as Excel.Range).Value2.ToString() != "blank"  )     )
                    {

                        double _timepoint = 0; string _timepointStr = "";
                        double _200;
                        double _205;
                        double _205_200;
                        double _ttr; double _ttr_subst_R0;
                        double _mp;
                        try
                        {
                            _timepoint    = (double)(_xlWS.Cells[i,  2] as Excel.Range).Value2; _timepointStr = (_xlWS.Cells[i, 1] as Excel.Range).Value2.ToString() ;
                            _200          = (double)(_xlWS.Cells[i,  9] as Excel.Range).Value2;
                            _205          = (double)(_xlWS.Cells[i, 11] as Excel.Range).Value2;
                            _205_200      = (double)(_xlWS.Cells[i, 13] as Excel.Range).Value2;
                            _ttr          = (double)(_xlWS.Cells[i, 14] as Excel.Range).Value2;
                            _ttr_subst_R0 = (double)(_xlWS.Cells[i, 15] as Excel.Range).Value2;
                            _mp           =  _205 / (_205 + _200);
                        }
                        catch
                        {
                            MessageBox.Show(String.Format
                                ("Manohar says: There is a problem with subject {0}, raw {1}", _subject, i),
                                 "Missing value");
                            continue;
                        }

                        if (_timepoint == 0.0) { _mp0 = _mp; };
                        double _mpe = (_mp - _mp0) * 100;

                        _timepointStr = _timepointStr.Replace(_subject, "").Trim();

                        if (_timepointStr.Contains("LP") && _timepointStr.Trim().Length == 3)
                        {
                           _timepointStr = _timepointStr.Insert(2, " ");
                        }
                        string _insert = "INSERT INTO TMP_GC_MS_IMPORT1 " +
                        "(SUBJECT_NUM, FLUID_TYPE_ID, ASSAY_DATE, TP_ID, NUM_200, NUM_205, MP, MPE, NUM205_200, NUM_TTR_BY_CAL, NUM_TTR_BY_CAL_SUBS_R0) ";
                        string _values = String.Format("VALUES ('{0}', {1}, '{2}', {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10});",
                               _subject,
                               _fluid_id, 
                               _assay_date,

                               //_timepoint,

                               Get_TP1(_timepointStr),
                                                              
                               _200,
                               _205,
                               _mp, _mpe,
                               _205_200,
                               _ttr,
                               _ttr_subst_R0);

                        _exportList.Add(_insert);
                        _exportList.Add(_values);

                    }
                }
            }

        }

        private void PattersonExport1(ref Excel.Workbook _xlWB, byte _fluid_id)
        {
            _exportList = new List<string>();

            foreach (Excel.Worksheet _xlWS in _xlWB.Worksheets)
            {
                double _mp0 = 0;

                Excel.Range _xlRange_AssayDate = _xlWS.get_Range(textBox_AssayDateAddr.Text, textBox_AssayDateAddr.Text);

                for (int i = 3; i < 50; i++)
                {
                    string _assay_date = DateTime.FromOADate((double)(_xlWS.Cells[_xlRange_AssayDate.Row, _xlRange_AssayDate.Column] as Excel.Range).Value2).Date.ToString("yyyy-MM-dd");
                    /* string _assay_date = DateTime.FromOADate((double)(_xlWS.Cells[5, 3] as Excel.Range).Value2).Date.ToString("yyyy-MM-dd");*/
                    string _subject = _xlWS.Name.Replace("p", "").Replace("_", "");

                    // if (i == 3) { (_xlWS.Cells[i, 1] as Excel.Range).Value2 = 0; }

                    if (((_xlWS.Cells[i, 4] as Excel.Range).Value2 != null) &&
                             ((_xlWS.Cells[i, 1] as Excel.Range).Value2.ToString() != "blank"))
                    {

                        double _timepoint = 0; string _timepointStr = "";
                        double _200;
                        double _205;
                        double _205_200;
                        double _ttr; double _ttr_subst_R0;
                        double _mp;
                        try
                        {
                            _timepoint = (double)(_xlWS.Cells[i, 2] as Excel.Range).Value2; _timepointStr = (_xlWS.Cells[i, 1] as Excel.Range).Value2.ToString();
                            _200 = (double)(_xlWS.Cells[i, 9] as Excel.Range).Value2;
                            _205 = (double)(_xlWS.Cells[i, 11] as Excel.Range).Value2;
                            _205_200 = (double)(_xlWS.Cells[i, 13] as Excel.Range).Value2;
                            _ttr = (double)(_xlWS.Cells[i, 14] as Excel.Range).Value2;
                            _ttr_subst_R0 = (double)(_xlWS.Cells[i, 15] as Excel.Range).Value2;
                            _mp = _205 / (_205 + _200);
                        }
                        catch
                        {
                            MessageBox.Show(String.Format
                                ("Terry Hicks says: There is a problem with subject {0}, hr {1}", _subject, _timepoint),
                                 "Missing value");
                            continue;
                        }

                        if (_timepoint == 0.0) { _mp0 = _mp; };
                        double _mpe = (_mp - _mp0) * 100;

                        _timepointStr = _timepointStr.Replace(_subject, "").Trim();

                        if (_timepointStr.Contains("LP") && _timepointStr.Trim().Length == 3)
                        {
                            _timepointStr = _timepointStr.Insert(2, " ");
                        }
                        string _insert = "INSERT INTO TMP_GC_MS_IMPORT1 " +
                        "(SUBJECT_NUM, FLUID_TYPE_ID, ASSAY_DATE, TP_ID, NUM_200, NUM_205, MP, MPE, NUM205_200, NUM_TTR_BY_CAL, NUM_TTR_BY_CAL_SUBS_R0) ";
                        string _values = String.Format("VALUES ('{0}', {1}, '{2}', {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10});",
                               _subject,
                               _fluid_id,
                               _assay_date,

                               //_timepoint,

                               Get_TP1(_timepointStr),

                               _200,
                               _205,
                               _mp, _mpe,
                               _205_200,
                               _ttr,
                               _ttr_subst_R0);

                        _exportList.Add(_insert);
                        _exportList.Add(_values);

                    }
                }
            }

        }


        private void button_SelectGSMSExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog _openDlg = new OpenFileDialog();
            _openDlg.Filter = "Excel(*.XLS,(*.xlsx)|*.xls;*.xlsx";


            if (_openDlg.ShowDialog() == DialogResult.OK)
            {
                textBox_GCMSfile.Text = _openDlg.FileName;
                _sqlFileName = Path.ChangeExtension(textBox_GCMSfile.Text, "sql");
            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void textBox_GSMSfile_TextChanged(object sender, EventArgs e)
        {
            string _fileName = Path.GetFileNameWithoutExtension(textBox_GCMSfile.Text).ToLower();

         /*   if ( _fileName.Contains("plasma") )
            {
                radioButton_hPlasma.Checked = true;
            }
            if ( _fileName.Contains("csf") )
            {
                radioButton_hCSF.Checked = true;
            }*/
            if ( _fileName.Contains("bruce") ||
                 _fileName.Contains("patterson")  )
            {
                radioButton_PattersonLab.Checked = true;
            }
            if (_fileName.Contains("kevin") ||
                 _fileName.Contains("yarasheski"))
            {
                radioButton_YarasheskiLab.Checked = true;
            }

        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fmAbout _fmAbout = new fmAbout();
            _fmAbout.ShowDialog();
        }

        private void fmMain_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'dsBatemanLabDB.FLUID_TYPE' table. You can move, or remove it, as needed.
            this.fLUID_TYPETableAdapter.Fill(this.dsBatemanLabDB.FLUID_TYPE);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.TIME_POINT' table. You can move, or remove it, as needed.
            this.tIME_POINTTableAdapter.Fill(this.dsBatemanLabDB.TIME_POINT);
            // TODO: This line of code loads data into the 'dsBatemanLabDB.STUDY' table. You can move, or remove it, as needed.
            this.sTUDYTableAdapter.Fill(this.dsBatemanLabDB.STUDY);

        }

        private void fmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
              RegistryKey _key;
            _key = Registry.CurrentUser.CreateSubKey("SOFTWARE\\GCMSReader");



            try
            {
                foreach (TextBox _txtBox in this.panel1.Controls.OfType<TextBox>())
                {
                    _key.SetValue(_txtBox.Name, _txtBox.Text);
                }

                foreach (RadioButton _rdbttn in panel1.Controls["groupBox_LabSource"].Controls.OfType<RadioButton>())
                {
                    _key.SetValue(_rdbttn.Name, _rdbttn.Checked);
                }
               /* foreach (RadioButton _rdbttn in panel2.Controls["groupBox_Matrix"].Controls.OfType<RadioButton>())
                {

                    _key.SetValue(_rdbttn.Name, _rdbttn.Checked);
                }*/

                foreach (ComboBox _com in this.Controls.OfType<ComboBox>())
                {
                    _key.SetValue(_com.Name, _com.Text);
                }
            }
            catch (Exception _ex)
            {
                MessageBox.Show(_ex.Message);
            }
        }

        private void fmMain_Shown(object sender, EventArgs e)
        {
            RegistryKey _key;
            _key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\GCMSReader");

            if (_key == null) { return; };
               
            try
            {
                foreach (TextBox _txtBox in this.panel1.Controls.OfType<TextBox>())
                {
                    _txtBox.Text = _key.GetValue(_txtBox.Name).ToString();
                }

                foreach (RadioButton _rdbttn in panel1.Controls["groupBox_LabSource"].Controls.OfType<RadioButton>())
                {
                    _rdbttn.Checked = Convert.ToBoolean(_key.GetValue(_rdbttn.Name).ToString());
                }

               /* foreach (RadioButton _rdbttn in panel2.Controls["groupBox_Matrix"].Controls.OfType<RadioButton>())
                {
                    _rdbttn.Checked = Convert.ToBoolean(_key.GetValue(_rdbttn.Name).ToString());
                }*/

                foreach (ComboBox _combo in this.Controls.OfType<ComboBox>())
                {
                    _combo.Text = _key.GetValue(_combo.Name).ToString();
                }

            }
            catch (Exception _ex) { MessageBox.Show(_ex.Message); }
                        
            } 

            
    } 

    
}
