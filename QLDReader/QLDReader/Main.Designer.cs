namespace QLDReader
{
    partial class fmMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.menuStrip_Main = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ToolStripMenuItem_Exit = new System.Windows.Forms.ToolStripMenuItem();
            this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem_Contents = new System.Windows.Forms.ToolStripMenuItem();
            this.ToolStripMenuItem_About = new System.Windows.Forms.ToolStripMenuItem();
            this.textBox_date_format = new System.Windows.Forms.TextBox();
            this.checkBox_CustomDate = new System.Windows.Forms.CheckBox();
            this.groupBox_ExperInfo = new System.Windows.Forms.GroupBox();
            this.comboBox_Project = new System.Windows.Forms.ComboBox();
            this.pROJECTBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dsBatemanLabDB = new QLDReader.dsBatemanLabDB();
            this.label12 = new System.Windows.Forms.Label();
            this.comboBox_Study = new System.Windows.Forms.ComboBox();
            this.sTUDYBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.label11 = new System.Windows.Forms.Label();
            this.comboBox_QuantitatedBy = new System.Windows.Forms.ComboBox();
            this.lABMEMBERSBindingSource2 = new System.Windows.Forms.BindingSource(this.components);
            this.label10 = new System.Windows.Forms.Label();
            this.comboBox_DoneBy = new System.Windows.Forms.ComboBox();
            this.lABMEMBERSBindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.label9 = new System.Windows.Forms.Label();
            this.comboBox_SampleProcessBy = new System.Windows.Forms.ComboBox();
            this.lABMEMBERSBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.label8 = new System.Windows.Forms.Label();
            this.textBox_subject = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.comboBox_Abody = new System.Windows.Forms.ComboBox();
            this.aNTIBODYBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.comboBox_Enzyme = new System.Windows.Forms.ComboBox();
            this.eNZYMEBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.comboBox_Instrument = new System.Windows.Forms.ComboBox();
            this.eQUIPMENTBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.comboBox_QuantType = new System.Windows.Forms.ComboBox();
            this.qUANTTYPEBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.label6 = new System.Windows.Forms.Label();
            this.comboBox_Fluid = new System.Windows.Forms.ComboBox();
            this.fLUIDTYPEBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dateTimePicker_AssayDate = new System.Windows.Forms.DateTimePicker();
            this.checkBox_fileName = new System.Windows.Forms.CheckBox();
            this.checkBox_OpenExcel = new System.Windows.Forms.CheckBox();
            this.button_StartExport_Batch = new System.Windows.Forms.Button();
            this.button_EditExcelfileName = new System.Windows.Forms.Button();
            this.button_SelectQLD = new System.Windows.Forms.Button();
            this.textBox_Excelfile = new System.Windows.Forms.TextBox();
            this.textBox_QLDfile = new System.Windows.Forms.TextBox();
            this.button_StartExport = new System.Windows.Forms.Button();
            this.statusStrip_Main = new System.Windows.Forms.StatusStrip();
            this.toolStripProgressBarMain = new System.Windows.Forms.ToolStripProgressBar();
            this.toolStripStatusLabelMain = new System.Windows.Forms.ToolStripStatusLabel();
            this.lAB_MEMBERSTableAdapter = new QLDReader.dsBatemanLabDBTableAdapters.LAB_MEMBERSTableAdapter();
            this.fbConnection1 = new FirebirdSql.Data.FirebirdClient.FbConnection();
            this.checkBox_ExportIntoDB = new System.Windows.Forms.CheckBox();
            this.checkBox_STD = new System.Windows.Forms.CheckBox();
            this.checkBox_Tableau = new System.Windows.Forms.CheckBox();
            this.sTUDYTableAdapter = new QLDReader.dsBatemanLabDBTableAdapters.STUDYTableAdapter();
            this.tIMEPOINTBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.tIME_POINTTableAdapter = new QLDReader.dsBatemanLabDBTableAdapters.TIME_POINTTableAdapter();
            this.pROJECTTableAdapter = new QLDReader.dsBatemanLabDBTableAdapters.PROJECTTableAdapter();
            this.fLUID_TYPETableAdapter = new QLDReader.dsBatemanLabDBTableAdapters.FLUID_TYPETableAdapter();
            this.eNZYMETableAdapter = new QLDReader.dsBatemanLabDBTableAdapters.ENZYMETableAdapter();
            this.aNTIBODYTableAdapter = new QLDReader.dsBatemanLabDBTableAdapters.ANTIBODYTableAdapter();
            this.eQUIPMENTTableAdapter = new QLDReader.dsBatemanLabDBTableAdapters.EQUIPMENTTableAdapter();
            this.qUANT_TYPETableAdapter = new QLDReader.dsBatemanLabDBTableAdapters.QUANT_TYPETableAdapter();
            this.checkBox_ShowError = new System.Windows.Forms.CheckBox();
            this.checkBox_ShowDebug = new System.Windows.Forms.CheckBox();
            this.menuStrip_Main.SuspendLayout();
            this.groupBox_ExperInfo.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pROJECTBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsBatemanLabDB)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sTUDYBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lABMEMBERSBindingSource2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lABMEMBERSBindingSource1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lABMEMBERSBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.aNTIBODYBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.eNZYMEBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.eQUIPMENTBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.qUANTTYPEBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.fLUIDTYPEBindingSource)).BeginInit();
            this.statusStrip_Main.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tIMEPOINTBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip_Main
            // 
            this.menuStrip_Main.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.optionsToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip_Main.Location = new System.Drawing.Point(0, 0);
            this.menuStrip_Main.Name = "menuStrip_Main";
            this.menuStrip_Main.Size = new System.Drawing.Size(695, 24);
            this.menuStrip_Main.TabIndex = 5;
            this.menuStrip_Main.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ToolStripMenuItem_Exit});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // ToolStripMenuItem_Exit
            // 
            this.ToolStripMenuItem_Exit.Name = "ToolStripMenuItem_Exit";
            this.ToolStripMenuItem_Exit.Size = new System.Drawing.Size(92, 22);
            this.ToolStripMenuItem_Exit.Text = "Exit";
            this.ToolStripMenuItem_Exit.Click += new System.EventHandler(this.ToolStripMenuItem_Exit_Click);
            // 
            // optionsToolStripMenuItem
            // 
            this.optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            this.optionsToolStripMenuItem.Size = new System.Drawing.Size(61, 20);
            this.optionsToolStripMenuItem.Text = "Options";
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem_Contents,
            this.ToolStripMenuItem_About});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // toolStripMenuItem_Contents
            // 
            this.toolStripMenuItem_Contents.Name = "toolStripMenuItem_Contents";
            this.toolStripMenuItem_Contents.Size = new System.Drawing.Size(122, 22);
            this.toolStripMenuItem_Contents.Text = "Contents";
            // 
            // ToolStripMenuItem_About
            // 
            this.ToolStripMenuItem_About.Name = "ToolStripMenuItem_About";
            this.ToolStripMenuItem_About.Size = new System.Drawing.Size(122, 22);
            this.ToolStripMenuItem_About.Text = "About ...";
            this.ToolStripMenuItem_About.Click += new System.EventHandler(this.ToolStripMenuItem_About_Click);
            // 
            // textBox_date_format
            // 
            this.textBox_date_format.Location = new System.Drawing.Point(437, 56);
            this.textBox_date_format.Name = "textBox_date_format";
            this.textBox_date_format.Size = new System.Drawing.Size(100, 20);
            this.textBox_date_format.TabIndex = 22;
            this.textBox_date_format.Text = "yyyy-MM-dd";
            this.textBox_date_format.Visible = false;
            this.textBox_date_format.TextChanged += new System.EventHandler(this.textBox_date_format_TextChanged);
            // 
            // checkBox_CustomDate
            // 
            this.checkBox_CustomDate.AutoSize = true;
            this.checkBox_CustomDate.Location = new System.Drawing.Point(325, 59);
            this.checkBox_CustomDate.Name = "checkBox_CustomDate";
            this.checkBox_CustomDate.Size = new System.Drawing.Size(106, 17);
            this.checkBox_CustomDate.TabIndex = 21;
            this.checkBox_CustomDate.Text = "Cust. date format";
            this.checkBox_CustomDate.UseVisualStyleBackColor = true;
            this.checkBox_CustomDate.CheckedChanged += new System.EventHandler(this.checkBox_CustomDate_CheckedChanged);
            // 
            // groupBox_ExperInfo
            // 
            this.groupBox_ExperInfo.Controls.Add(this.comboBox_Project);
            this.groupBox_ExperInfo.Controls.Add(this.label12);
            this.groupBox_ExperInfo.Controls.Add(this.comboBox_Study);
            this.groupBox_ExperInfo.Controls.Add(this.label11);
            this.groupBox_ExperInfo.Controls.Add(this.comboBox_QuantitatedBy);
            this.groupBox_ExperInfo.Controls.Add(this.label10);
            this.groupBox_ExperInfo.Controls.Add(this.comboBox_DoneBy);
            this.groupBox_ExperInfo.Controls.Add(this.label9);
            this.groupBox_ExperInfo.Controls.Add(this.comboBox_SampleProcessBy);
            this.groupBox_ExperInfo.Controls.Add(this.label8);
            this.groupBox_ExperInfo.Controls.Add(this.textBox_subject);
            this.groupBox_ExperInfo.Controls.Add(this.label7);
            this.groupBox_ExperInfo.Controls.Add(this.comboBox_Abody);
            this.groupBox_ExperInfo.Controls.Add(this.comboBox_Enzyme);
            this.groupBox_ExperInfo.Controls.Add(this.comboBox_Instrument);
            this.groupBox_ExperInfo.Controls.Add(this.comboBox_QuantType);
            this.groupBox_ExperInfo.Controls.Add(this.label6);
            this.groupBox_ExperInfo.Controls.Add(this.comboBox_Fluid);
            this.groupBox_ExperInfo.Controls.Add(this.label5);
            this.groupBox_ExperInfo.Controls.Add(this.label4);
            this.groupBox_ExperInfo.Controls.Add(this.label3);
            this.groupBox_ExperInfo.Controls.Add(this.label2);
            this.groupBox_ExperInfo.Controls.Add(this.label1);
            this.groupBox_ExperInfo.Controls.Add(this.dateTimePicker_AssayDate);
            this.groupBox_ExperInfo.Location = new System.Drawing.Point(17, 138);
            this.groupBox_ExperInfo.Name = "groupBox_ExperInfo";
            this.groupBox_ExperInfo.Size = new System.Drawing.Size(664, 149);
            this.groupBox_ExperInfo.TabIndex = 13;
            this.groupBox_ExperInfo.TabStop = false;
            this.groupBox_ExperInfo.Text = "Experiment Info";
            this.groupBox_ExperInfo.Visible = false;
            // 
            // comboBox_Project
            // 
            this.comboBox_Project.DataSource = this.pROJECTBindingSource;
            this.comboBox_Project.DisplayMember = "PROJECT_NAME";
            this.comboBox_Project.FormattingEnabled = true;
            this.comboBox_Project.Location = new System.Drawing.Point(456, 13);
            this.comboBox_Project.Name = "comboBox_Project";
            this.comboBox_Project.Size = new System.Drawing.Size(200, 21);
            this.comboBox_Project.TabIndex = 40;
            this.comboBox_Project.ValueMember = "PROJECT_ID";
            // 
            // pROJECTBindingSource
            // 
            this.pROJECTBindingSource.DataMember = "PROJECT";
            this.pROJECTBindingSource.DataSource = this.dsBatemanLabDB;
            // 
            // dsBatemanLabDB
            // 
            this.dsBatemanLabDB.DataSetName = "dsBatemanLabDB";
            this.dsBatemanLabDB.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label12.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label12.Location = new System.Drawing.Point(367, 16);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(47, 13);
            this.label12.TabIndex = 39;
            this.label12.Text = "Project";
            // 
            // comboBox_Study
            // 
            this.comboBox_Study.DataSource = this.sTUDYBindingSource;
            this.comboBox_Study.DisplayMember = "STUDY_NAME";
            this.comboBox_Study.FormattingEnabled = true;
            this.comboBox_Study.Location = new System.Drawing.Point(107, 13);
            this.comboBox_Study.Name = "comboBox_Study";
            this.comboBox_Study.Size = new System.Drawing.Size(174, 21);
            this.comboBox_Study.TabIndex = 36;
            this.comboBox_Study.ValueMember = "STUDY_ID";
            // 
            // sTUDYBindingSource
            // 
            this.sTUDYBindingSource.DataMember = "STUDY";
            this.sTUDYBindingSource.DataSource = this.dsBatemanLabDB;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label11.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label11.Location = new System.Drawing.Point(5, 16);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(90, 13);
            this.label11.TabIndex = 35;
            this.label11.Text = "Study (project)";
            // 
            // comboBox_QuantitatedBy
            // 
            this.comboBox_QuantitatedBy.DataSource = this.lABMEMBERSBindingSource2;
            this.comboBox_QuantitatedBy.DisplayMember = "DISPLAY_NAME";
            this.comboBox_QuantitatedBy.FormattingEnabled = true;
            this.comboBox_QuantitatedBy.Location = new System.Drawing.Point(535, 110);
            this.comboBox_QuantitatedBy.Name = "comboBox_QuantitatedBy";
            this.comboBox_QuantitatedBy.Size = new System.Drawing.Size(121, 21);
            this.comboBox_QuantitatedBy.TabIndex = 20;
            this.comboBox_QuantitatedBy.ValueMember = "LAB_MEMBERS_ID";
            this.comboBox_QuantitatedBy.SelectedValueChanged += new System.EventHandler(this.comboBox_QuantitatedBy_SelectedValueChanged);
            // 
            // lABMEMBERSBindingSource2
            // 
            this.lABMEMBERSBindingSource2.DataMember = "LAB_MEMBERS";
            this.lABMEMBERSBindingSource2.DataSource = this.dsBatemanLabDB;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(456, 116);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(76, 13);
            this.label10.TabIndex = 19;
            this.label10.Text = "Quantitated by";
            // 
            // comboBox_DoneBy
            // 
            this.comboBox_DoneBy.DataSource = this.lABMEMBERSBindingSource1;
            this.comboBox_DoneBy.DisplayMember = "DISPLAY_NAME";
            this.comboBox_DoneBy.FormattingEnabled = true;
            this.comboBox_DoneBy.Location = new System.Drawing.Point(307, 110);
            this.comboBox_DoneBy.Name = "comboBox_DoneBy";
            this.comboBox_DoneBy.Size = new System.Drawing.Size(121, 21);
            this.comboBox_DoneBy.TabIndex = 18;
            this.comboBox_DoneBy.ValueMember = "LAB_MEMBERS_ID";
            this.comboBox_DoneBy.SelectedValueChanged += new System.EventHandler(this.comboBox_DoneBy_SelectedValueChanged);
            // 
            // lABMEMBERSBindingSource1
            // 
            this.lABMEMBERSBindingSource1.DataMember = "LAB_MEMBERS";
            this.lABMEMBERSBindingSource1.DataSource = this.dsBatemanLabDB;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(257, 113);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(47, 13);
            this.label9.TabIndex = 17;
            this.label9.Text = "Done by";
            // 
            // comboBox_SampleProcessBy
            // 
            this.comboBox_SampleProcessBy.DataSource = this.lABMEMBERSBindingSource;
            this.comboBox_SampleProcessBy.DisplayMember = "DISPLAY_NAME";
            this.comboBox_SampleProcessBy.FormattingEnabled = true;
            this.comboBox_SampleProcessBy.Location = new System.Drawing.Point(107, 110);
            this.comboBox_SampleProcessBy.Name = "comboBox_SampleProcessBy";
            this.comboBox_SampleProcessBy.Size = new System.Drawing.Size(121, 21);
            this.comboBox_SampleProcessBy.TabIndex = 16;
            this.comboBox_SampleProcessBy.ValueMember = "LAB_MEMBERS_ID";
            this.comboBox_SampleProcessBy.SelectedValueChanged += new System.EventHandler(this.comboBox_SampleProcessBy_SelectedValueChanged);
            // 
            // lABMEMBERSBindingSource
            // 
            this.lABMEMBERSBindingSource.DataMember = "LAB_MEMBERS";
            this.lABMEMBERSBindingSource.DataSource = this.dsBatemanLabDB;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(8, 113);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(96, 13);
            this.label8.TabIndex = 15;
            this.label8.Text = "Sample process by";
            // 
            // textBox_subject
            // 
            this.textBox_subject.Location = new System.Drawing.Point(72, 74);
            this.textBox_subject.Name = "textBox_subject";
            this.textBox_subject.Size = new System.Drawing.Size(74, 20);
            this.textBox_subject.TabIndex = 14;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(6, 80);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(50, 13);
            this.label7.TabIndex = 13;
            this.label7.Text = "Subject#";
            // 
            // comboBox_Abody
            // 
            this.comboBox_Abody.DataSource = this.aNTIBODYBindingSource;
            this.comboBox_Abody.DisplayMember = "ANTIBODY_NAME";
            this.comboBox_Abody.FormattingEnabled = true;
            this.comboBox_Abody.Location = new System.Drawing.Point(361, 45);
            this.comboBox_Abody.Name = "comboBox_Abody";
            this.comboBox_Abody.Size = new System.Drawing.Size(107, 21);
            this.comboBox_Abody.TabIndex = 12;
            this.comboBox_Abody.ValueMember = "ANTIBODY_ID";
            this.comboBox_Abody.SelectedValueChanged += new System.EventHandler(this.comboBox_Abody_SelectedValueChanged);
            // 
            // aNTIBODYBindingSource
            // 
            this.aNTIBODYBindingSource.DataMember = "ANTIBODY";
            this.aNTIBODYBindingSource.DataSource = this.dsBatemanLabDB;
            // 
            // comboBox_Enzyme
            // 
            this.comboBox_Enzyme.DataSource = this.eNZYMEBindingSource;
            this.comboBox_Enzyme.DisplayMember = "ENZYME_NAME";
            this.comboBox_Enzyme.FormattingEnabled = true;
            this.comboBox_Enzyme.Location = new System.Drawing.Point(361, 73);
            this.comboBox_Enzyme.Name = "comboBox_Enzyme";
            this.comboBox_Enzyme.Size = new System.Drawing.Size(107, 21);
            this.comboBox_Enzyme.TabIndex = 11;
            this.comboBox_Enzyme.ValueMember = "ENZYME_ID";
            this.comboBox_Enzyme.SelectedValueChanged += new System.EventHandler(this.comboBox_Enzyme_SelectedValueChanged);
            // 
            // eNZYMEBindingSource
            // 
            this.eNZYMEBindingSource.DataMember = "ENZYME";
            this.eNZYMEBindingSource.DataSource = this.dsBatemanLabDB;
            // 
            // comboBox_Instrument
            // 
            this.comboBox_Instrument.DataSource = this.eQUIPMENTBindingSource;
            this.comboBox_Instrument.DisplayMember = "EQUIPMENT_NAME";
            this.comboBox_Instrument.FormattingEnabled = true;
            this.comboBox_Instrument.Location = new System.Drawing.Point(536, 45);
            this.comboBox_Instrument.Name = "comboBox_Instrument";
            this.comboBox_Instrument.Size = new System.Drawing.Size(122, 21);
            this.comboBox_Instrument.TabIndex = 10;
            this.comboBox_Instrument.ValueMember = "EQUIPMENT_ID";
            this.comboBox_Instrument.SelectedValueChanged += new System.EventHandler(this.comboBox_Instrument_SelectedValueChanged);
            // 
            // eQUIPMENTBindingSource
            // 
            this.eQUIPMENTBindingSource.DataMember = "EQUIPMENT";
            this.eQUIPMENTBindingSource.DataSource = this.dsBatemanLabDB;
            // 
            // comboBox_QuantType
            // 
            this.comboBox_QuantType.DataSource = this.qUANTTYPEBindingSource;
            this.comboBox_QuantType.DisplayMember = "QUANT_TYPE_NAME";
            this.comboBox_QuantType.FormattingEnabled = true;
            this.comboBox_QuantType.Location = new System.Drawing.Point(536, 73);
            this.comboBox_QuantType.Name = "comboBox_QuantType";
            this.comboBox_QuantType.Size = new System.Drawing.Size(122, 21);
            this.comboBox_QuantType.TabIndex = 9;
            this.comboBox_QuantType.ValueMember = "QUANT_TYPE_ID";
            this.comboBox_QuantType.SelectedValueChanged += new System.EventHandler(this.comboBox_QuantType_SelectedValueChanged);
            // 
            // qUANTTYPEBindingSource
            // 
            this.qUANTTYPEBindingSource.DataMember = "QUANT_TYPE";
            this.qUANTTYPEBindingSource.DataSource = this.dsBatemanLabDB;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(471, 76);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(59, 13);
            this.label6.TabIndex = 8;
            this.label6.Text = "Quant type";
            // 
            // comboBox_Fluid
            // 
            this.comboBox_Fluid.DataSource = this.fLUIDTYPEBindingSource;
            this.comboBox_Fluid.DisplayMember = "FLUID_TYPE_NAME";
            this.comboBox_Fluid.FormattingEnabled = true;
            this.comboBox_Fluid.Location = new System.Drawing.Point(191, 72);
            this.comboBox_Fluid.Name = "comboBox_Fluid";
            this.comboBox_Fluid.Size = new System.Drawing.Size(107, 21);
            this.comboBox_Fluid.TabIndex = 7;
            this.comboBox_Fluid.ValueMember = "FLUID_TYPE_ID";
            this.comboBox_Fluid.SelectedValueChanged += new System.EventHandler(this.comboBox_Fluid_SelectedValueChanged);
            // 
            // fLUIDTYPEBindingSource
            // 
            this.fLUIDTYPEBindingSource.DataMember = "FLUID_TYPE";
            this.fLUIDTYPEBindingSource.DataSource = this.dsBatemanLabDB;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(474, 48);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(56, 13);
            this.label5.TabIndex = 6;
            this.label5.Text = "Instrument";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(313, 77);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(44, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "Enzyme";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(313, 49);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(40, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "A-body";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(156, 76);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Fluid";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(30, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Date";
            // 
            // dateTimePicker_AssayDate
            // 
            this.dateTimePicker_AssayDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker_AssayDate.Location = new System.Drawing.Point(72, 44);
            this.dateTimePicker_AssayDate.Name = "dateTimePicker_AssayDate";
            this.dateTimePicker_AssayDate.Size = new System.Drawing.Size(90, 20);
            this.dateTimePicker_AssayDate.TabIndex = 1;
            this.dateTimePicker_AssayDate.ValueChanged += new System.EventHandler(this.dateTimePicker_AssayDate_ValueChanged);
            // 
            // checkBox_fileName
            // 
            this.checkBox_fileName.AutoSize = true;
            this.checkBox_fileName.Checked = true;
            this.checkBox_fileName.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_fileName.Location = new System.Drawing.Point(195, 57);
            this.checkBox_fileName.Name = "checkBox_fileName";
            this.checkBox_fileName.Size = new System.Drawing.Size(107, 17);
            this.checkBox_fileName.TabIndex = 20;
            this.checkBox_fileName.Text = "Parse file name ?";
            this.checkBox_fileName.UseVisualStyleBackColor = true;
            // 
            // checkBox_OpenExcel
            // 
            this.checkBox_OpenExcel.AutoSize = true;
            this.checkBox_OpenExcel.Checked = true;
            this.checkBox_OpenExcel.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_OpenExcel.Location = new System.Drawing.Point(156, 391);
            this.checkBox_OpenExcel.Name = "checkBox_OpenExcel";
            this.checkBox_OpenExcel.Size = new System.Drawing.Size(90, 17);
            this.checkBox_OpenExcel.TabIndex = 19;
            this.checkBox_OpenExcel.Text = "Open Excel ?";
            this.checkBox_OpenExcel.UseVisualStyleBackColor = true;
            // 
            // button_StartExport_Batch
            // 
            this.button_StartExport_Batch.Enabled = false;
            this.button_StartExport_Batch.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_StartExport_Batch.Location = new System.Drawing.Point(552, 391);
            this.button_StartExport_Batch.Name = "button_StartExport_Batch";
            this.button_StartExport_Batch.Size = new System.Drawing.Size(128, 42);
            this.button_StartExport_Batch.TabIndex = 18;
            this.button_StartExport_Batch.Text = "Export All";
            this.button_StartExport_Batch.UseVisualStyleBackColor = true;
            this.button_StartExport_Batch.Click += new System.EventHandler(this.button_StartExport_Batch_Click);
            // 
            // button_EditExcelfileName
            // 
            this.button_EditExcelfileName.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_EditExcelfileName.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.button_EditExcelfileName.Location = new System.Drawing.Point(16, 293);
            this.button_EditExcelfileName.Name = "button_EditExcelfileName";
            this.button_EditExcelfileName.Size = new System.Drawing.Size(173, 23);
            this.button_EditExcelfileName.TabIndex = 17;
            this.button_EditExcelfileName.Text = "Edit Excel output file name";
            this.button_EditExcelfileName.UseVisualStyleBackColor = true;
            this.button_EditExcelfileName.Click += new System.EventHandler(this.button_EditExcelfileName_Click);
            // 
            // button_SelectQLD
            // 
            this.button_SelectQLD.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_SelectQLD.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.button_SelectQLD.Location = new System.Drawing.Point(16, 53);
            this.button_SelectQLD.Name = "button_SelectQLD";
            this.button_SelectQLD.Size = new System.Drawing.Size(173, 23);
            this.button_SelectQLD.TabIndex = 16;
            this.button_SelectQLD.Text = "Select QLD input file ...";
            this.button_SelectQLD.UseVisualStyleBackColor = true;
            this.button_SelectQLD.Click += new System.EventHandler(this.button_SelectQLD_Click);
            // 
            // textBox_Excelfile
            // 
            this.textBox_Excelfile.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.57F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_Excelfile.Location = new System.Drawing.Point(16, 324);
            this.textBox_Excelfile.Multiline = true;
            this.textBox_Excelfile.Name = "textBox_Excelfile";
            this.textBox_Excelfile.ReadOnly = true;
            this.textBox_Excelfile.Size = new System.Drawing.Size(665, 51);
            this.textBox_Excelfile.TabIndex = 15;
            // 
            // textBox_QLDfile
            // 
            this.textBox_QLDfile.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.57F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_QLDfile.Location = new System.Drawing.Point(16, 82);
            this.textBox_QLDfile.Multiline = true;
            this.textBox_QLDfile.Name = "textBox_QLDfile";
            this.textBox_QLDfile.ReadOnly = true;
            this.textBox_QLDfile.Size = new System.Drawing.Size(665, 50);
            this.textBox_QLDfile.TabIndex = 14;
            // 
            // button_StartExport
            // 
            this.button_StartExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_StartExport.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.button_StartExport.Location = new System.Drawing.Point(16, 391);
            this.button_StartExport.Name = "button_StartExport";
            this.button_StartExport.Size = new System.Drawing.Size(134, 42);
            this.button_StartExport.TabIndex = 12;
            this.button_StartExport.Text = "Start an export";
            this.button_StartExport.UseVisualStyleBackColor = true;
            this.button_StartExport.Click += new System.EventHandler(this.button_StartExport_Click);
            // 
            // statusStrip_Main
            // 
            this.statusStrip_Main.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripProgressBarMain,
            this.toolStripStatusLabelMain});
            this.statusStrip_Main.Location = new System.Drawing.Point(0, 464);
            this.statusStrip_Main.Name = "statusStrip_Main";
            this.statusStrip_Main.Size = new System.Drawing.Size(695, 22);
            this.statusStrip_Main.TabIndex = 23;
            this.statusStrip_Main.Text = "statusStrip1";
            // 
            // toolStripProgressBarMain
            // 
            this.toolStripProgressBarMain.Name = "toolStripProgressBarMain";
            this.toolStripProgressBarMain.Size = new System.Drawing.Size(185, 16);
            this.toolStripProgressBarMain.Step = 1;
            this.toolStripProgressBarMain.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.toolStripProgressBarMain.Visible = false;
            // 
            // toolStripStatusLabelMain
            // 
            this.toolStripStatusLabelMain.MergeIndex = 1;
            this.toolStripStatusLabelMain.Name = "toolStripStatusLabelMain";
            this.toolStripStatusLabelMain.Size = new System.Drawing.Size(0, 17);
            this.toolStripStatusLabelMain.Visible = false;
            // 
            // lAB_MEMBERSTableAdapter
            // 
            this.lAB_MEMBERSTableAdapter.ClearBeforeFill = true;
            // 
            // checkBox_ExportIntoDB
            // 
            this.checkBox_ExportIntoDB.AutoSize = true;
            this.checkBox_ExportIntoDB.Checked = true;
            this.checkBox_ExportIntoDB.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_ExportIntoDB.Location = new System.Drawing.Point(156, 416);
            this.checkBox_ExportIntoDB.Name = "checkBox_ExportIntoDB";
            this.checkBox_ExportIntoDB.Size = new System.Drawing.Size(103, 17);
            this.checkBox_ExportIntoDB.TabIndex = 24;
            this.checkBox_ExportIntoDB.Text = "Export into DB ?";
            this.checkBox_ExportIntoDB.UseVisualStyleBackColor = true;
            // 
            // checkBox_STD
            // 
            this.checkBox_STD.AutoSize = true;
            this.checkBox_STD.Checked = true;
            this.checkBox_STD.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_STD.Location = new System.Drawing.Point(280, 391);
            this.checkBox_STD.Name = "checkBox_STD";
            this.checkBox_STD.Size = new System.Drawing.Size(105, 17);
            this.checkBox_STD.TabIndex = 25;
            this.checkBox_STD.Text = "Standard curve?";
            this.checkBox_STD.UseVisualStyleBackColor = true;
            // 
            // checkBox_Tableau
            // 
            this.checkBox_Tableau.AutoSize = true;
            this.checkBox_Tableau.Checked = true;
            this.checkBox_Tableau.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_Tableau.Location = new System.Drawing.Point(280, 416);
            this.checkBox_Tableau.Name = "checkBox_Tableau";
            this.checkBox_Tableau.Size = new System.Drawing.Size(74, 17);
            this.checkBox_Tableau.TabIndex = 33;
            this.checkBox_Tableau.Text = "Tableau ?";
            this.checkBox_Tableau.UseVisualStyleBackColor = true;
            // 
            // sTUDYTableAdapter
            // 
            this.sTUDYTableAdapter.ClearBeforeFill = true;
            // 
            // tIMEPOINTBindingSource
            // 
            this.tIMEPOINTBindingSource.DataMember = "TIME_POINT";
            this.tIMEPOINTBindingSource.DataSource = this.dsBatemanLabDB;
            // 
            // tIME_POINTTableAdapter
            // 
            this.tIME_POINTTableAdapter.ClearBeforeFill = true;
            // 
            // pROJECTTableAdapter
            // 
            this.pROJECTTableAdapter.ClearBeforeFill = true;
            // 
            // fLUID_TYPETableAdapter
            // 
            this.fLUID_TYPETableAdapter.ClearBeforeFill = true;
            // 
            // eNZYMETableAdapter
            // 
            this.eNZYMETableAdapter.ClearBeforeFill = true;
            // 
            // aNTIBODYTableAdapter
            // 
            this.aNTIBODYTableAdapter.ClearBeforeFill = true;
            // 
            // eQUIPMENTTableAdapter
            // 
            this.eQUIPMENTTableAdapter.ClearBeforeFill = true;
            // 
            // qUANT_TYPETableAdapter
            // 
            this.qUANT_TYPETableAdapter.ClearBeforeFill = true;
            // 
            // checkBox_ShowError
            // 
            this.checkBox_ShowError.AutoSize = true;
            this.checkBox_ShowError.Checked = true;
            this.checkBox_ShowError.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_ShowError.Location = new System.Drawing.Point(407, 391);
            this.checkBox_ShowError.Name = "checkBox_ShowError";
            this.checkBox_ShowError.Size = new System.Drawing.Size(88, 17);
            this.checkBox_ShowError.TabIndex = 34;
            this.checkBox_ShowError.Text = "Show errors?";
            this.checkBox_ShowError.UseVisualStyleBackColor = true;
            // 
            // checkBox_ShowDebug
            // 
            this.checkBox_ShowDebug.AutoSize = true;
            this.checkBox_ShowDebug.Checked = true;
            this.checkBox_ShowDebug.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_ShowDebug.Location = new System.Drawing.Point(407, 416);
            this.checkBox_ShowDebug.Name = "checkBox_ShowDebug";
            this.checkBox_ShowDebug.Size = new System.Drawing.Size(114, 17);
            this.checkBox_ShowDebug.TabIndex = 35;
            this.checkBox_ShowDebug.Text = "Show Debug info?";
            this.checkBox_ShowDebug.UseVisualStyleBackColor = true;
            // 
            // fmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(695, 486);
            this.Controls.Add(this.checkBox_ShowDebug);
            this.Controls.Add(this.checkBox_ShowError);
            this.Controls.Add(this.checkBox_Tableau);
            this.Controls.Add(this.checkBox_STD);
            this.Controls.Add(this.checkBox_ExportIntoDB);
            this.Controls.Add(this.statusStrip_Main);
            this.Controls.Add(this.textBox_date_format);
            this.Controls.Add(this.checkBox_CustomDate);
            this.Controls.Add(this.groupBox_ExperInfo);
            this.Controls.Add(this.checkBox_fileName);
            this.Controls.Add(this.checkBox_OpenExcel);
            this.Controls.Add(this.button_StartExport_Batch);
            this.Controls.Add(this.button_EditExcelfileName);
            this.Controls.Add(this.button_SelectQLD);
            this.Controls.Add(this.textBox_Excelfile);
            this.Controls.Add(this.textBox_QLDfile);
            this.Controls.Add(this.button_StartExport);
            this.Controls.Add(this.menuStrip_Main);
            this.Name = "fmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "QLD Reader";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.fmMain_FormClosing);
            this.Load += new System.EventHandler(this.fmMain_Load);
            this.Shown += new System.EventHandler(this.fmMain_Shown);
            this.menuStrip_Main.ResumeLayout(false);
            this.menuStrip_Main.PerformLayout();
            this.groupBox_ExperInfo.ResumeLayout(false);
            this.groupBox_ExperInfo.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pROJECTBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsBatemanLabDB)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sTUDYBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lABMEMBERSBindingSource2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lABMEMBERSBindingSource1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lABMEMBERSBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.aNTIBODYBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.eNZYMEBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.eQUIPMENTBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.qUANTTYPEBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.fLUIDTYPEBindingSource)).EndInit();
            this.statusStrip_Main.ResumeLayout(false);
            this.statusStrip_Main.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tIMEPOINTBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip_Main;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItem_Exit;
        private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem_Contents;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItem_About;
        private System.Windows.Forms.TextBox textBox_date_format;
        private System.Windows.Forms.CheckBox checkBox_CustomDate;
        private System.Windows.Forms.GroupBox groupBox_ExperInfo;
        private System.Windows.Forms.TextBox textBox_subject;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox comboBox_Abody;
        private System.Windows.Forms.ComboBox comboBox_Enzyme;
        private System.Windows.Forms.ComboBox comboBox_Instrument;
        private System.Windows.Forms.ComboBox comboBox_QuantType;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox comboBox_Fluid;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dateTimePicker_AssayDate;
        private System.Windows.Forms.CheckBox checkBox_fileName;
        private System.Windows.Forms.CheckBox checkBox_OpenExcel;
        private System.Windows.Forms.Button button_StartExport_Batch;
        private System.Windows.Forms.Button button_EditExcelfileName;
        private System.Windows.Forms.Button button_SelectQLD;
        private System.Windows.Forms.TextBox textBox_Excelfile;
        private System.Windows.Forms.TextBox textBox_QLDfile;
        private System.Windows.Forms.Button button_StartExport;
        private System.Windows.Forms.StatusStrip statusStrip_Main;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBarMain;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelMain;
        private System.Windows.Forms.ComboBox comboBox_QuantitatedBy;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox comboBox_DoneBy;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox comboBox_SampleProcessBy;
        private System.Windows.Forms.Label label8;
        private dsBatemanLabDB dsBatemanLabDB;
        private System.Windows.Forms.BindingSource lABMEMBERSBindingSource;
        private QLDReader.dsBatemanLabDBTableAdapters.LAB_MEMBERSTableAdapter lAB_MEMBERSTableAdapter;
        private System.Windows.Forms.BindingSource lABMEMBERSBindingSource2;
        private System.Windows.Forms.BindingSource lABMEMBERSBindingSource1;
        private FirebirdSql.Data.FirebirdClient.FbConnection fbConnection1;
        private System.Windows.Forms.CheckBox checkBox_ExportIntoDB;
        private System.Windows.Forms.CheckBox checkBox_STD;
        private System.Windows.Forms.CheckBox checkBox_Tableau;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.ComboBox comboBox_Study;
        private System.Windows.Forms.BindingSource sTUDYBindingSource;
        private QLDReader.dsBatemanLabDBTableAdapters.STUDYTableAdapter sTUDYTableAdapter;
        private System.Windows.Forms.BindingSource tIMEPOINTBindingSource;
        private QLDReader.dsBatemanLabDBTableAdapters.TIME_POINTTableAdapter tIME_POINTTableAdapter;
        private System.Windows.Forms.ComboBox comboBox_Project;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.BindingSource pROJECTBindingSource;
        private QLDReader.dsBatemanLabDBTableAdapters.PROJECTTableAdapter pROJECTTableAdapter;
        private System.Windows.Forms.BindingSource fLUIDTYPEBindingSource;
        private QLDReader.dsBatemanLabDBTableAdapters.FLUID_TYPETableAdapter fLUID_TYPETableAdapter;
        private System.Windows.Forms.BindingSource eNZYMEBindingSource;
        private QLDReader.dsBatemanLabDBTableAdapters.ENZYMETableAdapter eNZYMETableAdapter;
        private System.Windows.Forms.BindingSource aNTIBODYBindingSource;
        private QLDReader.dsBatemanLabDBTableAdapters.ANTIBODYTableAdapter aNTIBODYTableAdapter;
        private System.Windows.Forms.BindingSource eQUIPMENTBindingSource;
        private QLDReader.dsBatemanLabDBTableAdapters.EQUIPMENTTableAdapter eQUIPMENTTableAdapter;
        private System.Windows.Forms.BindingSource qUANTTYPEBindingSource;
        private QLDReader.dsBatemanLabDBTableAdapters.QUANT_TYPETableAdapter qUANT_TYPETableAdapter;
        private System.Windows.Forms.CheckBox checkBox_ShowError;
        private System.Windows.Forms.CheckBox checkBox_ShowDebug;

    }
}

