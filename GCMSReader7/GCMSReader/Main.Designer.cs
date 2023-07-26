namespace GCMSReader
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
            this.button_StartExport = new System.Windows.Forms.Button();
            this.textBox_GCMSfile = new System.Windows.Forms.TextBox();
            this.button_SelectGSMSExcel = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.textBox_SubjectAddr = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox_LastDataCell = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox_FirstDataCell = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox_AssayDateAddr = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox_LabSource = new System.Windows.Forms.GroupBox();
            this.radioButton_YarasheskiLab_TAU = new System.Windows.Forms.RadioButton();
            this.radioButton_YarasheskiLab_BACE = new System.Windows.Forms.RadioButton();
            this.radioButton_PattersonLab = new System.Windows.Forms.RadioButton();
            this.radioButton_YarasheskiLab = new System.Windows.Forms.RadioButton();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.contensToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.comboBox_Study = new System.Windows.Forms.ComboBox();
            this.sTUDYBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dsBatemanLabDB = new GCMSReader.dsBatemanLabDB();
            this.label11 = new System.Windows.Forms.Label();
            this.sTUDYTableAdapter = new GCMSReader.dsBatemanLabDBTableAdapters.STUDYTableAdapter();
            this.tIMEPOINTBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.tIME_POINTTableAdapter = new GCMSReader.dsBatemanLabDBTableAdapters.TIME_POINTTableAdapter();
            this.comboBox_FluidType = new System.Windows.Forms.ComboBox();
            this.fLUIDTYPEBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.label4 = new System.Windows.Forms.Label();
            this.fLUID_TYPETableAdapter = new GCMSReader.dsBatemanLabDBTableAdapters.FLUID_TYPETableAdapter();
            this.panel1.SuspendLayout();
            this.groupBox_LabSource.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sTUDYBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsBatemanLabDB)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tIMEPOINTBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.fLUIDTYPEBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // button_StartExport
            // 
            this.button_StartExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_StartExport.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.button_StartExport.Location = new System.Drawing.Point(6, 349);
            this.button_StartExport.Name = "button_StartExport";
            this.button_StartExport.Size = new System.Drawing.Size(134, 42);
            this.button_StartExport.TabIndex = 11;
            this.button_StartExport.Text = "Start an export";
            this.button_StartExport.UseVisualStyleBackColor = true;
            this.button_StartExport.Click += new System.EventHandler(this.button_StartExport_Click);
            // 
            // textBox_GCMSfile
            // 
            this.textBox_GCMSfile.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_GCMSfile.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.57F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_GCMSfile.Location = new System.Drawing.Point(12, 62);
            this.textBox_GCMSfile.Multiline = true;
            this.textBox_GCMSfile.Name = "textBox_GCMSfile";
            this.textBox_GCMSfile.ReadOnly = true;
            this.textBox_GCMSfile.Size = new System.Drawing.Size(582, 26);
            this.textBox_GCMSfile.TabIndex = 13;
            this.textBox_GCMSfile.TextChanged += new System.EventHandler(this.textBox_GSMSfile_TextChanged);
            // 
            // button_SelectGSMSExcel
            // 
            this.button_SelectGSMSExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_SelectGSMSExcel.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.button_SelectGSMSExcel.Location = new System.Drawing.Point(12, 33);
            this.button_SelectGSMSExcel.Name = "button_SelectGSMSExcel";
            this.button_SelectGSMSExcel.Size = new System.Drawing.Size(173, 23);
            this.button_SelectGSMSExcel.TabIndex = 12;
            this.button_SelectGSMSExcel.Text = "Select GSMS(xls) input file ...";
            this.button_SelectGSMSExcel.UseVisualStyleBackColor = true;
            this.button_SelectGSMSExcel.Click += new System.EventHandler(this.button_SelectGSMSExcel_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.textBox_SubjectAddr);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.textBox_LastDataCell);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.textBox_FirstDataCell);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.textBox_AssayDateAddr);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.groupBox_LabSource);
            this.panel1.Location = new System.Drawing.Point(12, 103);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(293, 179);
            this.panel1.TabIndex = 17;
            // 
            // textBox_SubjectAddr
            // 
            this.textBox_SubjectAddr.Location = new System.Drawing.Point(223, 124);
            this.textBox_SubjectAddr.Name = "textBox_SubjectAddr";
            this.textBox_SubjectAddr.Size = new System.Drawing.Size(51, 20);
            this.textBox_SubjectAddr.TabIndex = 24;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(163, 129);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(57, 13);
            this.label5.TabIndex = 23;
            this.label5.Text = "Subject in:";
            // 
            // textBox_LastDataCell
            // 
            this.textBox_LastDataCell.Location = new System.Drawing.Point(223, 150);
            this.textBox_LastDataCell.Name = "textBox_LastDataCell";
            this.textBox_LastDataCell.Size = new System.Drawing.Size(51, 20);
            this.textBox_LastDataCell.TabIndex = 22;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(148, 155);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(73, 13);
            this.label3.TabIndex = 21;
            this.label3.Text = "Last data cell:";
            // 
            // textBox_FirstDataCell
            // 
            this.textBox_FirstDataCell.Location = new System.Drawing.Point(91, 148);
            this.textBox_FirstDataCell.Name = "textBox_FirstDataCell";
            this.textBox_FirstDataCell.Size = new System.Drawing.Size(51, 20);
            this.textBox_FirstDataCell.TabIndex = 20;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 155);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(72, 13);
            this.label2.TabIndex = 19;
            this.label2.Text = "First data cell:";
            // 
            // textBox_AssayDateAddr
            // 
            this.textBox_AssayDateAddr.Location = new System.Drawing.Point(91, 125);
            this.textBox_AssayDateAddr.Name = "textBox_AssayDateAddr";
            this.textBox_AssayDateAddr.Size = new System.Drawing.Size(51, 20);
            this.textBox_AssayDateAddr.TabIndex = 18;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 128);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 13);
            this.label1.TabIndex = 17;
            this.label1.Text = "Assay date in:";
            // 
            // groupBox_LabSource
            // 
            this.groupBox_LabSource.Controls.Add(this.radioButton_YarasheskiLab_TAU);
            this.groupBox_LabSource.Controls.Add(this.radioButton_YarasheskiLab_BACE);
            this.groupBox_LabSource.Controls.Add(this.radioButton_PattersonLab);
            this.groupBox_LabSource.Controls.Add(this.radioButton_YarasheskiLab);
            this.groupBox_LabSource.Location = new System.Drawing.Point(5, 8);
            this.groupBox_LabSource.Name = "groupBox_LabSource";
            this.groupBox_LabSource.Size = new System.Drawing.Size(272, 110);
            this.groupBox_LabSource.TabIndex = 16;
            this.groupBox_LabSource.TabStop = false;
            this.groupBox_LabSource.Text = "Choose the source of data";
            // 
            // radioButton_YarasheskiLab_TAU
            // 
            this.radioButton_YarasheskiLab_TAU.AutoSize = true;
            this.radioButton_YarasheskiLab_TAU.Checked = true;
            this.radioButton_YarasheskiLab_TAU.Location = new System.Drawing.Point(24, 42);
            this.radioButton_YarasheskiLab_TAU.Name = "radioButton_YarasheskiLab_TAU";
            this.radioButton_YarasheskiLab_TAU.Size = new System.Drawing.Size(152, 17);
            this.radioButton_YarasheskiLab_TAU.TabIndex = 19;
            this.radioButton_YarasheskiLab_TAU.TabStop = true;
            this.radioButton_YarasheskiLab_TAU.Text = "Yarasheski Lab (Tau SILK)";
            this.radioButton_YarasheskiLab_TAU.UseVisualStyleBackColor = true;
            // 
            // radioButton_YarasheskiLab_BACE
            // 
            this.radioButton_YarasheskiLab_BACE.AutoSize = true;
            this.radioButton_YarasheskiLab_BACE.Checked = true;
            this.radioButton_YarasheskiLab_BACE.Location = new System.Drawing.Point(24, 64);
            this.radioButton_YarasheskiLab_BACE.Name = "radioButton_YarasheskiLab_BACE";
            this.radioButton_YarasheskiLab_BACE.Size = new System.Drawing.Size(135, 17);
            this.radioButton_YarasheskiLab_BACE.TabIndex = 18;
            this.radioButton_YarasheskiLab_BACE.TabStop = true;
            this.radioButton_YarasheskiLab_BACE.Text = "Yarasheski Lab (BACE)";
            this.radioButton_YarasheskiLab_BACE.UseVisualStyleBackColor = true;
            // 
            // radioButton_PattersonLab
            // 
            this.radioButton_PattersonLab.AutoSize = true;
            this.radioButton_PattersonLab.Location = new System.Drawing.Point(24, 87);
            this.radioButton_PattersonLab.Name = "radioButton_PattersonLab";
            this.radioButton_PattersonLab.Size = new System.Drawing.Size(91, 17);
            this.radioButton_PattersonLab.TabIndex = 16;
            this.radioButton_PattersonLab.Text = "Patterson Lab";
            this.radioButton_PattersonLab.UseVisualStyleBackColor = true;
            // 
            // radioButton_YarasheskiLab
            // 
            this.radioButton_YarasheskiLab.AutoSize = true;
            this.radioButton_YarasheskiLab.Checked = true;
            this.radioButton_YarasheskiLab.Location = new System.Drawing.Point(24, 19);
            this.radioButton_YarasheskiLab.Name = "radioButton_YarasheskiLab";
            this.radioButton_YarasheskiLab.Size = new System.Drawing.Size(98, 17);
            this.radioButton_YarasheskiLab.TabIndex = 15;
            this.radioButton_YarasheskiLab.TabStop = true;
            this.radioButton_YarasheskiLab.Text = "Yarasheski Lab";
            this.radioButton_YarasheskiLab.UseVisualStyleBackColor = true;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.optionsToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(604, 24);
            this.menuStrip1.TabIndex = 19;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(92, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
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
            this.contensToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // contensToolStripMenuItem
            // 
            this.contensToolStripMenuItem.Name = "contensToolStripMenuItem";
            this.contensToolStripMenuItem.Size = new System.Drawing.Size(122, 22);
            this.contensToolStripMenuItem.Text = "Contents";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(122, 22);
            this.aboutToolStripMenuItem.Text = "About ...";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // comboBox_Study
            // 
            this.comboBox_Study.DataSource = this.sTUDYBindingSource;
            this.comboBox_Study.DisplayMember = "STUDY_NAME";
            this.comboBox_Study.FormattingEnabled = true;
            this.comboBox_Study.Location = new System.Drawing.Point(104, 322);
            this.comboBox_Study.Name = "comboBox_Study";
            this.comboBox_Study.Size = new System.Drawing.Size(174, 21);
            this.comboBox_Study.TabIndex = 37;
            this.comboBox_Study.ValueMember = "STUDY_ID";
            // 
            // sTUDYBindingSource
            // 
            this.sTUDYBindingSource.DataMember = "STUDY";
            this.sTUDYBindingSource.DataSource = this.dsBatemanLabDB;
            // 
            // dsBatemanLabDB
            // 
            this.dsBatemanLabDB.DataSetName = "dsBatemanLabDB";
            this.dsBatemanLabDB.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label11.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label11.Location = new System.Drawing.Point(12, 325);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(90, 13);
            this.label11.TabIndex = 38;
            this.label11.Text = "Study (project)";
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
            // comboBox_FluidType
            // 
            this.comboBox_FluidType.DataSource = this.fLUIDTYPEBindingSource;
            this.comboBox_FluidType.DisplayMember = "FLUID_TYPE_NAME";
            this.comboBox_FluidType.FormattingEnabled = true;
            this.comboBox_FluidType.Location = new System.Drawing.Point(415, 104);
            this.comboBox_FluidType.Name = "comboBox_FluidType";
            this.comboBox_FluidType.Size = new System.Drawing.Size(174, 21);
            this.comboBox_FluidType.TabIndex = 39;
            this.comboBox_FluidType.ValueMember = "FLUID_TYPE_ID";
            // 
            // fLUIDTYPEBindingSource
            // 
            this.fLUIDTYPEBindingSource.DataMember = "FLUID_TYPE";
            this.fLUIDTYPEBindingSource.DataSource = this.dsBatemanLabDB;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label4.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label4.Location = new System.Drawing.Point(328, 109);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(70, 13);
            this.label4.TabIndex = 40;
            this.label4.Text = "Fluid Type:";
            // 
            // fLUID_TYPETableAdapter
            // 
            this.fLUID_TYPETableAdapter.ClearBeforeFill = true;
            // 
            // fmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(604, 403);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.comboBox_FluidType);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.comboBox_Study);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.textBox_GCMSfile);
            this.Controls.Add(this.button_SelectGSMSExcel);
            this.Controls.Add(this.button_StartExport);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "fmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "GCMS Reader";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.fmMain_FormClosing);
            this.Load += new System.EventHandler(this.fmMain_Load);
            this.Shown += new System.EventHandler(this.fmMain_Shown);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupBox_LabSource.ResumeLayout(false);
            this.groupBox_LabSource.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sTUDYBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsBatemanLabDB)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tIMEPOINTBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.fLUIDTYPEBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button_StartExport;
        private System.Windows.Forms.TextBox textBox_GCMSfile;
        private System.Windows.Forms.Button button_SelectGSMSExcel;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.GroupBox groupBox_LabSource;
        private System.Windows.Forms.RadioButton radioButton_PattersonLab;
        private System.Windows.Forms.RadioButton radioButton_YarasheskiLab;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem contensToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.RadioButton radioButton_YarasheskiLab_BACE;
        private System.Windows.Forms.ComboBox comboBox_Study;
        private System.Windows.Forms.Label label11;
        private dsBatemanLabDB dsBatemanLabDB;
        private System.Windows.Forms.BindingSource sTUDYBindingSource;
        private GCMSReader.dsBatemanLabDBTableAdapters.STUDYTableAdapter sTUDYTableAdapter;
        private System.Windows.Forms.TextBox textBox_AssayDateAddr;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox_LastDataCell;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox_FirstDataCell;
        private System.Windows.Forms.BindingSource tIMEPOINTBindingSource;
        private GCMSReader.dsBatemanLabDBTableAdapters.TIME_POINTTableAdapter tIME_POINTTableAdapter;
        private System.Windows.Forms.ComboBox comboBox_FluidType;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.BindingSource fLUIDTYPEBindingSource;
        private GCMSReader.dsBatemanLabDBTableAdapters.FLUID_TYPETableAdapter fLUID_TYPETableAdapter;
        private System.Windows.Forms.TextBox textBox_SubjectAddr;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.RadioButton radioButton_YarasheskiLab_TAU;
    }
}

