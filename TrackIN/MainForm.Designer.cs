namespace WashU.BatemanLab.MassSpec.TrackIN
{
    partial class MainForm
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
            this.btnTEST = new System.Windows.Forms.Button();
            this.btnTEST2 = new System.Windows.Forms.Button();
            this.zedGraphControlTest = new ZedGraph.ZedGraphControl();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.tabMainForm = new System.Windows.Forms.TabControl();
            this.tabHomePage = new System.Windows.Forms.TabPage();
            this.btnMakeLinkerToSkyline = new System.Windows.Forms.Button();
            this.tabNoiseAnalysis = new System.Windows.Forms.TabPage();
            this.graphNoiseAnalysis = new ZedGraph.ZedGraphControl();
            this.stsNoiseAnalysis = new System.Windows.Forms.StatusStrip();
            this.stsNoiseAnalysisLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.mnuNoiseAnalysis = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuAnalizeFromSkyline = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuNASelectMSRunsToAnalyze = new System.Windows.Forms.ToolStripMenuItem();
            this.tabPeptideRatios = new System.Windows.Forms.TabPage();
            this.stsPeptideRatios = new System.Windows.Forms.StatusStrip();
            this.stsPeptideRatiosLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.graphPeptideRatios = new ZedGraph.ZedGraphControl();
            this.mnuPeptideRatios = new System.Windows.Forms.MenuStrip();
            this.mnuRatioSelection = new System.Windows.Forms.ToolStripMenuItem();
            this.tabTEST = new System.Windows.Forms.TabPage();
            this.button6 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.listView1 = new System.Windows.Forms.ListView();
            this.button1 = new System.Windows.Forms.Button();
            this.tabSequenceTool = new System.Windows.Forms.TabPage();
            this.dgSampleList = new System.Windows.Forms.DataGridView();
            this.btnPasteFromExcel = new System.Windows.Forms.Button();
            this.btnGenSequence = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.tabMainForm.SuspendLayout();
            this.tabHomePage.SuspendLayout();
            this.tabNoiseAnalysis.SuspendLayout();
            this.stsNoiseAnalysis.SuspendLayout();
            this.mnuNoiseAnalysis.SuspendLayout();
            this.tabPeptideRatios.SuspendLayout();
            this.stsPeptideRatios.SuspendLayout();
            this.mnuPeptideRatios.SuspendLayout();
            this.tabTEST.SuspendLayout();
            this.tabSequenceTool.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgSampleList)).BeginInit();
            this.SuspendLayout();
            // 
            // btnTEST
            // 
            this.btnTEST.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnTEST.Location = new System.Drawing.Point(8, 12);
            this.btnTEST.Name = "btnTEST";
            this.btnTEST.Size = new System.Drawing.Size(94, 37);
            this.btnTEST.TabIndex = 0;
            this.btnTEST.Text = "Test ...";
            this.btnTEST.UseVisualStyleBackColor = true;
            this.btnTEST.Click += new System.EventHandler(this.btnTEST_Click);
            // 
            // btnTEST2
            // 
            this.btnTEST2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnTEST2.Location = new System.Drawing.Point(144, 12);
            this.btnTEST2.Name = "btnTEST2";
            this.btnTEST2.Size = new System.Drawing.Size(94, 37);
            this.btnTEST2.TabIndex = 3;
            this.btnTEST2.Text = "Debug ...";
            this.btnTEST2.UseVisualStyleBackColor = true;
            this.btnTEST2.Click += new System.EventHandler(this.btnTEST2_Click);
            // 
            // zedGraphControlTest
            // 
            this.zedGraphControlTest.Location = new System.Drawing.Point(43, 64);
            this.zedGraphControlTest.Name = "zedGraphControlTest";
            this.zedGraphControlTest.ScrollGrace = 0D;
            this.zedGraphControlTest.ScrollMaxX = 0D;
            this.zedGraphControlTest.ScrollMaxY = 0D;
            this.zedGraphControlTest.ScrollMaxY2 = 0D;
            this.zedGraphControlTest.ScrollMinX = 0D;
            this.zedGraphControlTest.ScrollMinY = 0D;
            this.zedGraphControlTest.ScrollMinY2 = 0D;
            this.zedGraphControlTest.Size = new System.Drawing.Size(473, 706);
            this.zedGraphControlTest.TabIndex = 7;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(251, 12);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(94, 37);
            this.button3.TabIndex = 8;
            this.button3.Text = "button3";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(360, 12);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(94, 37);
            this.button4.TabIndex = 9;
            this.button4.Text = "button4";
            this.button4.UseVisualStyleBackColor = true;
            // 
            // tabMainForm
            // 
            this.tabMainForm.Controls.Add(this.tabHomePage);
            this.tabMainForm.Controls.Add(this.tabNoiseAnalysis);
            this.tabMainForm.Controls.Add(this.tabPeptideRatios);
            this.tabMainForm.Controls.Add(this.tabSequenceTool);
            this.tabMainForm.Controls.Add(this.tabTEST);
            this.tabMainForm.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabMainForm.Location = new System.Drawing.Point(0, 0);
            this.tabMainForm.Name = "tabMainForm";
            this.tabMainForm.SelectedIndex = 0;
            this.tabMainForm.Size = new System.Drawing.Size(1559, 811);
            this.tabMainForm.TabIndex = 10;
            this.tabMainForm.Enter += new System.EventHandler(this.tabMainForm_Enter);
            // 
            // tabHomePage
            // 
            this.tabHomePage.Controls.Add(this.btnMakeLinkerToSkyline);
            this.tabHomePage.Location = new System.Drawing.Point(4, 22);
            this.tabHomePage.Name = "tabHomePage";
            this.tabHomePage.Padding = new System.Windows.Forms.Padding(3);
            this.tabHomePage.Size = new System.Drawing.Size(1551, 785);
            this.tabHomePage.TabIndex = 0;
            this.tabHomePage.Text = "Home Page";
            this.tabHomePage.UseVisualStyleBackColor = true;
            // 
            // btnMakeLinkerToSkyline
            // 
            this.btnMakeLinkerToSkyline.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnMakeLinkerToSkyline.Location = new System.Drawing.Point(19, 18);
            this.btnMakeLinkerToSkyline.Name = "btnMakeLinkerToSkyline";
            this.btnMakeLinkerToSkyline.Size = new System.Drawing.Size(151, 58);
            this.btnMakeLinkerToSkyline.TabIndex = 19;
            this.btnMakeLinkerToSkyline.Text = "Make a link for Skyline";
            this.btnMakeLinkerToSkyline.UseVisualStyleBackColor = true;
            this.btnMakeLinkerToSkyline.Click += new System.EventHandler(this.btnMakeLinkerToSkyline_Click);
            // 
            // tabNoiseAnalysis
            // 
            this.tabNoiseAnalysis.Controls.Add(this.graphNoiseAnalysis);
            this.tabNoiseAnalysis.Controls.Add(this.stsNoiseAnalysis);
            this.tabNoiseAnalysis.Controls.Add(this.mnuNoiseAnalysis);
            this.tabNoiseAnalysis.Location = new System.Drawing.Point(4, 22);
            this.tabNoiseAnalysis.Name = "tabNoiseAnalysis";
            this.tabNoiseAnalysis.Size = new System.Drawing.Size(1551, 785);
            this.tabNoiseAnalysis.TabIndex = 3;
            this.tabNoiseAnalysis.Text = "Noise Analysis";
            this.tabNoiseAnalysis.UseVisualStyleBackColor = true;
            // 
            // graphNoiseAnalysis
            // 
            this.graphNoiseAnalysis.Dock = System.Windows.Forms.DockStyle.Fill;
            this.graphNoiseAnalysis.Location = new System.Drawing.Point(0, 24);
            this.graphNoiseAnalysis.Name = "graphNoiseAnalysis";
            this.graphNoiseAnalysis.ScrollGrace = 0D;
            this.graphNoiseAnalysis.ScrollMaxX = 0D;
            this.graphNoiseAnalysis.ScrollMaxY = 0D;
            this.graphNoiseAnalysis.ScrollMaxY2 = 0D;
            this.graphNoiseAnalysis.ScrollMinX = 0D;
            this.graphNoiseAnalysis.ScrollMinY = 0D;
            this.graphNoiseAnalysis.ScrollMinY2 = 0D;
            this.graphNoiseAnalysis.Size = new System.Drawing.Size(1551, 739);
            this.graphNoiseAnalysis.TabIndex = 2;
            // 
            // stsNoiseAnalysis
            // 
            this.stsNoiseAnalysis.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.stsNoiseAnalysisLabel});
            this.stsNoiseAnalysis.Location = new System.Drawing.Point(0, 763);
            this.stsNoiseAnalysis.Name = "stsNoiseAnalysis";
            this.stsNoiseAnalysis.Size = new System.Drawing.Size(1551, 22);
            this.stsNoiseAnalysis.TabIndex = 1;
            this.stsNoiseAnalysis.Text = "Ok";
            // 
            // stsNoiseAnalysisLabel
            // 
            this.stsNoiseAnalysisLabel.Name = "stsNoiseAnalysisLabel";
            this.stsNoiseAnalysisLabel.Size = new System.Drawing.Size(22, 17);
            this.stsNoiseAnalysisLabel.Text = "Ok";
            // 
            // mnuNoiseAnalysis
            // 
            this.mnuNoiseAnalysis.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.mnuNoiseAnalysis.Location = new System.Drawing.Point(0, 0);
            this.mnuNoiseAnalysis.Name = "mnuNoiseAnalysis";
            this.mnuNoiseAnalysis.Size = new System.Drawing.Size(1551, 24);
            this.mnuNoiseAnalysis.TabIndex = 0;
            this.mnuNoiseAnalysis.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuAnalizeFromSkyline,
            this.mnuNASelectMSRunsToAnalyze});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // mnuAnalizeFromSkyline
            // 
            this.mnuAnalizeFromSkyline.Name = "mnuAnalizeFromSkyline";
            this.mnuAnalizeFromSkyline.Size = new System.Drawing.Size(242, 22);
            this.mnuAnalizeFromSkyline.Text = "Analyze from Skyline document";
            this.mnuAnalizeFromSkyline.Click += new System.EventHandler(this.mnuAnalizeFromSkyline_Click);
            // 
            // mnuNASelectMSRunsToAnalyze
            // 
            this.mnuNASelectMSRunsToAnalyze.Name = "mnuNASelectMSRunsToAnalyze";
            this.mnuNASelectMSRunsToAnalyze.Size = new System.Drawing.Size(242, 22);
            this.mnuNASelectMSRunsToAnalyze.Text = "Select runs for analysis...";
            this.mnuNASelectMSRunsToAnalyze.Click += new System.EventHandler(this.mnuNASelectMSRunsToAnalyze_Click);
            // 
            // tabPeptideRatios
            // 
            this.tabPeptideRatios.Controls.Add(this.stsPeptideRatios);
            this.tabPeptideRatios.Controls.Add(this.graphPeptideRatios);
            this.tabPeptideRatios.Controls.Add(this.mnuPeptideRatios);
            this.tabPeptideRatios.Location = new System.Drawing.Point(4, 22);
            this.tabPeptideRatios.Name = "tabPeptideRatios";
            this.tabPeptideRatios.Padding = new System.Windows.Forms.Padding(3);
            this.tabPeptideRatios.Size = new System.Drawing.Size(1551, 785);
            this.tabPeptideRatios.TabIndex = 2;
            this.tabPeptideRatios.Text = "Peptide Ratios";
            this.tabPeptideRatios.UseVisualStyleBackColor = true;
            this.tabPeptideRatios.Enter += new System.EventHandler(this.tabPeptideRatios_Enter);
            // 
            // stsPeptideRatios
            // 
            this.stsPeptideRatios.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.stsPeptideRatiosLabel});
            this.stsPeptideRatios.Location = new System.Drawing.Point(3, 760);
            this.stsPeptideRatios.Name = "stsPeptideRatios";
            this.stsPeptideRatios.Size = new System.Drawing.Size(1545, 22);
            this.stsPeptideRatios.TabIndex = 2;
            this.stsPeptideRatios.Text = "statusStrip1";
            // 
            // stsPeptideRatiosLabel
            // 
            this.stsPeptideRatiosLabel.Name = "stsPeptideRatiosLabel";
            this.stsPeptideRatiosLabel.Size = new System.Drawing.Size(22, 17);
            this.stsPeptideRatiosLabel.Text = "Ok";
            // 
            // graphPeptideRatios
            // 
            this.graphPeptideRatios.Dock = System.Windows.Forms.DockStyle.Fill;
            this.graphPeptideRatios.Location = new System.Drawing.Point(3, 27);
            this.graphPeptideRatios.Name = "graphPeptideRatios";
            this.graphPeptideRatios.ScrollGrace = 0D;
            this.graphPeptideRatios.ScrollMaxX = 0D;
            this.graphPeptideRatios.ScrollMaxY = 0D;
            this.graphPeptideRatios.ScrollMaxY2 = 0D;
            this.graphPeptideRatios.ScrollMinX = 0D;
            this.graphPeptideRatios.ScrollMinY = 0D;
            this.graphPeptideRatios.ScrollMinY2 = 0D;
            this.graphPeptideRatios.Size = new System.Drawing.Size(1545, 755);
            this.graphPeptideRatios.TabIndex = 1;
            // 
            // mnuPeptideRatios
            // 
            this.mnuPeptideRatios.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuRatioSelection});
            this.mnuPeptideRatios.Location = new System.Drawing.Point(3, 3);
            this.mnuPeptideRatios.Name = "mnuPeptideRatios";
            this.mnuPeptideRatios.Size = new System.Drawing.Size(1545, 24);
            this.mnuPeptideRatios.TabIndex = 0;
            this.mnuPeptideRatios.Text = "menuStrip1";
            // 
            // mnuRatioSelection
            // 
            this.mnuRatioSelection.Name = "mnuRatioSelection";
            this.mnuRatioSelection.Size = new System.Drawing.Size(90, 20);
            this.mnuRatioSelection.Text = "Shown Ratios";
            // 
            // tabTEST
            // 
            this.tabTEST.Controls.Add(this.button6);
            this.tabTEST.Controls.Add(this.button2);
            this.tabTEST.Controls.Add(this.listBox1);
            this.tabTEST.Controls.Add(this.listView1);
            this.tabTEST.Controls.Add(this.button1);
            this.tabTEST.Controls.Add(this.button3);
            this.tabTEST.Controls.Add(this.button4);
            this.tabTEST.Controls.Add(this.btnTEST2);
            this.tabTEST.Controls.Add(this.zedGraphControlTest);
            this.tabTEST.Controls.Add(this.btnTEST);
            this.tabTEST.Location = new System.Drawing.Point(4, 22);
            this.tabTEST.Name = "tabTEST";
            this.tabTEST.Padding = new System.Windows.Forms.Padding(3);
            this.tabTEST.Size = new System.Drawing.Size(1551, 785);
            this.tabTEST.TabIndex = 1;
            this.tabTEST.Text = "TEST only";
            this.tabTEST.UseVisualStyleBackColor = true;
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(1007, 21);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(193, 23);
            this.button6.TabIndex = 16;
            this.button6.Text = "Make a link for Skyline";
            this.button6.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(602, 27);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 14;
            this.button2.Text = "button2";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(522, 64);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(1021, 706);
            this.listBox1.TabIndex = 13;
            // 
            // listView1
            // 
            this.listView1.Location = new System.Drawing.Point(1245, 12);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(220, 38);
            this.listView1.TabIndex = 12;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.List;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(475, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(93, 37);
            this.button1.TabIndex = 10;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // tabSequenceTool
            // 
            this.tabSequenceTool.Controls.Add(this.button5);
            this.tabSequenceTool.Controls.Add(this.btnGenSequence);
            this.tabSequenceTool.Controls.Add(this.btnPasteFromExcel);
            this.tabSequenceTool.Controls.Add(this.dgSampleList);
            this.tabSequenceTool.Location = new System.Drawing.Point(4, 22);
            this.tabSequenceTool.Name = "tabSequenceTool";
            this.tabSequenceTool.Size = new System.Drawing.Size(1551, 785);
            this.tabSequenceTool.TabIndex = 4;
            this.tabSequenceTool.Text = "Sequence Tool";
            this.tabSequenceTool.UseVisualStyleBackColor = true;
            // 
            // dgSampleList
            // 
            this.dgSampleList.AllowUserToOrderColumns = true;
            this.dgSampleList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgSampleList.Location = new System.Drawing.Point(20, 144);
            this.dgSampleList.Name = "dgSampleList";
            this.dgSampleList.Size = new System.Drawing.Size(349, 615);
            this.dgSampleList.TabIndex = 0;
            // 
            // btnPasteFromExcel
            // 
            this.btnPasteFromExcel.Location = new System.Drawing.Point(20, 115);
            this.btnPasteFromExcel.Name = "btnPasteFromExcel";
            this.btnPasteFromExcel.Size = new System.Drawing.Size(75, 23);
            this.btnPasteFromExcel.TabIndex = 1;
            this.btnPasteFromExcel.Text = "Paste...";
            this.btnPasteFromExcel.UseVisualStyleBackColor = true;
            this.btnPasteFromExcel.Click += new System.EventHandler(this.btnPasteFromExcel_Click);
            // 
            // btnGenSequence
            // 
            this.btnGenSequence.Location = new System.Drawing.Point(592, 478);
            this.btnGenSequence.Name = "btnGenSequence";
            this.btnGenSequence.Size = new System.Drawing.Size(75, 23);
            this.btnGenSequence.TabIndex = 2;
            this.btnGenSequence.Text = "button5";
            this.btnGenSequence.UseVisualStyleBackColor = true;
            this.btnGenSequence.Click += new System.EventHandler(this.btnGenSequence_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(889, 478);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(75, 23);
            this.button5.TabIndex = 3;
            this.button5.Text = "button5";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1559, 811);
            this.Controls.Add(this.tabMainForm);
            this.MainMenuStrip = this.mnuPeptideRatios;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "TrackIN";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.tabMainForm.ResumeLayout(false);
            this.tabHomePage.ResumeLayout(false);
            this.tabNoiseAnalysis.ResumeLayout(false);
            this.tabNoiseAnalysis.PerformLayout();
            this.stsNoiseAnalysis.ResumeLayout(false);
            this.stsNoiseAnalysis.PerformLayout();
            this.mnuNoiseAnalysis.ResumeLayout(false);
            this.mnuNoiseAnalysis.PerformLayout();
            this.tabPeptideRatios.ResumeLayout(false);
            this.tabPeptideRatios.PerformLayout();
            this.stsPeptideRatios.ResumeLayout(false);
            this.stsPeptideRatios.PerformLayout();
            this.mnuPeptideRatios.ResumeLayout(false);
            this.mnuPeptideRatios.PerformLayout();
            this.tabTEST.ResumeLayout(false);
            this.tabSequenceTool.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgSampleList)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnTEST;
        private System.Windows.Forms.Button btnTEST2;
        private ZedGraph.ZedGraphControl zedGraphControlTest;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.TabControl tabMainForm;
        private System.Windows.Forms.TabPage tabHomePage;
        private System.Windows.Forms.TabPage tabTEST;
        private System.Windows.Forms.TabPage tabPeptideRatios;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.MenuStrip mnuPeptideRatios;
        private System.Windows.Forms.ToolStripMenuItem mnuRatioSelection;
        private ZedGraph.ZedGraphControl graphPeptideRatios;
        private System.Windows.Forms.StatusStrip stsPeptideRatios;
        private System.Windows.Forms.ToolStripStatusLabel stsPeptideRatiosLabel;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.TabPage tabNoiseAnalysis;
        private System.Windows.Forms.MenuStrip mnuNoiseAnalysis;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mnuAnalizeFromSkyline;
        private System.Windows.Forms.ToolStripMenuItem mnuNASelectMSRunsToAnalyze;
        private System.Windows.Forms.Button btnMakeLinkerToSkyline;
        private System.Windows.Forms.StatusStrip stsNoiseAnalysis;
        private System.Windows.Forms.ToolStripStatusLabel stsNoiseAnalysisLabel;
        private ZedGraph.ZedGraphControl graphNoiseAnalysis;
        private System.Windows.Forms.TabPage tabSequenceTool;
        private System.Windows.Forms.DataGridView dgSampleList;
        private System.Windows.Forms.Button btnPasteFromExcel;
        private System.Windows.Forms.Button btnGenSequence;
        private System.Windows.Forms.Button button5;
    }
}

