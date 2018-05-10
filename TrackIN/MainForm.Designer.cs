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
            this.tabTEST = new System.Windows.Forms.TabPage();
            this.tabMainForm.SuspendLayout();
            this.tabTEST.SuspendLayout();
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
            this.zedGraphControlTest.Location = new System.Drawing.Point(8, 55);
            this.zedGraphControlTest.Name = "zedGraphControlTest";
            this.zedGraphControlTest.ScrollGrace = 0D;
            this.zedGraphControlTest.ScrollMaxX = 0D;
            this.zedGraphControlTest.ScrollMaxY = 0D;
            this.zedGraphControlTest.ScrollMaxY2 = 0D;
            this.zedGraphControlTest.ScrollMinX = 0D;
            this.zedGraphControlTest.ScrollMinY = 0D;
            this.zedGraphControlTest.ScrollMinY2 = 0D;
            this.zedGraphControlTest.Size = new System.Drawing.Size(1520, 722);
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
            this.tabMainForm.Controls.Add(this.tabTEST);
            this.tabMainForm.Location = new System.Drawing.Point(0, 0);
            this.tabMainForm.Name = "tabMainForm";
            this.tabMainForm.SelectedIndex = 0;
            this.tabMainForm.Size = new System.Drawing.Size(1575, 850);
            this.tabMainForm.TabIndex = 10;
            // 
            // tabHomePage
            // 
            this.tabHomePage.Location = new System.Drawing.Point(4, 22);
            this.tabHomePage.Name = "tabHomePage";
            this.tabHomePage.Padding = new System.Windows.Forms.Padding(3);
            this.tabHomePage.Size = new System.Drawing.Size(1567, 824);
            this.tabHomePage.TabIndex = 0;
            this.tabHomePage.Text = "HomePage";
            this.tabHomePage.UseVisualStyleBackColor = true;
            // 
            // tabTEST
            // 
            this.tabTEST.Controls.Add(this.button3);
            this.tabTEST.Controls.Add(this.button4);
            this.tabTEST.Controls.Add(this.btnTEST2);
            this.tabTEST.Controls.Add(this.zedGraphControlTest);
            this.tabTEST.Controls.Add(this.btnTEST);
            this.tabTEST.Location = new System.Drawing.Point(4, 22);
            this.tabTEST.Name = "tabTEST";
            this.tabTEST.Padding = new System.Windows.Forms.Padding(3);
            this.tabTEST.Size = new System.Drawing.Size(1567, 824);
            this.tabTEST.TabIndex = 1;
            this.tabTEST.Text = "TEST only";
            this.tabTEST.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1559, 811);
            this.Controls.Add(this.tabMainForm);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "TrackIN";
            this.tabMainForm.ResumeLayout(false);
            this.tabTEST.ResumeLayout(false);
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
    }
}

