namespace AquityTimestamp
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
            this.btShowTimestamp = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btShowTimestamp
            // 
            this.btShowTimestamp.Location = new System.Drawing.Point(166, 121);
            this.btShowTimestamp.Name = "btShowTimestamp";
            this.btShowTimestamp.Size = new System.Drawing.Size(119, 37);
            this.btShowTimestamp.TabIndex = 0;
            this.btShowTimestamp.Text = "Call \"Datetime.now()\"";
            this.btShowTimestamp.UseVisualStyleBackColor = true;
            this.btShowTimestamp.Click += new System.EventHandler(this.btShowTimestamp_Click);
            // 
            // fmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(453, 296);
            this.Controls.Add(this.btShowTimestamp);
            this.Name = "fmMain";
            this.Text = "Timestamp Test";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btShowTimestamp;
    }
}

