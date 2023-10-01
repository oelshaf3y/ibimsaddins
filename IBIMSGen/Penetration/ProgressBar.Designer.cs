namespace IBIMSGen
{
    partial class ProgressBar
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
            this.progresBarRatio = new System.Windows.Forms.ProgressBar();
            this.Lb = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // progresBarRatio
            // 
            this.progresBarRatio.Location = new System.Drawing.Point(9, 21);
            this.progresBarRatio.Margin = new System.Windows.Forms.Padding(2);
            this.progresBarRatio.Name = "progresBarRatio";
            this.progresBarRatio.Size = new System.Drawing.Size(380, 29);
            this.progresBarRatio.Step = 1;
            this.progresBarRatio.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progresBarRatio.TabIndex = 0;
            // 
            // Lb
            // 
            this.Lb.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.Lb.AutoSize = true;
            this.Lb.BackColor = System.Drawing.Color.Transparent;
            this.Lb.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Lb.Font = new System.Drawing.Font("Lucida Bright", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Lb.Location = new System.Drawing.Point(339, 76);
            this.Lb.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Lb.Name = "Lb";
            this.Lb.Size = new System.Drawing.Size(50, 16);
            this.Lb.TabIndex = 1;
            this.Lb.Text = "label1";
            this.Lb.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            this.Lb.UseMnemonic = false;
            // 
            // ProgressBar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(398, 101);
            this.ControlBox = false;
            this.Controls.Add(this.Lb);
            this.Controls.Add(this.progresBarRatio);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProgressBar";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Progress_Bar";
            this.Load += new System.EventHandler(this.Progress_Bar_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.ProgressBar progresBarRatio;
        public System.Windows.Forms.Label Lb;
    }
}