namespace IBIMSGen.Hangers
{
    partial class HangerUC
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.checkListPanel = new System.Windows.Forms.Panel();
            this.worksetNames = new System.Windows.Forms.CheckedListBox();
            this.sizesPanel = new System.Windows.Forms.Panel();
            this.allSizesSpacing = new System.Windows.Forms.TextBox();
            this.fixedSpacingLabel = new System.Windows.Forms.Label();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.copyFromCB = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.useDims = new System.Windows.Forms.CheckBox();
            this.checkListPanel.SuspendLayout();
            this.sizesPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // checkListPanel
            // 
            this.checkListPanel.Controls.Add(this.label2);
            this.checkListPanel.Controls.Add(this.worksetNames);
            this.checkListPanel.Location = new System.Drawing.Point(0, 0);
            this.checkListPanel.Name = "checkListPanel";
            this.checkListPanel.Size = new System.Drawing.Size(209, 367);
            this.checkListPanel.TabIndex = 0;
            // 
            // worksetNames
            // 
            this.worksetNames.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.worksetNames.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.worksetNames.FormattingEnabled = true;
            this.worksetNames.Location = new System.Drawing.Point(0, 22);
            this.worksetNames.Name = "worksetNames";
            this.worksetNames.Size = new System.Drawing.Size(209, 345);
            this.worksetNames.TabIndex = 0;
            // 
            // sizesPanel
            // 
            this.sizesPanel.Controls.Add(this.useDims);
            this.sizesPanel.Controls.Add(this.allSizesSpacing);
            this.sizesPanel.Controls.Add(this.fixedSpacingLabel);
            this.sizesPanel.Controls.Add(this.dgv);
            this.sizesPanel.Location = new System.Drawing.Point(215, 3);
            this.sizesPanel.Name = "sizesPanel";
            this.sizesPanel.Size = new System.Drawing.Size(300, 311);
            this.sizesPanel.TabIndex = 1;
            // 
            // allSizesSpacing
            // 
            this.allSizesSpacing.Location = new System.Drawing.Point(140, 156);
            this.allSizesSpacing.Name = "allSizesSpacing";
            this.allSizesSpacing.Size = new System.Drawing.Size(100, 20);
            this.allSizesSpacing.TabIndex = 2;
            this.allSizesSpacing.Text = "2000";
            this.allSizesSpacing.Visible = false;
            // 
            // fixedSpacingLabel
            // 
            this.fixedSpacingLabel.AutoSize = true;
            this.fixedSpacingLabel.Location = new System.Drawing.Point(48, 159);
            this.fixedSpacingLabel.Name = "fixedSpacingLabel";
            this.fixedSpacingLabel.Size = new System.Drawing.Size(86, 13);
            this.fixedSpacingLabel.TabIndex = 1;
            this.fixedSpacingLabel.Text = "Spacing in (mm):";
            this.fixedSpacingLabel.Visible = false;
            // 
            // dgv
            // 
            this.dgv.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Location = new System.Drawing.Point(0, 0);
            this.dgv.Name = "dgv";
            this.dgv.RowHeadersVisible = false;
            this.dgv.Size = new System.Drawing.Size(300, 278);
            this.dgv.TabIndex = 0;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.label1);
            this.panel3.Controls.Add(this.copyFromCB);
            this.panel3.Location = new System.Drawing.Point(215, 320);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(300, 47);
            this.panel3.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(0, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(107, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Duplicate table from:";
            this.label1.Visible = false;
            // 
            // copyFromCB
            // 
            this.copyFromCB.BackColor = System.Drawing.Color.White;
            this.copyFromCB.Enabled = false;
            this.copyFromCB.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.copyFromCB.FormattingEnabled = true;
            this.copyFromCB.Location = new System.Drawing.Point(113, 21);
            this.copyFromCB.Name = "copyFromCB";
            this.copyFromCB.Size = new System.Drawing.Size(104, 21);
            this.copyFromCB.TabIndex = 0;
            this.copyFromCB.Visible = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 3);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(191, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Check Link Name Where System Exists";
            // 
            // useDims
            // 
            this.useDims.AutoSize = true;
            this.useDims.Location = new System.Drawing.Point(3, 284);
            this.useDims.Name = "useDims";
            this.useDims.Size = new System.Drawing.Size(214, 17);
            this.useDims.TabIndex = 2;
            this.useDims.Text = "Use these dimensions with my selection";
            this.useDims.UseVisualStyleBackColor = true;
            // 
            // HangerUC
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.sizesPanel);
            this.Controls.Add(this.checkListPanel);
            this.Name = "HangerUC";
            this.Size = new System.Drawing.Size(520, 370);
            this.checkListPanel.ResumeLayout(false);
            this.checkListPanel.PerformLayout();
            this.sizesPanel.ResumeLayout(false);
            this.sizesPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.DataGridView dgv;
        public System.Windows.Forms.Label fixedSpacingLabel;
        public System.Windows.Forms.TextBox allSizesSpacing;
        public System.Windows.Forms.Panel checkListPanel;
        public System.Windows.Forms.Panel sizesPanel;
        public System.Windows.Forms.Panel panel3;
        public System.Windows.Forms.CheckedListBox worksetNames;
        public System.Windows.Forms.ComboBox copyFromCB;
        public System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox useDims;
    }
}
