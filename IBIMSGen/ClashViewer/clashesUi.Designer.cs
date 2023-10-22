namespace IBIMSGen.ClashViewer
{
    partial class clashesUi
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(clashesUi));
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.inspectCheckBox = new System.Windows.Forms.CheckBox();
            this.OkBut = new System.Windows.Forms.Button();
            this.cancelBut = new System.Windows.Forms.Button();
            this.clashBallCheckBox = new System.Windows.Forms.CheckBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.create3DCheckBox = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.collectBallCheckBox = new System.Windows.Forms.CheckBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.White;
            this.button1.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Location = new System.Drawing.Point(308, 87);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Browse";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(375, 39);
            this.label1.TabIndex = 1;
            this.label1.Text = "this app should create views for each clash point exported from navisworks.\r\nfor " +
    "this app to work you should export tests results (Separated or combined)\r\nas xml" +
    ", then import the xml file below.\r\n";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 71);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(74, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "XML File Path:";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(12, 87);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(273, 20);
            this.textBox1.TabIndex = 3;
            this.toolTip1.SetToolTip(this.textBox1, "you can copy/paste path manually, make sure path ends with .xml");
            // 
            // inspectCheckBox
            // 
            this.inspectCheckBox.AutoSize = true;
            this.inspectCheckBox.Location = new System.Drawing.Point(15, 123);
            this.inspectCheckBox.Name = "inspectCheckBox";
            this.inspectCheckBox.Size = new System.Drawing.Size(102, 17);
            this.inspectCheckBox.TabIndex = 4;
            this.inspectCheckBox.Text = "Inspect Clashes";
            this.toolTip1.SetToolTip(this.inspectCheckBox, "Use this options to review clashes instead of creating\r\nseparate 3D Views for eac" +
        "h clash.");
            this.inspectCheckBox.UseVisualStyleBackColor = true;
            this.inspectCheckBox.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // OkBut
            // 
            this.OkBut.BackColor = System.Drawing.Color.White;
            this.OkBut.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.OkBut.FlatAppearance.BorderSize = 0;
            this.OkBut.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.OkBut.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.OkBut.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.OkBut.Location = new System.Drawing.Point(289, 208);
            this.OkBut.Name = "OkBut";
            this.OkBut.Size = new System.Drawing.Size(75, 23);
            this.OkBut.TabIndex = 5;
            this.OkBut.Text = "OK";
            this.OkBut.UseVisualStyleBackColor = false;
            this.OkBut.Click += new System.EventHandler(this.button2_Click);
            // 
            // cancelBut
            // 
            this.cancelBut.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.cancelBut.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelBut.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.cancelBut.FlatAppearance.BorderSize = 0;
            this.cancelBut.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Red;
            this.cancelBut.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.cancelBut.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cancelBut.Location = new System.Drawing.Point(188, 208);
            this.cancelBut.Name = "cancelBut";
            this.cancelBut.Size = new System.Drawing.Size(75, 23);
            this.cancelBut.TabIndex = 6;
            this.cancelBut.Text = "Cancel";
            this.cancelBut.UseVisualStyleBackColor = false;
            // 
            // clashBallCheckBox
            // 
            this.clashBallCheckBox.AutoSize = true;
            this.clashBallCheckBox.Location = new System.Drawing.Point(3, 3);
            this.clashBallCheckBox.Name = "clashBallCheckBox";
            this.clashBallCheckBox.Size = new System.Drawing.Size(71, 17);
            this.clashBallCheckBox.TabIndex = 7;
            this.clashBallCheckBox.Text = "Clash Ball";
            this.toolTip1.SetToolTip(this.clashBallCheckBox, "Create a ball on clash point exported from navisworks,\r\nclash ball will have all " +
        "information about the clash in it\'s\r\nparameters.");
            this.clashBallCheckBox.UseVisualStyleBackColor = true;
            this.clashBallCheckBox.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged);
            // 
            // create3DCheckBox
            // 
            this.create3DCheckBox.AutoSize = true;
            this.create3DCheckBox.Location = new System.Drawing.Point(15, 146);
            this.create3DCheckBox.Name = "create3DCheckBox";
            this.create3DCheckBox.Size = new System.Drawing.Size(105, 17);
            this.create3DCheckBox.TabIndex = 8;
            this.create3DCheckBox.Text = "Create 3D Views";
            this.toolTip1.SetToolTip(this.create3DCheckBox, "Create 3D views for each clash");
            this.create3DCheckBox.UseVisualStyleBackColor = true;
            this.create3DCheckBox.CheckedChanged += new System.EventHandler(this.checkBox3_CheckedChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(178, 255);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(171, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "© Omar O.Elshafey | I-BIMS 2023";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.textBox2);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.clashBallCheckBox);
            this.panel1.Location = new System.Drawing.Point(13, 189);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(136, 76);
            this.panel1.TabIndex = 10;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(93, 48);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(23, 13);
            this.label6.TabIndex = 12;
            this.label6.Text = "mm";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(3, 45);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(84, 20);
            this.textBox2.TabIndex = 11;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(3, 29);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(58, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Ball Radius";
            // 
            // comboBox1
            // 
            this.comboBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(164, 142);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 21);
            this.comboBox1.TabIndex = 9;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(164, 126);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(98, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Coordinaates Units";
            // 
            // collectBallCheckBox
            // 
            this.collectBallCheckBox.AutoSize = true;
            this.collectBallCheckBox.Location = new System.Drawing.Point(15, 169);
            this.collectBallCheckBox.Name = "collectBallCheckBox";
            this.collectBallCheckBox.Size = new System.Drawing.Size(82, 17);
            this.collectBallCheckBox.TabIndex = 11;
            this.collectBallCheckBox.Text = "Collect Balls";
            this.collectBallCheckBox.UseVisualStyleBackColor = true;
            // 
            // clashesUi
            // 
            this.AcceptButton = this.OkBut;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.CancelButton = this.cancelBut;
            this.ClientSize = new System.Drawing.Size(395, 277);
            this.Controls.Add(this.collectBallCheckBox);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.create3DCheckBox);
            this.Controls.Add(this.cancelBut);
            this.Controls.Add(this.OkBut);
            this.Controls.Add(this.inspectCheckBox);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "clashesUi";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "I-BIMS Clash Finder";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Label label1;
        public System.Windows.Forms.Label label2;
        public System.Windows.Forms.Button button1;
        public System.Windows.Forms.TextBox textBox1;
        public System.Windows.Forms.CheckBox inspectCheckBox;
        public System.Windows.Forms.Button OkBut;
        public System.Windows.Forms.Button cancelBut;
        public System.Windows.Forms.CheckBox clashBallCheckBox;
        public System.Windows.Forms.ToolTip toolTip1;
        public System.Windows.Forms.CheckBox create3DCheckBox;
        public System.Windows.Forms.Label label3;
        public System.Windows.Forms.Panel panel1;
        public System.Windows.Forms.ComboBox comboBox1;
        public System.Windows.Forms.Label label4;
        public System.Windows.Forms.TextBox textBox2;
        public System.Windows.Forms.Label label5;
        public System.Windows.Forms.Label label6;
        public System.Windows.Forms.CheckBox collectBallCheckBox;
    }
}