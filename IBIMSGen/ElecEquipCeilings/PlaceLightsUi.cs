using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IBIMSGen.ElecEquipCeilings
{
    public partial class PlaceLightsUi : System.Windows.Forms.Form
    {
        public double x, y;
        public int nx, ny;
        public List<FamilySymbol> families;
        public PlaceLightsUi(List<FamilySymbol> Fams)
        {
            InitializeComponent();
            families = Fams;
            comboBox1.Items.AddRange(families.Select(x => x.FamilyName + " - " + x.Name).ToArray());
        }

        private void label3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.linkedin.com/in/oelshaf3y");

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                label1.Text = "Spacing 1-1 Dir.(mm)";
                label2.Text = "Spacing 2-2 Dir.(mm)";
            }
            else
            {
                label1.Text = "N 1-1 Dir";
                label2.Text = "N 2-2 Dir";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            this.DialogResult = DialogResult.Cancel;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.textBox1.Text.Trim().Length == 0 || this.textBox2.Text.Trim().Length == 0)
            {
                MessageBox.Show("Please Don't Leave any field empty!.");
                return;
            }

            if (this.radioButton2.Checked)
            {
                try
                {
                    nx = Convert.ToInt32(this.textBox1.Text);
                    ny = Convert.ToInt32(this.textBox2.Text);
                }
                catch
                {
                    MessageBox.Show("Number of items must be INTEGERS!.");
                }
            }
            else
            {
                try
                {
                    x = Convert.ToDouble(textBox1.Text.Trim());
                    y = Convert.ToDouble(textBox2.Text.Trim());
                }
                catch
                {
                    MessageBox.Show("Please insert numbers for spacing, no letters or characters are allowed!.");
                }
            }

            this.Close();
            this.DialogResult = DialogResult.OK;
        }
    }
}
