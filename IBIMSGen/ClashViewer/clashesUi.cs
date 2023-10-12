using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IBIMSGen.ClashViewer
{
    public partial class clashesUi : Form
    {
        OpenFileDialog dialog = new OpenFileDialog();
        public clashesUi(bool ins, bool coll)
        {
            InitializeComponent();
            collectBallCheckBox.Visible = coll;
            clashBallCheckBox.Visible = ins;
            clashBallCheckBox.Checked = !ins;
            inspectCheckBox.Visible = ins;
            create3DCheckBox.Visible = ins;
            this.comboBox1.Items.Add("Millimeter (mm)");//0 >> /304.8
            this.comboBox1.Items.Add("Centimeter (cm)");//1 >> /30.48
            this.comboBox1.Items.Add("Meter (M)");//2 >> /0.3048
            this.comboBox1.Items.Add("Feet (ft)");//3 >> *1
            this.comboBox1.Items.Add("Inch (in)");//4 >> /12
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (collectBallCheckBox.Checked)
            {
                this.Close();
                this.DialogResult = DialogResult.OK;
            }
            else if (textBox1.Text.Trim().Length == 0)
            {
                TaskDialog.Show("Error", "Error!!\nNo xml file selected. please select the test report file and try again.");

            }
            else if (clashBallCheckBox.Checked)
            {
                double dx;
                if (textBox2.Text.Trim().Length == 0)
                {
                    TaskDialog.Show("Error", "Ball Radius can not be empty or 0");
                }
                else if (!double.TryParse(textBox2.Text.Trim(), out dx))
                {
                    TaskDialog.Show("Error", "Ball Radius must be a real number!");

                }
                else if (comboBox1.SelectedIndex == -1)
                {
                    TaskDialog.Show("Error", "please select proper coordinates units!");

                }
                else
                {
                    this.Close();
                    this.DialogResult = DialogResult.OK;
                }
            }
            else
            {
                if (comboBox1.SelectedIndex == -1)
                {
                    TaskDialog.Show("Error", "please select proper coordinates units!");
                }
                else if (clashBallCheckBox.Checked || create3DCheckBox.Checked || inspectCheckBox.Checked)
                {
                    this.Close();
                    this.DialogResult = DialogResult.OK;
                }
                else
                {
                    TaskDialog.Show("Error", "You have to select at least one option\nClash Ball or 3D Views.");
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dialog.Title = "Open clash report as xml";
            dialog.Filter = "Xml Files (*.xml)|*.xml";
            dialog.ShowDialog();
            textBox1.Text = dialog.FileName;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (inspectCheckBox.Checked)
            {
                clashBallCheckBox.Checked = false;
                create3DCheckBox.Checked = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            inspectCheckBox.Checked = false;
            clashBallCheckBox.Checked = false;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            inspectCheckBox.Checked = false;
            create3DCheckBox.Checked = false;
        }

        private void label3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.linkedin.com/in/oelshaf3y");
        }
    }
}
