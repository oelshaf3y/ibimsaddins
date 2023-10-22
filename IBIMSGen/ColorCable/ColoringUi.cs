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

namespace IBIMSGen.ColorCable
{
    public partial class coloringUi : Form
    {
        public coloringUi()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (rtb.Text.Length != 0 || btb.Text.Length != 0 || ytb.Text.Length != 0 || stb.Text.Length != 0)
            {

                DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                TaskDialog.Show("Error", "All fields must have a value.");
            }
        }

        private void label8_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.linkedin.com/in/oelshaf3y");

        }
    }
}
