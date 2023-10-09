using Autodesk.Revit.DB;
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

namespace IBIMSGen.Rooms
{
    public partial class RoomsUI : System.Windows.Forms.Form
    {
        public int visualStyle, scale;
        public ViewDetailLevel DetailLevel;
        public RoomsUI()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            this.DialogResult = DialogResult.Cancel;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (radioButton5.Checked)
            {
                try
                {

                    this.scale = Convert.ToInt32(textBox1.Text);
                }
                catch
                {
                    TaskDialog.Show("Error", "Please make sure you typed a proper scale");
                }
            }
            else
            {
                this.scale = getscale();
            }
            this.DetailLevel = getDetailLevel();
            this.visualStyle = getVisualStyle();
            this.Close();
            this.DialogResult = DialogResult.OK;
        }

        private int getVisualStyle()
        {
            if (radioButton9.Checked) return 1;
            else if (radioButton10.Checked) return 2;
            else if (radioButton11.Checked) return 4;
            else if (radioButton12.Checked) return 5;
            else if (radioButton13.Checked) return 6;
            else { return 4; }
        }

        private ViewDetailLevel getDetailLevel()
        {
            if (radioButton6.Checked) return ViewDetailLevel.Coarse;
            else if (radioButton7.Checked) return ViewDetailLevel.Medium;
            else if (radioButton8.Checked) return ViewDetailLevel.Fine;
            else
            {
                return ViewDetailLevel.Fine;
            }
        }

        private int getscale()
        {
            if (radioButton1.Checked) return 20;
            else if (radioButton2.Checked) { return 25; }
            else if (radioButton3.Checked) { return 50; }
            else if (radioButton4.Checked) { return 100; }


            return 25;
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            textBox1.Visible = radioButton5.Checked;
            label4.Visible = radioButton5.Checked;
        }
    }
}
