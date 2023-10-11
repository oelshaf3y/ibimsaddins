using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using View = Autodesk.Revit.DB.View;
namespace IBIMSGen.Rooms
{
    public partial class ParentView : System.Windows.Forms.Form
    {
        public List<View> views;
        public ParentView(List<View> views,string roomName,string calloutType)
        {
            InitializeComponent();
            this.views = views;
            this.label3.Text = calloutType;
            this.label5.Text= roomName;
            this.label6.Text = $"i found {views.Count} views.\r\nplease choose a parent view from the list below.\r\nthanks.";
            this.comboBox1.Items.AddRange(this.views.Select(v => v.Name).ToArray());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Please Select a parent view.");
                return;
            }
            this.Close();
            this.DialogResult = DialogResult.OK;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult= DialogResult.Cancel;
            this.Close();
        }
    }
}
