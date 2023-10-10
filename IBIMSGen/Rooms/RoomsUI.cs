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
        public List<ElementId> viewTempsIds;
        public ElementId sectionVTId, celingVTId, floor1VTId, floor2VTId, floor3VTId;
        public string sheetName, sectionName, ceilingName, floor1Name, floor2Name, floor3Name, sheetNumber;
        public RoomsUI(Document doc, List<ElementId> viewTempsIds)
        {
            InitializeComponent();
            this.viewTempsIds = viewTempsIds;
            this.comboBox1.Items.AddRange(this.viewTempsIds.Select(x => doc.GetElement(x).Name).ToArray());
            this.comboBox2.Items.AddRange(this.viewTempsIds.Select(x => doc.GetElement(x).Name).ToArray());
            this.comboBox3.Items.AddRange(this.viewTempsIds.Select(x => doc.GetElement(x).Name).ToArray());
            this.comboBox4.Items.AddRange(this.viewTempsIds.Select(x => doc.GetElement(x).Name).ToArray());
            this.comboBox5.Items.AddRange(this.viewTempsIds.Select(x => doc.GetElement(x).Name).ToArray());
            this.sectionVTId = null; this.celingVTId = null; this.floor1VTId = null; this.floor2VTId = null; this.floor3VTId = null;
            this.sheetName = null; this.sectionName = null; this.ceilingName = null; this.floor1Name = null; this.floor2Name = null; this.floor3Name = null;
            this.sheetName = null;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            this.DialogResult = DialogResult.Cancel;
        }


        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            comboBox1.Enabled = checkBox1.Checked;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            comboBox2.Enabled = checkBox2.Checked;

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            comboBox3.Enabled = checkBox3.Checked;

        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            comboBox4.Enabled = checkBox4.Checked;

        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            comboBox5.Enabled = checkBox5.Checked;

        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            textBox7.Enabled = checkBox6.Checked;
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            textBox2.Enabled = checkBox7.Checked;
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            textBox3.Enabled = checkBox8.Checked;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked) { this.sectionVTId = this.viewTempsIds[comboBox1.SelectedIndex]; }
            if (checkBox2.Checked) { this.celingVTId = this.viewTempsIds[comboBox2.SelectedIndex]; }
            if (checkBox3.Checked) { this.floor1VTId = this.viewTempsIds[comboBox3.SelectedIndex]; }
            if (checkBox4.Checked) { this.floor2VTId = this.viewTempsIds[comboBox4.SelectedIndex]; }
            if (checkBox5.Checked) { this.floor3VTId = this.viewTempsIds[comboBox5.SelectedIndex]; }

            if (checkBox6.Checked) { this.sheetName = this.textBox7.Text; }
            if (checkBox7.Checked) { this.sectionName = this.textBox2.Text; }
            if (checkBox8.Checked) { this.ceilingName = this.textBox3.Text; }
            if (checkBox9.Checked) { this.floor1Name = this.textBox4.Text; }
            if (checkBox10.Checked) { this.floor2Name = this.textBox5.Text; }
            if (checkBox11.Checked) { this.floor3Name = this.textBox6.Text; }

            if (checkBox12.Checked) { this.sheetNumber = this.textBox8.Text; }



            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            textBox4.Enabled = checkBox9.Checked;

        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            textBox5.Enabled = checkBox10.Checked;

        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            textBox6.Enabled = checkBox11.Checked;

        }

        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {
            textBox8.Enabled = checkBox12.Checked;

        }

    }
}
