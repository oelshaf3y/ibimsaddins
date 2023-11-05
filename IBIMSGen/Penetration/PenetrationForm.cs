using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace IBIMSGen.Penetration
{
    public partial class PenetrationForm : System.Windows.Forms.Form
    {
        public List<string> linksnames;
        public List<int> linksinds = new List<int>();
        public bool sel, click;
        List<FamilySymbol> familySymbols;
        List<string> families;
        public FamilySymbol famSymb;
        IList<FamilySymbol> symbols;
        public PenetrationForm(List<string> linksnames, List<FamilySymbol> famSymbols, List<string> families)
        {
            InitializeComponent();
            this.linksnames = linksnames;
            click = false;
            this.familySymbols = famSymbols;
            this.families = families;
            comboBox1.Items.AddRange(this.families.Select(x => x).Distinct().ToArray());
            comboBox3.Items.Add("Pipe");
            comboBox3.Items.Add("Duct");
            comboBox3.Items.Add("Cable Tray");
            comboBox3.Items.Add("Conduit");
        }

        private void Form9_Load(object sender, EventArgs e)
        {
            foreach (string s in linksnames)
            {
                checkedListBox1.Items.Add(s, true);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex > 0)
            {
                famSymb = this.symbols.Where(x => x.Name == comboBox2.SelectedItem.ToString() && x.Family.Name == comboBox1.SelectedItem.ToString()).FirstOrDefault();
            }
            foreach (int i in checkedListBox1.CheckedIndices)
            {
                linksinds.Add(i);
            }
            sel = checkBox1.Checked;
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            sel = checkBox1.Checked;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

            label1.Visible = radioButton2.Checked;
            label2.Visible = radioButton2.Checked;
            label3.Visible = radioButton2.Checked;
            comboBox1.Visible = radioButton2.Checked;
            comboBox2.Visible = radioButton2.Checked;
            comboBox3.Visible = radioButton2.Checked;
        }

        private void label5_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.linkedin.com/in/oelshaf3y");

        }

        private void label6_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.linkedin.com/in/tarek-mahmoud-ahmed-103041204");

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            symbols = this.familySymbols.Where(x => x.Family.Name == comboBox1.SelectedItem.ToString()).ToList();
            comboBox2.Items.AddRange(symbols.Select(x => x.Name).ToList().ToArray());

        }


    }
}
