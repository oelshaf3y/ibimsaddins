using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Color = System.Drawing.Color;
using Control = System.Windows.Forms.Control;
using Point = System.Drawing.Point;

namespace IBIMSGen.Hangers
{

    public partial class HangersFM : System.Windows.Forms.Form
    {

        public bool canc, ook, selc;
        public int dup, frin, toin, linkIndex = -1;
        public List<string> linksNames, levelsNames;
        public List<List<string>> AllworksetsNames;
        public List<List<Dictionary<string, double>>> AllworksetsDIMS;
        public UserControl Duc, Wuc, CWuc, DRuc, FFuc, CTuc;
        public Button lastButton = null;
        public List<double> DRdias, WSdias, CHWdias, Firedias, DRspcs, WSspcs, CHWspcs, Firespcs;
        public List<FamilySymbol> familySymbols;

        public HangersFM(List<string> linksNames, List<string> levelsNames, List<FamilySymbol> familySymbols)
        {
            InitializeComponent();
            DRdias = new List<double>() { 20, 25, 32, 40, 50, 75, 110, 125, 130 };
            WSdias = new List<double>() { 20, 25, 32, 40, 50, 75, 110, 125, 130 };
            CHWdias = new List<double>() { 15, 20, 25, 32, 40, 50, 65, 80, 100, 125, 150, 200, 250, 300, 350, 400 };
            Firedias = new List<double>() {  20, 25, 32, 40, 50, 65, 80, 90, 100, 125, 150, 200, 250, 300, 350 };
            DRspcs = new List<double>() { 1000, 1000, 1000, 1000, 1065, 1370, 1525, 1680, 1830 };
            WSspcs = new List<double>() { 1000, 1000, 1000, 1000, 1065, 1370, 1525, 1680, 1830 };
            CHWspcs = new List<double>() { 2100, 2100, 2100, 2400, 2700, 3100, 3400, 3600, 4300, 4300, 5100, 5800, 6100, 7000, 7600, 8200 };
            Firespcs = new List<double>() {  2400, 2400, 2700, 3000, 3000, 3300, 3600, 3700, 3900, 4200, 4500, 4500, 4500, 4500, 4500 };
            AllworksetsDIMS = new List<List<Dictionary<string, double>>>();
            AllworksetsNames = new List<List<string>>();
            this.linksNames = linksNames;
            this.levelsNames = levelsNames;
            this.familySymbols = familySymbols;
            Duc = null; Wuc = null; CWuc = null; DRuc = null; FFuc = null; CTuc = null;
        }
        private void Form7_Load(object sender, EventArgs e)
        {
            comboBox1.Items.AddRange(linksNames.ToArray());
            comboBox2.Items.AddRange(levelsNames.ToArray());
            comboBox3.Items.AddRange(levelsNames.ToArray());
            Duc = createUserControl("Duct", true);
            Wuc = createUserControl("WS", false, WSdias, WSspcs);
            CWuc = createUserControl("CHW", false, CHWdias, CHWspcs);
            DRuc = createUserControl("DR", false, DRdias, DRspcs);
            FFuc = createUserControl("FF", false, Firedias, Firespcs);
            CTuc = createUserControl("CT");
            CTuc.Controls.Find("dgv", true).First().Enabled = false;
            lastButton = CT;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = comboBox3.Items.Count - 1;
            useLink.Checked = true;
            buttonClicked(ductButton.Name);
        }

        private void D_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;
            foreach (DataGridViewRow row in dgv.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.ErrorText != "")
                    {
                        cell.ErrorText = "";
                    }
                }
            }
        }
        public UserControl getActiveUC()
        {
            foreach (UserControl us in panel1.Controls)
            {
                if (us.Visible == true) return us;
            }
            return null;
        }

        public UserControl createUserControl(string name, bool vis = false, List<double> sizes = null, List<double> spaces = null)
        {

            UserControl userControl = new HangerUC();
            userControl.Name = name;
            panel1.Controls.Add(userControl);
            userControl.Dock = DockStyle.Fill;
            ComboBox hangersFamilies = userControl.Controls.Find("hangerFamily", true).First() as ComboBox;
            hangersFamilies.Items.AddRange(familySymbols.Select(x => x.FamilyName).Distinct().ToArray());

            if (name == "WS" || name == "CHW" || name == "DR" || name.Contains("stem"))
            {
                userControl.Controls.Find("label4", true).First().Visible = true;
                userControl.Controls.Find("label4", true).First().Text = "larger than 200mm";
                userControl.Controls.Find("label3", true).First().Text = "smaller than 200mm";
                userControl.Controls.Find("hangerFamily2", true).First().Visible = true;
                ((ComboBox)userControl.Controls.Find("hangerFamily2", true).First()).Items.AddRange(familySymbols.Select(x => x.FamilyName).Distinct().ToArray());
            }

            CheckedListBox worksetNames = userControl.Controls.Find("worksetNames", true).First() as CheckedListBox;
            worksetNames.Items.AddRange(linksNames.ToArray());
            ComboBox copyFromCB = userControl.Controls.Find("copyFromCB", true).First() as ComboBox;
            copyFromCB.SelectedIndexChanged += copyFrom_SelectedIndexChanged;
            DataGridView dgv = userControl.Controls.Find("dgv", true).First() as DataGridView;
            dgv.DefaultCellStyle.NullValue = "0";
            dgv.CellLeave += D_CellLeave;
            if (name != "Duct")
            {
                if (sizes == null) { sizes = new List<double>(); }
                if (spaces == null) { spaces = new List<double>(); }
                dgv.Columns.Add("Size", "Size ( mm )");
                dgv.Columns.Add("Spacing", "Spacing ( mm )");
                dgv.Columns[0].Width = 150;
                dgv.Columns[1].Width = 146;
                if (sizes.Count == 0) { dgv.RowCount = 12; }
                else { dgv.RowCount = sizes.Count; }
                for (int i = 0; i < sizes.Count; i++)
                {
                    if (sizes[i] != null) dgv[0, i].Value = sizes[i]; else dgv[0, i].Value = "0";
                    if (spaces[i] != null) dgv[1, i].Value = spaces[i]; else dgv[1, i].Value = "0";
                }
            }
            else
            {
                sizes = new List<double>() { 1, 533.4, 838.2, 1041.4, 1524 };
                List<string> sizes2 = new List<string>() { "508", "812.8", "1016", "1524", "10000" };
                spaces = new List<double>() { 3000, 2500, 2500, 2000, 1500 };
                dgv.Columns.Add("Size1", "Size from (mm)");
                dgv.Columns.Add("Size2", "Size to (mm)");
                dgv.Columns.Add("Spacing", "Spacing (mm)");
                dgv.Columns[0].Width = 98;
                dgv.Columns[1].Width = 98;
                dgv.Columns[2].Width = 100;
                dgv.RowCount = sizes.Count;
                for (int i = 0; i < sizes.Count; i++)
                {
                    dgv[0, i].Value = sizes[i];
                    dgv[1, i].Value = sizes2[i];
                    dgv[2, i].Value = spaces[i];
                }
            }

            return userControl;
        }
        private void copyFrom_SelectedIndexChanged(object sender, EventArgs e)
        {
            dup = ((ComboBox)sender).SelectedIndex;
            DataGridView neww = getActiveUC().Controls.Find("dgv", true).FirstOrDefault() as DataGridView;
            if (dup == 0)
            {
                neww.RowCount = 12;
                for (int i = 0; i < 2; i++)
                {
                    for (int j = 0; j < 12; j++)
                    {
                        neww[i, j].Value = "0";
                    }
                }
            }
            else
            {
                DataGridView old = null;
                foreach (UserControl c in panel1.Controls)
                {
                    if (((ComboBox)sender).Items[dup].ToString().Contains(c.Name.ToString()))
                    {
                        old = c.Controls.Find("dgv", true).FirstOrDefault() as DataGridView;
                        break;
                    }
                }
                neww.RowCount = old.Rows.Count; int cl = 0;
                for (int i = 0; i < 2; i++)
                {
                    for (int j = 0; j < old.Rows.Count; j++)
                    {
                        neww[i, j].Value = old[i, j].Value;
                    }
                }
            }
        }

        private void CT_Click(object sender, EventArgs e)
        {
            if (isValidDims())
            {
                showUserControl("CT");
                buttonClicked(CT.Name);
            }
            else
            {
                MessageBox.Show("Size or Spacing Value must be a number and greater than 0!");
            }
        }

        private void label6_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.linkedin.com/in/tarek-mahmoud-ahmed-103041204");

        }

        private void label5_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.linkedin.com/in/oelshaf3y");

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

            foreach (var control in panel1.Controls.Find("worksetNames", true))
            {
                control.Enabled = useLink.Checked;
            }
            foreach (var control in panel1.Controls.Find("useDims", true))
            {
                control.Enabled = !useLink.Checked;
            }
        }

        bool isValidDims()
        {
            DataGridView activeD = (DataGridView)getActiveUC().Controls.Find("dgv", true)?.FirstOrDefault();
            if (!activeD.Enabled) return true;
            foreach (DataGridViewRow row in activeD.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    double x;
                    bool conv = double.TryParse(cell.Value?.ToString(), out x);
                    if (cell.Value != null && !conv || x < 0)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        void getWorksetDims(string name)
        {
            Control c = panel1.Controls.Find(name, true).First();
            DataGridView d = (DataGridView)c.Controls.Find("dgv", true).FirstOrDefault();
            List<Dictionary<string, double>> newDims = new List<Dictionary<string, double>>();
            if (useLink.Checked)
            {
                if (d.Enabled == true)
                {
                    for (int i = 0; i < d.RowCount; i++)
                    {
                        double size = Convert.ToDouble(d[0, i].Value);
                        double spacing = Convert.ToDouble(d[1, i].Value);
                        Dictionary<string, double> dict = new Dictionary<string, double>();
                        dict.Add("size", size);
                        dict.Add("spacing", spacing);
                        dict.Add("family", ((ComboBox)c.Controls.Find("hangerFamily", true).FirstOrDefault()).SelectedIndex);
                        if (name == "WS" || name == "CHW" || name == "DR" || name.Contains("stem"))
                        {
                            dict.Add("family2", ((ComboBox)c.Controls.Find("hangerFamily2", true).FirstOrDefault()).SelectedIndex);
                        }
                        if (name == "FF") dict.Add("FF", 1);
                        else dict.Add("FF", 0);
                        newDims.Add(dict);
                    }
                    AllworksetsDIMS.Add(newDims);
                }
                else
                {
                    TextBox t = (TextBox)c.Controls.Find("allSizesSpacing", true).FirstOrDefault();
                    Dictionary<string, double> dict = new Dictionary<string, double>();
                    dict.Add("spacing", Convert.ToDouble(t.Text));
                    dict.Add("family", ((ComboBox)c.Controls.Find("hangerFamily", true).FirstOrDefault()).SelectedIndex);
                    if (name == "WS" || name == "CHW" || name == "DR" || name.Contains("stem"))
                    {
                        dict.Add("family2", ((ComboBox)c.Controls.Find("hangerFamily2", true).FirstOrDefault()).SelectedIndex);
                    }
                    if (name == "FF") dict.Add("FF", 1);
                    else dict.Add("FF", 0);
                    newDims.Add(dict);
                    AllworksetsDIMS.Add(newDims);
                }

            }
            else
            {

                CheckBox use = c.Controls.Find("useDims", true).First() as CheckBox;
                if (use.Checked)
                {

                    if (d.Enabled == true)
                    {
                        for (int i = 0; i < d.RowCount; i++)
                        {
                            double size = Convert.ToDouble(d[0, i].Value);
                            double spacing = Convert.ToDouble(d[1, i].Value);
                            Dictionary<string, double> dict = new Dictionary<string, double>();
                            dict.Add("size", size);
                            dict.Add("spacing", spacing);
                            dict.Add("family", ((ComboBox)c.Controls.Find("hangerFamily", true).FirstOrDefault()).SelectedIndex);
                            if (name == "WS" || name == "CHW" || name == "DR" || name.Contains("stem"))
                            {
                                dict.Add("family2", ((ComboBox)c.Controls.Find("hangerFamily2", true).FirstOrDefault()).SelectedIndex);
                            }
                            if (name == "FF") dict.Add("FF", 1);
                            else dict.Add("FF", 0);
                            newDims.Add(dict);
                        }
                        AllworksetsDIMS.Add(newDims);
                    }
                    else
                    {
                        //MessageBox.Show("Info", "Single spacing");
                        TextBox t = (TextBox)c.Controls.Find("allSizesSpacing", true).FirstOrDefault();
                        Dictionary<string, double> dict = new Dictionary<string, double>();
                        dict.Add("spacing", Convert.ToDouble(t.Text));
                        dict.Add("family", ((ComboBox)c.Controls.Find("hangerFamily", true).FirstOrDefault()).SelectedIndex);
                        if (name == "WS" || name == "CHW" || name == "DR" || name.Contains("stem"))
                        {
                            dict.Add("family2", ((ComboBox)c.Controls.Find("hangerFamily2", true).FirstOrDefault()).SelectedIndex);
                        }
                        if (name == "FF") dict.Add("FF", 1);
                        else dict.Add("FF", 0);
                        newDims.Add(dict);
                        AllworksetsDIMS.Add(newDims);
                    }
                }
                else
                {
                    Dictionary<string, double> dict = new Dictionary<string, double>();
                    dict.Add("spacing", 0);
                    dict.Add("family", ((ComboBox)c.Controls.Find("hangerFamily", true).FirstOrDefault()).SelectedIndex);
                    if (name == "WS" || name == "CHW" || name == "DR" || name.Contains("stem"))
                    {
                        dict.Add("family2", ((ComboBox)c.Controls.Find("hangerFamily2", true).FirstOrDefault()).SelectedIndex);
                    }
                    if (name == "FF") dict.Add("FF", 1);
                    else dict.Add("FF", 0);
                    newDims.Add(dict);
                    AllworksetsDIMS.Add(newDims);
                }
            }
        }

        void showUserControl(string name)
        {
            UserControl uc = null;
            foreach (Control c in panel1.Controls)
            {
                if (c is UserControl)
                {
                    if (c.Name == name)
                    {
                        c.Visible = true;
                        uc = (UserControl)c;
                    }
                    else
                    {
                        c.Visible = false;
                    }
                }
            }
            DataGridView dd = uc.Controls.Find("dgv", true).FirstOrDefault() as DataGridView;
            if (dd.Enabled)
            {
                checkBox1.Checked = false;
            }
            else
            {
                checkBox1.Checked = true;
            }
        }
        void buttonClicked(string btname)
        {
            foreach (Control c in panel3.Controls)
            {
                if (c is Button)
                {
                    if (c.Name == btname)
                    {
                        c.ForeColor = Color.Black;
                        c.BackColor = Color.LightSkyBlue;
                    }
                    else
                    {
                        c.ForeColor = Color.Brown;
                        c.BackColor = Color.Gainsboro;
                    }
                }
            }
            ComboBox cbx = getActiveUC().Controls.Find("copyFromCB", true).FirstOrDefault() as ComboBox;
            Label label = getActiveUC().Controls.Find("label1", true).FirstOrDefault() as Label;
            if (getActiveUC().Name.Contains("System"))
            {
                cbx.Enabled = true;
                cbx.Visible = true;
                label.Visible = true;
                cbx.Items.Clear();
                cbx.Items.Add("New Table");
                foreach (Control c in panel3.Controls)
                {
                    if (c is Button && c.Text != "Ducts")
                    {
                        cbx.Items.Add(c.Text);
                    }
                }
            }
            else
            {
                cbx.Enabled = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (linkIndex == -1)
            {
                ook = false;
                button1.DialogResult = DialogResult.None;
                MessageBox.Show("Select a Revit Link before Run.");
            }
            else if (!isValidDims())
            {
                ook = false;
                button1.DialogResult = DialogResult.None;
                MessageBox.Show("Size or Spacing Value must be a number and greater than 0!");
            }
            else
            {
                List<Control> CLBs = new List<Control>();
                CLBs.Add(panel1.Controls.Find("Duct", true).First());
                CLBs.Add(panel1.Controls.Find("WS", true).First());
                CLBs.Add(panel1.Controls.Find("CHW", true).First());
                CLBs.Add(panel1.Controls.Find("DR", true).First());
                CLBs.Add(panel1.Controls.Find("FF", true).First());
                CLBs.Add(panel1.Controls.Find("CT", true).First());
                if (useLink.Checked)
                {

                    foreach (var item in CLBs)
                    {
                        CheckedListBox CLB = item.Controls.Find("worksetNames", true).FirstOrDefault() as CheckedListBox;
                        if (CLB.CheckedItems.Cast<string>().ToList().Count > 0)
                        {
                            AllworksetsNames.Add(CLB.CheckedItems.Cast<string>().ToList());
                            ComboBox comboBox = item.Controls.Find("hangerFamily", true).First() as ComboBox;
                            if (comboBox.SelectedIndex == -1)
                            {
                                MessageBox.Show($"Please select the family for {item.Name}");
                                return;
                            }
                            if (item.Controls.Find("hangerFamily2", true).First().Visible)
                            {
                                ComboBox comboBox2 = item.Controls.Find("hangerFamily2", true).First() as ComboBox;
                                if (comboBox2.SelectedIndex == -1)
                                {
                                    MessageBox.Show($"Please select the larger family for {item.Name}");
                                    return;
                                }
                            }
                        }
                        else
                        {
                            AllworksetsNames.Add(new List<string>());
                        }
                    }
                    if (CLBs.Select(x => x.Controls.Find("worksetNames", true).FirstOrDefault()).Cast<CheckedListBox>().ToList().Count == 0)
                    {
                        MessageBox.Show("Nothing selected!!");
                        return;
                    }
                    getDimensions();
                }
                else
                {
                    foreach (var item in CLBs)
                    {
                        CheckBox useDimension = item.Controls.Find("useDims", true).First() as CheckBox;
                        if (useDimension.Checked)
                        {
                            AllworksetsNames.Add(new List<string>() { "used"});
                        }
                        else
                        {
                            AllworksetsNames.Add(new List<string>());
                        }
                    }
                    getDimensions();
                }
                foreach (Control cont in panel1.Controls)
                {
                    if (cont is UserControl && cont.Name.Contains("System"))
                    {
                        CheckedListBox clb = cont.Controls.Find("worksetNames", true).FirstOrDefault() as CheckedListBox;
                        AllworksetsNames.Add(clb.CheckedItems.Cast<string>().ToList());
                    }
                }

                ook = true;
                button1.DialogResult = DialogResult.OK;
                Close();
            }
            frin = comboBox2.SelectedIndex;
            toin = comboBox3.SelectedIndex;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            canc = true;
        }

        private void getDimensions()
        {
            Control c = panel1.Controls.Find("Duct", true).First();
            DataGridView ductSizesList = c.Controls.Find("dgv", true).First() as DataGridView;
            List<Dictionary<string, double>> DIMS = new List<Dictionary<string, double>>();
            if (AllworksetsNames[0].Count > 0)
            {
                if (ductSizesList.Enabled)
                {
                    for (int j = 0; j < ductSizesList.RowCount; j++)
                    {
                        Dictionary<string, double> dict = new Dictionary<string, double>();
                        dict["from"] = Convert.ToDouble(ductSizesList[0, j].Value.ToString());
                        dict["to"] = Convert.ToDouble(ductSizesList[1, j].Value.ToString());
                        dict["spacing"] = Convert.ToDouble(ductSizesList[2, j].Value.ToString());
                        dict["family"] = ((ComboBox)c.Controls.Find("hangerFamily", true).FirstOrDefault()).SelectedIndex;
                        DIMS.Add(dict);
                    }
                    AllworksetsDIMS.Add(DIMS);
                }
                else
                {
                    TextBox t = (TextBox)c.Controls.Find("allSizesSpacing", true).First();
                    Dictionary<string, double> dict = new Dictionary<string, double>();
                    dict["spacing"] = Convert.ToDouble(t.Text);
                    dict["family"] = ((ComboBox)c.Controls.Find("hangerFamily", true).FirstOrDefault()).SelectedIndex;
                    DIMS.Add(dict);
                    AllworksetsDIMS.Add(DIMS);
                }
            }
            else { AllworksetsDIMS.Add(new List<Dictionary<string, double>>()); }

            if (AllworksetsNames[1].Count > 0) getWorksetDims("WS"); else AllworksetsDIMS.Add(new List<Dictionary<string, double>>());
            if (AllworksetsNames[2].Count > 0) getWorksetDims("CHW"); else AllworksetsDIMS.Add(new List<Dictionary<string, double>>());
            if (AllworksetsNames[3].Count > 0) getWorksetDims("DR"); else AllworksetsDIMS.Add(new List<Dictionary<string, double>>());
            if (AllworksetsNames[4].Count > 0) getWorksetDims("FF"); else AllworksetsDIMS.Add(new List<Dictionary<string, double>>());
            if (AllworksetsNames[5].Count > 0) getWorksetDims("CT"); else AllworksetsDIMS.Add(new List<Dictionary<string, double>>());
            foreach (Control control in panel1.Controls)
            {
                if (control is UserControl && c.Name.Contains("System"))
                {
                    DataGridView d = (DataGridView)control.Controls.Find("dgv", true).FirstOrDefault();
                    DIMS = new List<Dictionary<string, double>>();
                    if (d.Enabled == true)
                    {
                        for (int j = 0; j < d.RowCount; j++)
                        {
                            Dictionary<string, double> dict = new Dictionary<string, double>();
                            dict["size"] = Convert.ToDouble(d[0, j].Value.ToString());
                            dict["spacing"] = Convert.ToDouble(d[1, j].Value.ToString());
                            dict["family"] = ((ComboBox)control.Controls.Find("hangerFamily", true).FirstOrDefault()).SelectedIndex;
                            DIMS.Add(dict);
                        }
                        AllworksetsDIMS.Add(DIMS);
                    }
                    else
                    {
                        TextBox t = (TextBox)c.Controls.Find("allSizesSpacing", true).FirstOrDefault();
                        Dictionary<string, double> dict = new Dictionary<string, double>();
                        dict["spacing"] = Convert.ToDouble(t.Text);
                        dict["family"] = ((ComboBox)c.Controls.Find("hangerFamily", true).FirstOrDefault()).SelectedIndex;
                        DIMS.Add(dict);
                        AllworksetsDIMS.Add(DIMS);
                    }
                }
            }
        }
        private void Ductbt_Click(object sender, EventArgs e)
        {
            if (isValidDims())
            {
                showUserControl("Duct");
                buttonClicked(ductButton.Name);
            }
            else
            {
                MessageBox.Show("Size or Spacing Value must be a number and greater than 0!");
            }
        }

        private void WSbt_Click(object sender, EventArgs e)
        {
            if (isValidDims())
            {
                showUserControl("WS");
                buttonClicked(WSButton.Name);
            }
            else
            {
                MessageBox.Show("Size or Spacing Value must be a number and greater than 0!");
            }
        }

        private void CHWbt_Click(object sender, EventArgs e)
        {
            if (isValidDims())
            {
                showUserControl("CHW");
                buttonClicked(CHWButton.Name);
            }
            else
            {
                MessageBox.Show("Size or Spacing Value must be a number and greater than 0!");
            }
        }


        private void DRbt_Click(object sender, EventArgs e)
        {
            if (isValidDims())
            {
                showUserControl("DR");
                buttonClicked(DRButton.Name);
            }
            else
            {
                MessageBox.Show("Size or Spacing Value must be a number and greater than 0!");
            }
        }

        private void FFbt_Click(object sender, EventArgs e)
        {
            if (isValidDims())
            {
                showUserControl("FF");
                buttonClicked(FFButton.Name);
            }
            else
            {
                MessageBox.Show("Size or Spacing Value must be a number and greater than 0!");
            }
        }

        private void addSystem(object sender, EventArgs e)
        {
            button3.DialogResult = DialogResult.None;
            Button systemButton = new Button();
            systemButton.Click += genButtonClicked;
            panel3.Controls.Add(systemButton);
            systemButton.Location = new Point(lastButton.Location.X, lastButton.Location.Y + 50);
            systemButton.Size = lastButton.Size;
            systemButton.BackColor = lastButton.BackColor;
            systemButton.Font = lastButton.Font;
            systemButton.ForeColor = lastButton.ForeColor; systemButton.FlatStyle = FlatStyle.Flat;
            systemButton.FlatAppearance.BorderSize = 0;
            string systemName = "System " + (panel3.Controls.Count - 6).ToString();
            systemButton.Text = systemName;
            systemButton.Name = systemName;
            UserControl newUserControl = createUserControl(systemName);
            lastButton = systemButton;
        }

        private void genButtonClicked(object sender, EventArgs e)
        {
            if (isValidDims())
            {
                showUserControl(((Button)sender).Text);
                buttonClicked(((Button)sender).Text);
            }
            else
            {
                MessageBox.Show("Size or Spacing Value must be a number and greater than 0!");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            linkIndex = comboBox1.SelectedIndex;
        }

        private void fixedSpacing_CheckedChanged(object sender, EventArgs e)
        {
            DataGridView dgv = getActiveUC().Controls.Find("dgv", true).FirstOrDefault() as DataGridView;
            TextBox textBox = getActiveUC().Controls.Find("allSizesSpacing", true).FirstOrDefault() as TextBox;
            Label label = getActiveUC().Controls.Find("fixedSpacingLabel", true).FirstOrDefault() as Label;
            if (checkBox1.Checked)
            {
                dgv.Visible = false;
                dgv.Enabled = false;
                textBox.Visible = true;
                label.Visible = true;
            }
            else
            {
                dgv.Visible = true;
                dgv.Enabled = true;
                textBox.Visible = false;
                label.Visible = false;

            }
        }


        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked) { selc = true; panel7.Enabled = false; }
            else { selc = false; panel7.Enabled = true; }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked) { selc = false; panel7.Enabled = true; }
            else { selc = true; panel7.Enabled = false; }
        }
    }
}
