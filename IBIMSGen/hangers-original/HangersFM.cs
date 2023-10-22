using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IBIMS_MEP
{
    public partial class HangersFM : Form
    {

        public bool canc; public bool ook; public int lnk =-1; public int dup;  public bool selc; public int frin,toin;
        public List<string> worksetnames = new List<string>(); public List<string> Linkes = new List<string>();
        public List<List<string>> AllworksetsNames = new List<List<string>>(); public List<string> Levels = new List<string>();
        public List<List<List<string>>> AllworksetsDIMS = new List<List<List<string>>>();
        public UserControl Duc = null;
        public UserControl Wuc = null;
        public UserControl CWuc = null;
        public UserControl DRuc = null;
        public UserControl FFuc = null;
        public Button LastB = null;
        //public List<string> DRdias = new List<string>() { "12","15", "20", "25", "28", "32", "40", "50", "75", "110", "125", "160", "200", "250", "315" };
        //public List<string> WSdias = new List<string>() { "15", "20", "25", "32", "40", "50","65", "80","90", "100","125", "150","200","250","300","350" };
        //public List<string> CHWdias = new List<string>() { "15", "20", "25", "32", "40", "50", "65", "80", "90", "100", "125", "150", "200", "250", "300", "350" };
        public List<string> DRdias = new List<string>() { "20", "25" , "32", "40", "50", "75", "110", "125","130" };
        public List<string> WSdias = new List<string>() { "20", "25" , "32", "40", "50", "75", "110", "125" ,"130"};
        public List<string> CHWdias = new List<string>() { "15", "20", "25", "32", "40", "50", "65", "80", "100", "125", "150", "200", "250", "300", "350", "400" };
        public List<string> Firedias = new List<string>() { "15", "20", "25", "32", "40", "50", "65", "80", "90", "100", "125", "150", "200", "250", "300", "350" };
        //public List<double> DRspcs = new List<double>() { 530, 610, 685, 720, 760, 840, 915, 1065, 1370, 1525, 1680, 1830,1830,1830,1830 };
        //public List<double> WSspcs = new List<double>() { 1800, 2400, 2400, 2700, 3000, 3000, 3300, 3600, 3700, 3900, 4200, 4500, 4500, 4500, 4500, 4500 };
        //public List<double> CHWspcs = new List<double>() { 1800, 2400, 2400, 2700, 3000, 3000, 3300, 3600, 3700, 3900, 4200, 4500, 4500, 4500, 4500, 4500 };
        public List<double> DRspcs = new List<double>() {1000,1000,1000,1000,1065,1370,1525,1680,1830 };
        public List<double> WSspcs = new List<double>() {1000,1000,1000,1000,1065,1370,1525,1680,1830 };
        public List<double> CHWspcs = new List<double>() { 2100,2100,2100,2400,2700,3100,3400,3600,4300,4300,5100,5800,6100,7000,7600,8200};
        public List<double> Firespcs = new List<double>() { 1800, 2400, 2400, 2700, 3000, 3000, 3300, 3600, 3700, 3900, 4200,4500,4500,4500,4500,4500 };
        public void textlev(TextBox txtbx, double defu, double min, double max)
        {
            try
            {
                if (Convert.ToDouble(txtbx.Text) < min || Convert.ToDouble(txtbx.Text) > max)
                {
                    txtbx.Text = defu.ToString();
                }
            }
            catch (FormatException)
            {
                txtbx.Text = defu.ToString();
            }
        }
        DataGridView DGVCreator(List<string> sizes = null , List<double> spaces = null,bool duct=false)
        {
            DataGridView d = new DataGridView();
            d.DefaultCellStyle.NullValue = "0";
            d.CellLeave += D_CellLeave;
            if (!duct)
            {
                if (sizes == null) { sizes = new List<string>(); }
                if (spaces == null) { spaces = new List<double>(); }
                d.Columns.Add("Size", "Size ( mm )");
                d.Columns.Add("Spacing", "Spacing ( mm )");
                d.Columns[0].Width = 150; d.Columns[1].Width = 146;
                if (sizes.Count == 0) { d.RowCount = 12; }
                else { d.RowCount = sizes.Count; }
                int o = 0;
                foreach (string s in sizes)
                {
                    d[0, o].Value = s;
                    d[1, o].Value = spaces[o];
                    o++;
                }
            }
            else
            {
                sizes = new List<string>() { "1","533.4","838.2", "1041.4","1524" };
                List<string> sizes2 = new List<string>() { "508", "812.8", "1016", "1524", "10000" };
                spaces = new List<double>() { 3000, 2500, 2500, 2000, 1500 };
                d.Columns.Add("Size1", "Size from (mm)");
                d.Columns.Add("Size2", "Size to (mm)");
                d.Columns.Add("Spacing", "Spacing (mm)");
                d.Columns[0].Width = 98; d.Columns[1].Width = 98; d.Columns[2].Width = 100;
                d.RowCount = sizes.Count;
                int o = 0;
                foreach (string s in sizes)
                {
                    d[0, o].Value = s;
                    d[1, o].Value = sizes2[o];
                    d[2, o].Value = spaces[o];
                    o++;
                }
            }
            return d;
        }

        private void D_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dd = (DataGridView)sender;
            foreach (DataGridViewRow row in dd.Rows)
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
        UserControl ActiveUS()
        {
            UserControl uss = null;
            foreach (UserControl us in panel1.Controls)
            {
                if (us.Visible == true)
                {
                    uss = us;
                    break;
                }
            }
            return uss;
        }
        public UserControl US(string name, bool vis = false,List<string> sizes = null, List<double> spaces = null)
        {
            UserControl us = new UserControl();
            us.Name = name;
            panel1.Controls.Add(us);
            us.Dock = DockStyle.Fill;
            Panel pl = new Panel(); Panel pc = new Panel(); Panel pr = new Panel(); Panel pb = new Panel();
            us.Controls.Add(pl);  us.Controls.Add(pc); us.Controls.Add(pr); us.Controls.Add(pb); pb.BackColor = Color.Snow; pb.BorderStyle = BorderStyle.FixedSingle;
            pb.Dock = DockStyle.Bottom; pb.Height = 50; pl.Dock = DockStyle.Left; pl.Width = 300;
            pr.Dock = DockStyle.Right; pr.Width = 300; 
            //pc.Dock = DockStyle.Fill;
             
            CheckedListBox clb = new CheckedListBox();
            pl.Controls.Add(clb);
            foreach (string ss in worksetnames)
            {
                clb.Items.Add(ss);
            }
            clb.Location = new Point(211, 38);
            clb.Dock = DockStyle.Fill; clb.BorderStyle = BorderStyle.None;
            clb.Font = new Font("Microsoft JhengHei", 13,FontStyle.Regular,GraphicsUnit.Pixel);
            clb.CheckOnClick = true;
            ComboBox cbx = new ComboBox(); Label lab = new Label();
            pb.Controls.Add(cbx); pb.Controls.Add(lab); lab.Location = new Point(210, 20); 
            lab.Size = new Size(180, 13); lab.Text = "Duplicate table from: ";
            lab.Font = new Font("Arial", 14, FontStyle.Regular, GraphicsUnit.Pixel);
            cbx.Location = new Point(390, 15); cbx.Size = new Size(150, 15);
            cbx.Enabled = false;    
            cbx.SelectedIndexChanged += Cbx_SelectedIndexChanged;

            DataGridView dgv = null;
            if (name == "Duc")
            {
                dgv = DGVCreator(sizes, spaces, true);
            }
            else
            {
                dgv = DGVCreator(sizes, spaces);
            }
            pr.Controls.Add(dgv);
            dgv.Dock = DockStyle.Fill;
            dgv.Location = new Point(600, 30); dgv.RowHeadersVisible = false;
            TextBox tb = new TextBox(); pr.Controls.Add(tb); 
            tb.Location = new Point(120, 100); tb.Visible = false; tb.Text = "2000"; tb.Leave += Tb_Leave;
            Label lbb = new Label(); pr.Controls.Add(lbb); lbb.Size = new Size(180, 18); lbb.Text = "Spacing in ( mm ): "; 
            lbb.Font = new Font("Arial", 14, FontStyle.Regular, GraphicsUnit.Pixel); lbb.Visible = false; lbb.Location = new Point(115, 60);
            if (vis) { us.Visible = true; }
            else { us.Visible = false; }
            return us;
        }

        private void Tb_Leave(object sender, EventArgs e)
        {
            textlev((TextBox)sender, 2000, 500, 10000);
        }


        
        private void Cbx_SelectedIndexChanged(object sender, EventArgs e)
        {
            dup = ((ComboBox)sender).SelectedIndex;
            DataGridView neww = ActiveUS().Controls[2].Controls[0] as DataGridView;
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
                        old = c.Controls[2].Controls[0] as DataGridView;
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

        bool Go()
        {
            bool go = true;
            DataGridView activeD = null;
            activeD = (DataGridView)ActiveUS().Controls[2].Controls[0];
            foreach (DataGridViewRow row in activeD.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    try
                    {
                        if (cell.Value != null)
                        {
                            if (Convert.ToDouble(cell.Value) < 0)
                            {
                                cell.ErrorText = "Value Must be Integar and greater than 0.";
                            }
                        }
                    }
                    catch (FormatException)
                    {
                        cell.ErrorText = "Value Must be Integar and greater than 0.";
                    }
                }
            }
            foreach (DataGridViewRow row in activeD.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.ErrorText != "")
                    {
                        go = false;
                        break;
                    }
                }
            }
            return go;
        }
        
        void ListAdd(string name)
        {
            foreach (Control c in panel1.Controls)
            {
                if (c is UserControl && c.Name == name)
                {
                    CheckedListBox clb = c.Controls[0].Controls[0] as CheckedListBox;
                    List<string> worksetcats = new List<string>();
                    foreach (int ii in clb.CheckedIndices)
                    {
                        worksetcats.Add(worksetnames[ii]);
                    }
                    AllworksetsNames.Add(worksetcats); 
                    break;
                }
            }
        }
        void DGVAdd( string name)
        {
            foreach (Control c in panel1.Controls)
            {
                if (c is UserControl && c.Name == name)
                {
                    DataGridView d = (DataGridView)c.Controls[2].Controls[0];
                    List<List<string>> DIMS = new List<List<string>>();
                    if (d.Enabled == true)
                    {
                        for (int i = 0; i < 2; i++)
                        {
                            List<string> list = new List<string>();
                            for (int j = 0; j < d.RowCount; j++)
                            {
                                if (d[i, j].Value != null)
                                {
                                    list.Add(d[i, j].Value.ToString());
                                }
                                else
                                {
                                    list.Add("0");
                                }
                            }
                            DIMS.Add(list);
                        }
                        AllworksetsDIMS.Add(DIMS);
                        break;
                    }
                    else
                    {
                        TextBox t = (TextBox)c.Controls[2].Controls[1];
                        List<string> list1 = new List<string>();
                        List<string> list2 = new List<string>();
                        list1.Add("-10"); DIMS.Add(list1);
                        list2.Add(t.Text); DIMS.Add(list2);
                        AllworksetsDIMS.Add(DIMS);
                        break;
                    }
                }
            }
        }

        void USvis(string name)
        {
            UserControl uc = null;
            foreach (Control c in panel1.Controls)
            {
                if (c is UserControl)
                {
                    if (c.Name == name)
                    {
                        c.Visible = true; uc = (UserControl)c; 
                    }
                    else
                    {
                        c.Visible = false;
                    }
                }
            }
            DataGridView dd = uc.Controls[2].Controls[0] as DataGridView;
            if (dd.Enabled)
            {
                checkBox1.Checked = false;
            }
            else
            {
                checkBox1.Checked = true;
            }
        }
        void btnclr(string btname)
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
                        c.BackColor = Color.Snow;
                    }
                }
            }
            ComboBox cbx = ActiveUS().Controls[3].Controls[0] as ComboBox;
            if (ActiveUS().Name.Contains("System"))
            {
                cbx.Enabled = true;
                cbx.Items.Clear();
                cbx.Items.Add("New Table");
                foreach (Control c in panel3.Controls )
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
        public HangersFM()
        {
            InitializeComponent();
        }



        private void button1_Click(object sender, EventArgs e)
        {
            AllworksetsNames = new List<List<string>>();
            AllworksetsDIMS = new List<List<List<string>>>();
            if (lnk == -1)
            {
                ook = false;
                button1.DialogResult = DialogResult.None;
                MessageBox.Show("Select a Revit Link before Run.");
            }
             else if (!Go())
            {
                ook = false;
                button1.DialogResult = DialogResult.None;
                MessageBox.Show("Value Must be Integar and greater than 0.");
            }
            else 
            {
                ListAdd("Duc");
                ListAdd("WS");
                ListAdd("CHW");
                ListAdd("DR");
                ListAdd("FF");
                foreach (Control c in panel1.Controls)
                {
                    if (c is UserControl && c.Name.Contains("System"))
                    {
                        CheckedListBox clb = c.Controls[0].Controls[0] as CheckedListBox;
                        List<string> list = new List<string>();
                        foreach (int ii in clb.CheckedIndices)
                        {
                            list.Add(worksetnames[ii]);
                        }
                        AllworksetsNames.Add(list);
                        break;
                    }
                }
                foreach (Control c in panel1.Controls)
                {
                    if (c is UserControl && c.Name == "Duc")
                    {
                        DataGridView d = (DataGridView)c.Controls[2].Controls[0];
                        List<List<string>> DIMS = new List<List<string>>();
                        if (d.Enabled == true)
                        {
                            for (int i = 0; i < 3; i++)
                            {
                                List<string> list = new List<string>();
                                for (int j = 0; j < d.RowCount; j++)
                                {
                                    list.Add(d[i, j].Value.ToString());
                                }
                                DIMS.Add(list);
                            }
                            AllworksetsDIMS.Add(DIMS);
                            break;
                        }
                        else
                        {
                            TextBox t = (TextBox)c.Controls[2].Controls[1];
                            List<string> list1 = new List<string>();
                            List<string> list2 = new List<string>();
                            list1.Add("-10"); DIMS.Add(list1);
                            list2.Add(t.Text); DIMS.Add(list2);
                            AllworksetsDIMS.Add(DIMS);
                            break;
                        }
                    }
                }
                DGVAdd("WS");
                DGVAdd("CHW");
                DGVAdd("DR");
                DGVAdd("FF");
                foreach (Control c in panel1.Controls)
                {
                    if (c is UserControl && c.Name.Contains("System"))
                    {
                        DataGridView d=(DataGridView)c.Controls[2].Controls[0];
                        List<List<string>> DIMS = new List<List<string>>();
                        if (d.Enabled == true)
                        {
                            for (int i = 0; i < 2; i++)
                            {
                                List<string> list = new List<string>();
                                for (int j = 0; j < d.RowCount; j++)
                                {
                                    if (d[i, j].Value != null)
                                    {
                                        list.Add(d[i, j].Value.ToString());
                                    }
                                    else
                                    {
                                        list.Add("0");
                                    }
                                }
                                DIMS.Add(list);
                            }
                            AllworksetsDIMS.Add(DIMS);
                            break;
                        }
                        else
                        {
                            TextBox t = (TextBox)c.Controls[2].Controls[1];
                            List<string> list1 = new List<string>();
                            List<string> list2 = new List<string>();
                            list1.Add("-10"); DIMS.Add(list1);
                            list2.Add(t.Text); DIMS.Add(list2);
                            AllworksetsDIMS.Add(DIMS);
                            break;
                        }
                    }
                }
                ook = true;
                button1.DialogResult = DialogResult.OK;
                this.Close();
            }
            frin = comboBox2.SelectedIndex; toin = comboBox3.SelectedIndex;
        }
            
        private void button2_Click(object sender, EventArgs e)
        {
            canc = true;
        }
        private void Form7_Load(object sender, EventArgs e)
        {
            foreach(string ss in Linkes) { comboBox1.Items.Add(ss); }
            foreach (string ss in Levels) { comboBox2.Items.Add(ss); comboBox3.Items.Add(ss); }
            Duc = US("Duc",true);
            Wuc = US("WS",false,WSdias,WSspcs);
            CWuc = US("CHW",false,CHWdias,CHWspcs);
            DRuc = US("DR",false,DRdias,DRspcs);
            FFuc = US("FF",false,Firedias,Firespcs);
            LastB = FFbt;
            comboBox2.SelectedItem = Levels[0]; comboBox2.SelectedIndex= 0;
            comboBox3.SelectedItem = Levels[1]; comboBox3.SelectedIndex =1;
        }
        private void Ductbt_Click(object sender, EventArgs e)
        {
            if (Go())
            {
                USvis("Duc"); btnclr(Ductbt.Name);
            }
            else
            {
                MessageBox.Show("Value Must be Integar and greater than 0.");
            }
        }

        private void WSbt_Click(object sender, EventArgs e)
        {
            if (Go())
            {
                USvis("WS"); btnclr(WSbt.Name);
            }
            else
            {
                MessageBox.Show("Value Must be Integar and greater than 0.");
            }
        }

        private void CHWbt_Click(object sender, EventArgs e)
        {
            if (Go())
            {
                USvis("CHW"); btnclr(CHWbt.Name);
            }
            else
            {
                MessageBox.Show("Value Must be Integar and greater than 0.");
            }
        } 


        private void DRbt_Click(object sender, EventArgs e)
        {
            if (Go())
            {
                USvis("DR"); btnclr(DRbt.Name);
            }
            else
            {
                MessageBox.Show("Value Must be Integar and greater than 0.");
            }
        }

        private void FFbt_Click(object sender, EventArgs e)
        {
            if (Go())
            {
                USvis("FF"); btnclr(FFbt.Name);
            }
            else
            {
                MessageBox.Show("Value Must be Integar and greater than 0.");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            button3.DialogResult = DialogResult.None;
            Button b = new Button(); 
            b.Click += B_Click;
            panel3.Controls.Add(b); 
            b.Location = new Point(LastB.Location.X, LastB.Location.Y+50);
            b.Size= LastB.Size; 
            b.BackColor = LastB.BackColor; 
            b.Font = LastB.Font; b.ForeColor = LastB.ForeColor; b.FlatStyle= FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 0;
            string s = "System " + (panel3.Controls.Count - 5).ToString();
            b.Text = s; b.Name = s;
            UserControl uc6 = US(s);
            LastB = b;
        }

        private void B_Click(object sender, EventArgs e)
        {
            if (Go())
            {
                USvis(((Button)sender).Text); btnclr(((Button)sender).Text);
            }
            else
            {
                MessageBox.Show("Value Must be Integar and greater than 0.");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            lnk=comboBox1.SelectedIndex;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            DataGridView d = ActiveUS().Controls[2].Controls[0] as DataGridView;
            TextBox tb = ActiveUS().Controls[2].Controls[1] as TextBox;
            Label lbb = ActiveUS().Controls[2].Controls[2] as Label;
            if (checkBox1.Checked)
            {
                d.Visible = false;
                d.Enabled = false;
                tb.Visible = true;
                lbb.Visible = true;
            }
            else
            {
                d.Visible = true;
                d.Enabled = true;
                tb.Visible = false;
                lbb.Visible = false;

            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton1.Checked) { selc = true; panel7.Enabled = false; }
            else { selc = false; panel7.Enabled = true; }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked) { selc = false; panel7.Enabled = true; }
            else { selc = true; panel7.Enabled = false; }
        }
    }
}
