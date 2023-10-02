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

namespace IBIMSGen.ClashViewer
{
    public partial class clashBalls : System.Windows.Forms.Form
    {
        public IList<Element> balls;
        public BindingList<ClashPoint> clashes = new BindingList<ClashPoint>();
        UIDocument uidoc;
        Document doc;
        ProjectPosition position;
        clashPoints parent;
        UIApplication uiapp;
        public IList<ClashPoint> points;

        public clashBalls(clashPoints parent, UIApplication uiapp, IList<Element> balls, bool vis)
        {
            points = new List<ClashPoint>();
            this.uiapp = uiapp;
            this.uidoc = this.uiapp.ActiveUIDocument;
            this.doc = this.uidoc.Document;
            this.position = this.doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero);
            this.balls = balls;
            this.parent = parent;
            InitializeComponent();
            button3.Visible = vis;
            getClashes();
            DGV();
        }

        public clashBalls(clashPoints parent, UIApplication uiapp, IList<ClashPoint> clashes, bool vis)
        {
            points = new List<ClashPoint>();
            this.uiapp = uiapp;
            this.uidoc = this.uiapp.ActiveUIDocument;
            this.doc = this.uidoc.Document;
            this.position = this.doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero);
            this.parent = parent;
            this.clashes.AllowNew = true;
            this.clashes.AllowEdit = true;
            foreach (ClashPoint clash in clashes)
            {
                this.clashes.Add(clash);
            }
            //this.clashes = clashes;
            InitializeComponent();
            button3.Visible = vis;
            DGV();
        }

        private void getClashes()
        {
            foreach (Element ball in this.balls)
            {
                string testName = ball.LookupParameter("Test Name").AsString();
                string clashName = ball.LookupParameter("Clash Name").AsString();
                string elementId1 = ball.LookupParameter("Element ID (1)").AsString();
                string elementId2 = ball.LookupParameter("Element ID (2)").AsString();
                string testType = ball.LookupParameter("Test Type").AsString();
                string tolerance = ball.LookupParameter("Tolerance").AsString();
                string right = ball.LookupParameter("Group A").AsString();
                string left = ball.LookupParameter("Group B").AsString();
                double x = ball.LookupParameter("X").AsDouble();
                double y = ball.LookupParameter("Y").AsDouble();
                double z = ball.LookupParameter("Z").AsDouble();
                string comment = ball.LookupParameter("Clash Comment").AsString();
                XYZ clashPoint = new XYZ(x, y, z);
                int status = ball.LookupParameter("Status (Resolved/Not)").AsInteger();
                bool resolved = true;
                if (status > 0)
                {
                    resolved = true;
                }
                else
                {
                    resolved = false;
                }
                this.clashes.AllowNew = true;
                this.clashes.AllowEdit = true;
                this.clashes.Add(new ClashPoint(testName, testType, tolerance, left, right, clashName, clashPoint, this.position, elementId1, elementId2, resolved, comment));
            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell cell = this.dataGridView1.SelectedCells[0];
            if (cell == null) { return; }
            int index = cell.RowIndex;
            try
            {
                using (Transaction tx = new Transaction(this.doc))
                {
                    tx.Start("zoom in");
                    Autodesk.Revit.DB.View3D activeView = this.doc.ActiveView as View3D;
                    Element elem;
                    if (this.balls != null)
                    {

                        elem = this.balls.ElementAt(index);
                    }
                    else
                    {
                        elem = this.doc.GetElement(this.clashes.ElementAt(index).elementId);
                    }
                    BoundingBoxXYZ bx = elem.get_BoundingBox(activeView);
                    BoundingBoxXYZ bxx = new BoundingBoxXYZ();
                    bxx.Min = bx.Min;
                    bxx.Max = bx.Max;
                    bxx.Transform = Transform.Identity;
                    activeView.SetSectionBox(bxx);
                    this.doc.Regenerate();
                    this.uidoc.ShowElements(elem);
                    tx.Commit();
                    tx.Dispose();
                }
            }
            catch { }

        }

        private void DGV()
        {
            this.dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;
            this.dataGridView1.DataSource = this.clashes;
            try { this.dataGridView1.Columns["testname"].HeaderText = "Test"; } catch { }
            try { this.dataGridView1.Columns["clashName"].HeaderText = "Clash Name"; } catch { }
            try { this.dataGridView1.Columns["testType"].HeaderText = "Type"; } catch { }
            try { this.dataGridView1.Columns["elementId1"].HeaderText = "ID 1"; } catch { }
            try { this.dataGridView1.Columns["elementId2"].HeaderText = "ID 2"; } catch { }
            try { this.dataGridView1.Columns["left"].HeaderText = "Left"; } catch { }
            try { this.dataGridView1.Columns["right"].HeaderText = "Right"; } catch { }
            try { this.dataGridView1.Columns["tolerance"].HeaderText = "Tolerance"; } catch { }
            try { this.dataGridView1.Columns["resolved"].HeaderText = "Resolved"; } catch { }
            try { this.dataGridView1.Columns["comment"].HeaderText = "Comment"; } catch { }
            try { this.dataGridView1.Columns["x"].HeaderText = "X"; } catch { }
            try { this.dataGridView1.Columns["y"].HeaderText = "Y"; } catch { }
            try { this.dataGridView1.Columns["z"].HeaderText = "Z"; } catch { }


        }


        private void OK_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {

                try
                {
                    if (row.Cells["testName"].Value == null)
                    {
                        continue;
                    }
                    points.Add(
                        new ClashPoint(
                        row.Cells["testName"].Value.ToString(), row.Cells["testType"].Value.ToString(), row.Cells["tolerance"].Value.ToString(),
                        row.Cells["left"].Value.ToString(), row.Cells["right"].Value.ToString(), row.Cells["clashName"].Value.ToString(),
                        this.clashes.ElementAt(dataGridView1.Rows.IndexOf(row)).clashPoint, this.position,
                        row.Cells["elementId1"].Value.ToString(), row.Cells["elementId2"].Value.ToString(), Convert.ToBoolean(row.Cells["resolved"].Value), row.Cells["comment"].Value.ToString()
                        )
                        );
                }
                catch (Exception ex)
                {
                    //TaskDialog.Show("error", ex.Message);
                }
            }

            this.Close();
            this.DialogResult = DialogResult.OK;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
            this.parent.readAgain(false, false);
            this.DialogResult = DialogResult.OK;
        }

    }
}
