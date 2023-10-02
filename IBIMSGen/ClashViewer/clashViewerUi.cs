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
    public partial class clashViewerUi : System.Windows.Forms.Form
    {
        UIDocument uidoc;
        Document doc;
        IList<ClashPoint> clashes;
        Element element;
        public ElementId elementId;
        ClashPoint clash = null;
        double offset;
        public clashViewerUi(UIDocument uidoc, IList<ClashPoint> clashes)
        {
            this.uidoc = uidoc;
            this.doc = uidoc.Document;
            this.clashes = clashes;
            InitializeComponent();
            List<string> ids = this.clashes.Select(x => "Test#" + x.testName.Split(' ')[0] + " : " + x.clashName + "    /   ID: " + x.elementId.ToString()).Distinct().ToList();
            comboBox1.Items.AddRange(ids.ToArray());
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                this.clash = clashes.ElementAt(comboBox1.SelectedIndex);
                this.element = this.doc.GetElement(this.clash.elementId);
                textBox1.Text = element.Name;
                textBox2.Text = this.clash.testName;
                if (this.clash.elementId.ToString() == this.clash.elementId1)
                {
                    textBox4.Text = this.clash.elementId2;
                }
                else
                {
                    textBox4.Text = this.clash.elementId1;
                }
                textBox5.Text = this.clash.left;
                textBox6.Text = this.clash.right;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox7.Text.Trim().Length > 0)
            {
                if (double.TryParse(textBox7.Text, out double value))
                {
                    this.offset = value / 304.8;
                }
                else
                {
                    TaskDialog.Show("Error", "Offset value must be a real number");
                    return;
                }
            }
            else { this.offset = 0; }
            ElementId elementId = this.element.Id;
            using (Transaction t = new Transaction(doc))
            {
                t.Start("3D Views");
                View3D view3D = doc.ActiveView as View3D;
                if (view3D != null)
                {
                    BoundingBoxXYZ bx2 = this.element.get_BoundingBox(view3D);
                    XYZ min = bx2.Min;
                    XYZ max = bx2.Max;
                    BoundingBoxXYZ bx = new BoundingBoxXYZ();
                    bx.Min = new XYZ(min.X - offset, min.Y - offset, min.Z - offset);
                    bx.Max = new XYZ(max.X + offset, max.Y + offset, max.Z + offset);
                    bx.Transform = Transform.Identity;
                    view3D.SetSectionBox(bx);
                    doc.Regenerate();
                    uidoc.ShowElements(this.element);
                }
                uidoc.Selection.SetElementIds(new List<ElementId> { elementId });
                t.Commit();
                t.Dispose();
            }
            //this.DialogResult = DialogResult.OK;

            //this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox4.Text != null)
            {

                System.Windows.Clipboard.SetText(textBox4.Text);
            }
        }
    }

}
