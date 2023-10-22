using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IBIMSGen.ColorCable
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class Coloring : IExternalCommand
    {

        UIApplication app;
        UIDocument uidoc;
        Document doc;
        coloringUi form;
        List<int> R = new List<int>();
        List<int> Y = new List<int>();
        List<int> B = new List<int>();
        FilteredElementCollector textNotes;
        string place;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            app = commandData.Application;
            uidoc = app.ActiveUIDocument;
            doc = uidoc.Document;
            form = new coloringUi();
            form.ShowDialog();

            if (form.DialogResult == System.Windows.Forms.DialogResult.Cancel)
            {
                return Result.Cancelled;
            }
            if (form.cb1.Checked)
            {

                textNotes = new FilteredElementCollector(doc, doc.ActiveView.Id);
                place = "in this view only.";
            }
            else
            {
                textNotes = new FilteredElementCollector(doc);
                place = "in the whole project.";
            }
            textNotes.OfClass(typeof(TextNote)).WhereElementIsNotElementType().ToElements();
            string redValues = form.rtb.Text;
            string blueValues = form.btb.Text;
            string yellowValues = form.ytb.Text;

            char c = form.stb.Text.Trim()[0];
            foreach (string a in redValues.Split(','))
            {
                if (int.TryParse(a, out _))
                {
                    R.Add(int.Parse(a));
                }
            }
            foreach (string a in blueValues.Split(','))
            {
                if (int.TryParse(a, out _))
                {
                    B.Add(int.Parse(a));
                }
            }
            foreach (string a in yellowValues.Split(','))
            {
                if (int.TryParse(a, out _))
                {
                    Y.Add(int.Parse(a));
                }
            }
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("replace text");
                //uidoc.Selection.SetElementIds(textNotes.Select(x => x.Id).ToList());
                foreach (Element note in textNotes)
                {
                    if (note is TextNote)
                    {
                        TextNote textNote = note as TextNote;
                        string num = textNote.Text.Split(c)[0].Split(' ').LastOrDefault();
                        if (int.TryParse(num, out _))
                        {
                            string sent = textNote.Text.Split(c)[0];
                            string ts = textNote.Text;
                            IList<string> splited = textNote.Text.Split(c);
                            if (ts.Split(c).Count() > 1)
                            {
                                if (R.Contains(int.Parse(num))) { sent += "R" + c; }
                                else if (Y.Contains(int.Parse(num))) { sent += "Y" + c; }
                                else if (B.Contains(int.Parse(num))) { sent += "B" + c; }
                                for (int i = 1; i < splited.Count(); i++)
                                {
                                    sent += splited[i] + " ";
                                }
                                textNote.Text = sent;
                            }
                        }
                        else
                        {
                            //textNote.Text = textNote.Text.Replace(t, t + "Omar");
                            //break;
                        }

                        //textNote.Text.Replace();
                    }
                }
                tx.Commit();
                tx.Dispose();
            }
            TaskDialog.Show("Done", $"All text have been replaced {place}");
            return Result.Succeeded;

        }
    }
}
