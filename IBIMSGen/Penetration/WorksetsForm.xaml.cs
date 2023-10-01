using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ComboBox = System.Windows.Controls.ComboBox;

namespace IBIMSGen.Penetration
{
    /// <summary>
    /// Interaction logic for WorksetsForm.xaml
    /// </summary>
    public partial class WorksetsForm : Window
    {
        IList<string> worksetNames;
        IList<FamilySymbol> familySymbols;
        ICollection<WrapPanel> Panels;
        public WorksetsForm(IList<string> worksetNames, IList<FamilySymbol> familySymbols)
        {
            InitializeComponent();
            this.worksetNames = worksetNames;
            this.familySymbols = familySymbols;
            Panels = new List<WrapPanel>();
        }

        private void cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void addMore_Click(object sender, RoutedEventArgs e)
        {

        }

        private void selectionChange(object sender, SelectionChangedEventArgs e)
        {
            ComboBox c = Panels.Where(x => x.Children[3] == sender).FirstOrDefault().Children[5] as ComboBox;
            foreach(var a in this.familySymbols.Where(x => x.Family.Name == c.SelectedItem.ToString()).Select(x => x.Name))
            {

            c.Items.Add(a);
            }
        }

        private void addMore_Click_1(object sender, RoutedEventArgs e)
        {
            Label l1 = new Label();
            l1.Content = "Worksets";
            Label l2 = new Label();
            l1.Content = "FamilyName";
            Label l3 = new Label();
            l1.Content = "Type Name";
            ComboBox wscb = new ComboBox();
            foreach (string a in this.worksetNames)
            {
                wscb.Items.Add(a);
            }
            wscb.Width = 200;
            wscb.Height = 40;
            ComboBox fname = new ComboBox();
            foreach (var a in this.familySymbols.Select(x => x.Family.Name).Distinct().ToList())
            {

                fname.Items.Add(a);
            }
            fname.Width = 200;
            fname.Height = 40;
            ComboBox tyName = new ComboBox();
            tyName.Width = 200;
            tyName.Height = 40;

            WrapPanel wrapPanel = new WrapPanel();
            wrapPanel.VerticalAlignment = VerticalAlignment.Top;
            wrapPanel.Height = 50;
            wrapPanel.Margin = new Thickness(10, GRID.Children.Count * 110, 10, 10);
            wrapPanel.Children.Add(l1);
            wrapPanel.Children.Add(wscb);
            wrapPanel.Children.Add(l2);
            wrapPanel.Children.Add(fname);
            wrapPanel.Children.Add(l3);
            wrapPanel.Children.Add(tyName);
            GRID.Children.Add(wrapPanel);
            Panels.Add(wrapPanel);
            foreach (WrapPanel panel in Panels)
            {
                ComboBox c = panel.Children[3] as ComboBox;
                c.SelectionChanged += selectionChange;
            }
        }
    }
}
