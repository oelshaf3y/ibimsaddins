﻿using Autodesk.Revit.DB;
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
        IList<Workset> worksets;
        IList<string> worksetNames;
        IList<FamilySymbol> familySymbols;
        public ICollection<WrapPanel> Panels;
        public IList<Tuple<Workset, FamilySymbol>> worksetCollection;
        public bool state;
        public WorksetsForm(IList<Workset> worksets, IList<FamilySymbol> familySymbols)
        {
            InitializeComponent();
            this.worksets = worksets;
            this.worksetNames = this.worksets.Select(x => x.Name).ToList();
            this.familySymbols = familySymbols;
            this.Panels = new List<WrapPanel>();
            this.worksetCollection=new List<Tuple<Workset, FamilySymbol>>();
            this.state = false;
            addMore_Click(null,null);
        }

        
        private void selectionChange(object sender, SelectionChangedEventArgs e)
        {
            string selected = ((ComboBox)sender).SelectedItem.ToString();
            var list = this.familySymbols.Where(x => x.Family.Name == selected).ToList();
            ComboBox tochange = this.Panels.Where(panel => panel.Children[3] == sender as ComboBox).FirstOrDefault().Children[5] as ComboBox;
            tochange.Items.Clear();
            foreach (var item in list)
            {
                tochange.Items.Add(item.Name);
            }
        }

        private void addMore_Click(object sender, RoutedEventArgs e)
        {
            addMore.Margin = new Thickness(addMore.Margin.Left, addMore.Margin.Top + 45, addMore.Margin.Right, addMore.Margin.Bottom);
            OKBut.Margin = new Thickness(OKBut.Margin.Left, OKBut.Margin.Top + 45, OKBut.Margin.Right, OKBut.Margin.Bottom);
            cancelBut.Margin = new Thickness(cancelBut.Margin.Left, cancelBut.Margin.Top + 45, cancelBut.Margin.Right, cancelBut.Margin.Bottom);
            Label l1 = new Label();
            l1.Content = "Worksets";
            Label l2 = new Label();
            l2.Content = "FamilyName";
            Label l3 = new Label();
            l3.Content = "Type Name";
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
            wrapPanel.Height = 40;
            wrapPanel.Margin = new Thickness(10, (GRID.Children.Count - 3) * 45, 10, 10);
            wrapPanel.Children.Add(l1);
            wrapPanel.Children.Add(wscb);
            wrapPanel.Children.Add(l2);
            wrapPanel.Children.Add(fname);
            fname.SelectionChanged += selectionChange;
            wrapPanel.Children.Add(l3);
            wrapPanel.Children.Add(tyName);
            GRID.Children.Add(wrapPanel);
            this.Panels.Add(wrapPanel);

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.state = false;
            this.Close();
        }

        private void OKBut_Click(object sender, RoutedEventArgs e)
        {
            foreach (WrapPanel panel in this.Panels)
            {
                Workset workset =this.worksets.Where(x=> x.Name == (panel.Children[1] as ComboBox).SelectedItem.ToString()).FirstOrDefault();
                FamilySymbol familySymbol = this.familySymbols.Where(x=>
                x.Family.Name == (panel.Children[3] as ComboBox).SelectedItem.ToString() 
                && x.Name== (panel.Children[5] as ComboBox).SelectedItem.ToString()
                ).FirstOrDefault();
                worksetCollection.Add(new Tuple<Workset, FamilySymbol>(workset, familySymbol));
            }
            this.state = true;
            this.Close();
        }
    }
}
