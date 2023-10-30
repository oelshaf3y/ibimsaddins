using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Interop;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace IBIMSGen
{
    internal class App : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Failed;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            String assemblyName = Assembly.GetExecutingAssembly().Location;
            String asPath = System.IO.Path.GetDirectoryName(assemblyName);
            String tabName = "IBIMS";
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch { }
            RibbonPanel genTools = application.CreateRibbonPanel(tabName, "General Tools");
            RibbonPanel systems = application.CreateRibbonPanel(tabName, "System Tools");
            RibbonPanel Arch = application.CreateRibbonPanel(tabName, "Arch. Tools");


            PushButtonData SelectSimilarDL = new PushButtonData("Similar Detail Line", "Similar Lines", assemblyName, "IBIMSGen.SelectSimilar")
            {
                LargeImage = Properties.Resources.similar.ToImageSource(),
                Image = Properties.Resources.similar_s.ToImageSource(),
                ToolTip = "Find and Select all Lines with the same line style in the current view."
            };

            PushButtonData ReplaceElements = new PushButtonData("Find & Replace Elements", "Replace Elements", assemblyName, "IBIMSGen.ReplaceFamilies.Replace")
            {
                LargeImage = Properties.Resources.replace2.ToImageSource(),
                Image = Properties.Resources.replace2_s.ToImageSource(),
                ToolTip = "Find all instances of element and replace it with another element"
            };

            PushButtonData RoomSections = new PushButtonData("Room Sheets", "Room Sheet", assemblyName, "IBIMSGen.Rooms.Rooms")
            {
                LargeImage = Properties.Resources.Room.ToImageSource(),
                Image = Properties.Resources.Room_s.ToImageSource(),
                ToolTip = "Create a sheet for each selected room containing all wall elevations, flooring plans and ceiling plan."
            };

            PushButtonData copyCallouts = new PushButtonData("Copy all Callouts", "Copy Callouts", assemblyName, "IBIMSGen.CallOutsCopy")
            {
                LargeImage = Properties.Resources.callouts.ToImageSource(),
                Image = Properties.Resources.callouts_s.ToImageSource(),
                ToolTip = "Copy all callouts in all view ports in a specific sheet to the corresponding view in other project."
            };

            PushButtonData clash = new PushButtonData("Clash Finder", "Clashes Finder", assemblyName, "IBIMSGen.ClashViewer.clashPoints")
            {
                LargeImage = Properties.Resources.clash.ToImageSource(),
                Image = Properties.Resources.clash_s.ToImageSource(),
                ToolTip = "Retrive clashes from navisworks"
            };

            PushButtonData BWork = new PushButtonData("Penetrate structural elements using sleeves", "Builder Work", assemblyName, "IBIMSGen.Penetration.Penetration")
            {
                LargeImage = Properties.Resources.bw.ToImageSource(),
                Image = Properties.Resources.bw_s.ToImageSource(),
                ToolTip = "Insert sleeves of a choosen type at penetration location."
            };

            PushButtonData Fillet = new PushButtonData("Fillet Electrical Lines", "Fillet Lines", assemblyName, "IBIMSGen.ElecCables.Fillet")
            {
                LargeImage = Properties.Resources.fillet.ToImageSource(),
                Image = Properties.Resources.fillet_s.ToImageSource(),
                ToolTip = "Fillet lines with the same line style at the ends where they meet."
            };

            PushButtonData CutLines = new PushButtonData("Cut Electrical Lines", "Cut Lines", assemblyName, "IBIMSGen.ElecCables.CablesTrim")
            {
                LargeImage = Properties.Resources.CutLines.ToImageSource(),
                Image = Properties.Resources.CutLines_s.ToImageSource(),
                ToolTip = "Cut intersecting lines with the selection set of lines."
            };

            PushButtonData CenterElement = new PushButtonData("Center Element 2pts", "Center Element", assemblyName, "IBIMSGen.AlignBetween2Pts")
            {
                LargeImage = Properties.Resources.centerElement.ToImageSource(),
                Image = Properties.Resources.centerElement_s.ToImageSource(),
                ToolTip = "Allign the mid point of the element between two points in 2D"
            };

            PushButtonData Lights = new PushButtonData("Distribute Lighting Fixtures", "Distribute Fixtures", assemblyName, "IBIMSGen.ElecEquipCeilings.CeilingLightsPlacement")
            {
                LargeImage = Properties.Resources.gridLights.ToImageSource(),
                Image = Properties.Resources.gridLights_s.ToImageSource(),
                ToolTip = "Distribute Lighting Fixtures in a certain area"
            };


            PushButtonData ColorCables = new PushButtonData("Color code for wires", "Color Wires", assemblyName, "IBIMSGen.ColorCable.Coloring")
            {
                LargeImage = Properties.Resources.colorCables.ToImageSource(),
                Image = Properties.Resources.colorCables_s.ToImageSource(),
                ToolTip = "Color Code for wires in the project (find & replace method)"
            };
            PushButtonData cableTrayHeights = new PushButtonData("Tray's Height Offset", "Trays' Height", assemblyName, "IBIMSGen.ElecCables.CableTraysElevations")
            {
                LargeImage = Properties.Resources.cableTray.ToImageSource(),
                Image = Properties.Resources.cableTray_s.ToImageSource(),
                ToolTip = "Get Trays' Height Offset From The Nearest Slab Above."
            };

            PushButtonData RemoveCad = new PushButtonData("Remove CAD Imports", "Remove CAD", assemblyName, "IBIMSGen.DeleteCad")
            {
                LargeImage = Properties.Resources.delete_cad.ToImageSource(),
                Image = Properties.Resources.delete_cad_s.ToImageSource(),
                ToolTip = "Delete All CAD Imports From The Entire Project."
            };

            PushButtonData sections = new PushButtonData("Make 2 Sections", "2 Sections", assemblyName, "IBIMSGen.ElementSections")
            {
                LargeImage = Properties.Resources.section.ToImageSource(),
                Image = Properties.Resources.section_s.ToImageSource(),
                ToolTip = "Make two perpendicular sections for the selected element(s)."
            };

            genTools.AddItem(copyCallouts);
            genTools.AddItem(clash);
            genTools.AddStackedItems(ReplaceElements,SelectSimilarDL);
            genTools.AddStackedItems(RemoveCad, sections);

            Arch.AddItem(RoomSections);
            systems.AddItem(BWork);
            systems.AddItem(Lights);
            systems.AddStackedItems(CenterElement, ColorCables);
            systems.AddStackedItems(Fillet, CutLines,cableTrayHeights);
            return Result.Succeeded;
        }
    }
    public static class Methods
    {
        public static ImageSource ToImageSource(this Icon icon)
        {
            ImageSource imageSource = Imaging.CreateBitmapSourceFromHIcon(
                icon.Handle,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());

            return imageSource;
        }
    }
}
