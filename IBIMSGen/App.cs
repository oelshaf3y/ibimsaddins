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

            Image rep = Properties.Resources.replace;
            Image roo = Properties.Resources.Room;
            Image co = Properties.Resources.callouts;
            Image srch = Properties.Resources.clash;
            Image insu = Properties.Resources.inulation;
            Image bw = Properties.Resources.bw;
            Image sim = Properties.Resources.similar;

            ImageSource search = GetImgSrc(srch);
            ImageSource callouts = GetImgSrc(co);
            ImageSource insulation = GetImgSrc(insu);
            ImageSource bwIS = GetImgSrc(bw);
            ImageSource roomIco = GetImgSrc(roo);
            ImageSource replaceIco = GetImgSrc(rep);
            ImageSource similarIco = GetImgSrc(sim);

            PushButtonData SelectSimilarDL = new PushButtonData("Similar Detail Line", "Similar Lines", assemblyName, "IBIMSGen.SelectSimilar")
            {
                LargeImage = similarIco,
                Image = similarIco,
                ToolTip = "Find and Select all Lines with the same line style in the current view."
            };

            PushButtonData ReplaceElements = new PushButtonData("Find & Replace Elements", "Replace Elements", assemblyName, "IBIMSGen.ReplaceFamilies.Replace")
            {
                LargeImage = replaceIco,
                Image = replaceIco,
                ToolTip = "Find all instances of element and replace it with another element"
            };

            PushButtonData RoomSections = new PushButtonData("Room Sheets", "Room Sheet", assemblyName, "IBIMSGen.Rooms.Rooms")
            {
                LargeImage = roomIco,
                Image = roomIco,
                ToolTip = "Create a sheet for each selected room containing all wall elevations, flooring plans and ceiling plan."
            };

            PushButtonData copyCallouts = new PushButtonData("Copy all Callouts", "Copy Callouts", assemblyName, "IBIMSGen.CallOutsCopy")
            {
                LargeImage = callouts,
                Image = callouts,
                ToolTip = "Copy all callouts in all view ports in a specific sheet to the corresponding view in other project."
            };

            PushButtonData clash = new PushButtonData("Clash Finder", "Clashes Finder", assemblyName, "IBIMSGen.ClashViewer.clashPoints")
            {
                LargeImage = search,
                Image = search,
                ToolTip = "Retrive clashes from navisworks"
            };

            PushButtonData BWork = new PushButtonData( "Penetrate structural elements using sleeves", "Builder Work", assemblyName, "IBIMSGen.Penetration.Penetration")
            {
                LargeImage = bwIS,
                Image = bwIS,
                ToolTip = "Insert sleeves of a choosen type at penetration location."
            };

            genTools.AddItem(copyCallouts);
            genTools.AddItem(clash);
            genTools.AddItem(ReplaceElements);
            genTools.AddItem(SelectSimilarDL);
            Arch.AddItem(RoomSections);
            systems.AddItem(BWork);
            return Result.Succeeded;
        }
        private BitmapSource GetImgSrc(Image img)
        {
            BitmapImage bmi = new BitmapImage();
            using (MemoryStream ms = new MemoryStream())
            {
                img.Save(ms, ImageFormat.Png);
                ms.Position = 0;
                bmi.BeginInit();
                bmi.CacheOption = BitmapCacheOption.OnLoad;
                bmi.UriSource = null;
                bmi.StreamSource = ms;
                bmi.EndInit();
            }
            return bmi;
        }
    }
}
