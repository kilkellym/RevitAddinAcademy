#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.IO;

#endregion

namespace RevitAddinAcademy
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            // step 1. create ribbon tab
            try
            {
                a.CreateRibbonTab("Revit Add-in Academy");
            }
            catch (Exception)
            {
                Debug.Print("Tab already exists");
            }

            // step 2. create ribbon panel
            RibbonPanel curRibbonPanel = CreateRibbonPanel(a, "Revit Add-in Academy", "Revit Tools");

            // step 3. create button data instances
            string assemblyName = GetAssemblyName();
            PushButtonData pData1 = new PushButtonData("tool1", "Delete Backups", GetAssemblyName(), "RevitAddinAcademy.cmdDeleteBackups");
            PushButtonData pData2 = new PushButtonData("tool2", "Elements from Lines", GetAssemblyName(), "RevitAddinAcademy.cmdElementsFromLines");
            PushButtonData pData3 = new PushButtonData("tool3", "Tool 3", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData4 = new PushButtonData("tool4", "Tool 4", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData5 = new PushButtonData("tool5", "Tool 5", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData6 = new PushButtonData("tool6", "Tool 6", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData7 = new PushButtonData("tool7", "Tool 7", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData8 = new PushButtonData("tool8", "Tool 8", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData9 = new PushButtonData("tool9", "Tool 9", GetAssemblyName(), "RevitAddinAcademy.Command");
            PushButtonData pData10 = new PushButtonData("tool10", "Tool 10", GetAssemblyName(), "RevitAddinAcademy.Command");

            PulldownButtonData pdData1 = new PulldownButtonData("pulldown1", "More Tools");
            SplitButtonData sbData1 = new SplitButtonData("splitButton1", "Split Button 1");

            // step 4. add images
            pData1.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_32);
            pData1.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_16);

            pData2.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Yellow_32);
            pData2.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Yellow_16);

            pData3.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Red_32);
            pData3.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Red_16);

            pData4.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Green_32);
            pData4.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Green_16);

            pData5.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_32);
            pData5.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_16);

            pData6.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Red_32);
            pData6.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Red_16);
           
            pData7.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Yellow_32);
            pData7.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Yellow_16);
            
            pData8.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_32);
            pData8.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_16);
            
            pData9.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_32);
            pData9.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_16);

            pData10.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_32);
            pData10.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Blue_16);

            pdData1.LargeImage = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Green_32);
            pdData1.Image = BitmapToImageSource(RevitAddinAcademy.Properties.Resources.Green_16);

            // step 5. add tool tips
            pData1.ToolTip = "Tool tip 1";
            pData2.ToolTip = "Tool tip 2";
            pData3.ToolTip = "Tool tip 3";
            pData4.ToolTip = "Tool tip 4";
            pData5.ToolTip = "Tool tip 5";
            pData6.ToolTip = "Tool tip 6";
            pData7.ToolTip = "Tool tip 7";
            pData8.ToolTip = "Tool tip 8";
            pData9.ToolTip = "Tool tip 9";
            pData10.ToolTip = "Tool tip 10";
            pdData1.ToolTip = "Pulldown tool tip";

            // step 6. create buttons
            PushButton button1 = curRibbonPanel.AddItem(pData1) as PushButton;
            PushButton button2 = curRibbonPanel.AddItem(pData2) as PushButton;

            curRibbonPanel.AddStackedItems(pData3, pData4, pData5);

            SplitButton sb1 = curRibbonPanel.AddItem(sbData1) as SplitButton;
            sb1.AddPushButton(pData6);
            sb1.AddPushButton(pData7);

            PulldownButton pb1 = curRibbonPanel.AddItem(pdData1) as PulldownButton;
            pb1.AddPushButton(pData8);
            pb1.AddPushButton(pData9);
            pb1.AddPushButton(pData10);

            return Result.Succeeded;
        }

        private BitmapImage BitmapToImageSource(Bitmap bm)
        {
            using(MemoryStream mem = new MemoryStream())
            {
                bm.Save(mem, System.Drawing.Imaging.ImageFormat.Png);
                mem.Position = 0;
                BitmapImage bmi = new BitmapImage();
                bmi.BeginInit();
                bmi.StreamSource = mem;
                bmi.CacheOption = BitmapCacheOption.OnLoad;
                bmi.EndInit();

                return bmi;
            }
        }

        private string GetAssemblyName()
        {
            return Assembly.GetExecutingAssembly().Location;
        }

        private RibbonPanel CreateRibbonPanel(UIControlledApplication a, string tabName, string panelName)
        {
            foreach(RibbonPanel tmpPanel in a.GetRibbonPanels(tabName))
            {
                if(tmpPanel.Name == panelName)
                    return tmpPanel;
            }

            RibbonPanel returnPanel = a.CreateRibbonPanel(tabName, panelName);

            return returnPanel;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            TaskDialog.Show("Hello", "Leaving Revit Add-in Academy");
            return Result.Succeeded;
        }
    }
}
