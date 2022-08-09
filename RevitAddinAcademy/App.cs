#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;

#endregion

namespace RevitAddinAcademy
{
    internal class App : IExternalApplication
    {
        public string logPath = @"C:\temp\_RevitLog.csv";

        public Result OnStartup(UIControlledApplication a)
        {
            try
            {
                // Register event. 
                a.ControlledApplication.DocumentOpened += new EventHandler<DocumentOpenedEventArgs>(application_DocumentOpened);
                a.ControlledApplication.DocumentClosing += new EventHandler<DocumentClosingEventArgs> (application_DocumentClosing);
            }
            catch (Exception)
            {
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            // remove the event.
            a.ControlledApplication.DocumentOpened -= application_DocumentOpened;
            a.ControlledApplication.DocumentClosing -= application_DocumentClosing;

            return Result.Succeeded;
        }

        public void application_DocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            // get document from event args.
            Document doc = args.Document;

            WriteToLog(doc.PathName, "open");
        }

        public void application_DocumentClosing(object sender, DocumentClosingEventArgs args)
        {
            Document doc = args.Document;

            WriteToLog(doc.PathName, "close");
        }

        private void WriteToLog(string pathName, string mode)
        {
            string logEntry = mode + "," + pathName + "," + DateTime.Now.ToString();

            using (StreamWriter writer = File.AppendText(logPath))
            {
                writer.WriteLine(logEntry);
            }
        }
    }
}
