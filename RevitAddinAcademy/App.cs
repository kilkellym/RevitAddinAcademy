#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;

#endregion

namespace RevitAddinAcademy
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            TaskDialog.Show("Hello", "Welcome to Revit Add-in Academy");
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            TaskDialog.Show("Hello", "Leaving Revit Add-in Academy");
            return Result.Succeeded;
        }
    }
}
