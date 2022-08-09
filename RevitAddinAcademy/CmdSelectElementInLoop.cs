#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class CmdSelectElementInLoop : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            bool stopSelect = false;

            do
            {
                Reference r = null;
                try
                {
                    r = uidoc.Selection.PickObject(ObjectType.Element, "Select an element");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    TaskDialog.Show("Complete", "Selection complete!");
                    stopSelect = true;
                }

                if (r != null)
                {
                    // get element from reference
                    Element elem = doc.GetElement(r);

                    // do something with element
                    TaskDialog.Show("Complete", elem.Name);
                }
            }
            while (stopSelect == false);

            return Result.Succeeded;
        }
    }
}
