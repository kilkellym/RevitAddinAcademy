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
    public class CmdAddLegend : IExternalCommand
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

            // get all sheets
            List<ViewSheet> sheetList = GetAllSheets(doc);

            // get selected sheet by prefix
            string sheetFilter = "A2";
            List<ViewSheet> filteredSheetList = FilterSheetListByPrefix(sheetList, sheetFilter);

            // get legend view
            string legendName = "Legend 1";
            View legendView = GetViewByName(doc, legendName);

            // add legend to sheet
            int counter = 0;
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Add legend to sheet");
                counter = AddLegendToSheets(doc, filteredSheetList, legendView, new XYZ(2.75, 0.5, 0));
                t.Commit();
            }
             
            // alert user
            TaskDialog.Show("Complete", "Added legend to " + counter.ToString() + " sheet.");
            return Result.Succeeded;
        }

        private int AddLegendToSheets(Document doc, List<ViewSheet> sheetList, View legendView, XYZ insPoint)
        {
            int counter = 0;
            foreach(ViewSheet curSheet in sheetList)
            {
                Viewport.Create(doc, curSheet.Id, legendView.Id, insPoint);
                counter++;
            }

            return counter;
        }

        private View GetViewByName(Document doc, string legendName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Views);

            foreach(View curView in collector)
            {
                if (curView.Name == legendName)
                    return curView;
            }

            return null;
        }

        private List<ViewSheet> FilterSheetListByPrefix(List<ViewSheet> sheetList, string sheetFilter)
        {
            List<ViewSheet> returnList = new List<ViewSheet>();

            foreach(ViewSheet sheet in sheetList)
            {
                if (sheet.SheetNumber.Contains(sheetFilter))
                    returnList.Add(sheet);
            }

            return returnList;
        }

        private List<ViewSheet> GetAllSheets(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Sheets);

            List<ViewSheet> returnList = new List<ViewSheet>();

            foreach (ViewSheet curSheet in collector)
                returnList.Add(curSheet);

            return returnList;

        }
    }
}
