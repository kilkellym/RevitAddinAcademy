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
    public class cmdDeleteUnusedViews : IExternalCommand
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

			// list of views to be deleted
			List<View> viewsToDelete = new List<View>();
			List<ViewSchedule> schedulesToDelete = new List<ViewSchedule>();

			// get all views in the project
			FilteredElementCollector viewColl = new FilteredElementCollector(doc);
			viewColl.OfCategory(BuiltInCategory.OST_Views);

			// get all schedules in the project
			FilteredElementCollector schedColl = new FilteredElementCollector(doc);
			schedColl.OfClass(typeof(ViewSchedule));

			// get all sheets in the project
			FilteredElementCollector sheetColl = new FilteredElementCollector(doc);
			sheetColl.OfClass(typeof(ViewSheet));

			// make sure there are sheets in the project
			if (sheetColl.GetElementCount() < 1)
			{
				TaskDialog.Show("Error", "There are no sheets in this project. Please add one before running the macro.");
				return Result.Failed;
			}

			foreach (View curView in viewColl)
			{
				// check if view has prefix
				if (curView.Name.Contains("working_") == false)
				{
					// check if view already on sheet
					if (Viewport.CanAddViewToSheet(doc, sheetColl.FirstElementId(), curView.Id) == true)
					{
						// check if view has dependent views
						if (curView.GetDependentViewIds().Count == 0)
						{
							// add to list of views to be deleted
							viewsToDelete.Add(curView);
						}
					}
				}
			}

			foreach(ViewSchedule curSched in schedColl)
            {
				if(curSched.Name.Contains("working_") == false)
                {
					if(IsScheduleOnSheet(doc, curSched) == false)
                    {
						schedulesToDelete.Add(curSched);
                    }
				}
            }

			// start a transaction
			Transaction curTrans = new Transaction(doc, "Delete unused views and schedules");
			curTrans.Start();

			// loop through views to delete and delete from file
			foreach (View viewToDelete in viewsToDelete)
			{
				string curViewName = "";
                try
				{      
					curViewName = viewToDelete.Name;
					doc.Delete(viewToDelete.Id);
				}
				catch (Exception)
				{
					TaskDialog.Show("Error", "Could not delete view: " + curViewName);
				}
			}

			foreach(ViewSchedule schedToDelete in schedulesToDelete)
            {
				string curSchedName = "";
				try
                {
					curSchedName = schedToDelete.Name;
					doc.Delete(schedToDelete.Id);
                }
				catch(Exception)
                {
					TaskDialog.Show("Error", "Could not delete schedule: " + curSchedName);
                }
            }

			//close transaction
			curTrans.Commit();
			curTrans.Dispose();

			//alert user
			TaskDialog.Show("Complete", "Deleted " + viewsToDelete.Count.ToString() + " views." + Environment.NewLine 
				+ "Deleted " + schedulesToDelete.Count.ToString() + " schedules.");

			return Result.Succeeded;
		}

		private bool IsScheduleOnSheet(Document doc, ViewSchedule curSched)
        {
			FilteredElementCollector collector = new FilteredElementCollector(doc);
			collector.OfClass(typeof(ScheduleSheetInstance));

			foreach(Element elem in collector)
            {
				ScheduleSheetInstance curSSI = elem as ScheduleSheetInstance;

				if(curSSI.ScheduleId == curSched.Id)
					return true;
            }

			return false;
        }
    }
}
