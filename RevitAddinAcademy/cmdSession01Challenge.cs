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
    public class cmdSession01Challenge : IExternalCommand
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

            int range = 100;
            XYZ insPoint = new XYZ(0, 0, 0);
            double offset = 0.041;
            double calcOffset = offset * doc.ActiveView.Scale;
            XYZ offsetPoint = new XYZ(0, calcOffset, 0);

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(TextNoteType));

            //------------------------------------------------------
            // calculate fizzbuzz then create text note
            using(Transaction t = new Transaction(doc))
            {
                t.Start("FizzBuzz");

                for (int i = 1; i <= range; i++)
                {
                    string result = CheckFizzBuzz(i);

                    CreateTextNote(doc, result, insPoint, collector.FirstElementId());
                    insPoint = insPoint.Subtract(offsetPoint);
                }

                t.Commit();
            }
            
            return Result.Succeeded;
        }

        internal void CreateTextNote(Document doc, string text, XYZ insPoint, ElementId id)
        {
            TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, insPoint, text, id);
        }

        internal string CheckFizzBuzz(int number)
        {
            string result = "";

            if (number % 3 == 0)
            {
                result = "FIZZ";
            }

            if (number % 5 == 0)
            {
                result = result + "BUZZ";
            }

            if (number % 3 != 0 && number % 5 != 0)
            {
                result = number.ToString();
            }

            Debug.Print(result);

            return result;
        }
    }
}
