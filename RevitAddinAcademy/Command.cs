#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
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

            string modelPath = doc.PathName;
            string fileName = Path.GetFileName(modelPath);
            string fileName2 = Path.GetFileNameWithoutExtension(fileName);
            string folderPath = Path.GetDirectoryName(modelPath);
            string txtFile = folderPath + "\\" + fileName2 + ".txt";

            List<string> stringList = new List<string>();
            stringList.Add("line 1");
            stringList.Add("line 2");
            stringList.Add("line 3");

            using (StreamWriter writer = File.CreateText(txtFile))
            {
                foreach (string curLine in stringList)
                {
                    writer.WriteLine(curLine);

                }
            }

            if (File.Exists(txtFile))
            {
                string[] textFile = File.ReadAllLines(txtFile);

                foreach(string text in textFile)
                {
                    Debug.Print(text);
                }

            }

            FrmTestForm form1 = new FrmTestForm(txtFile);
            form1.ShowDialog();

            TestData test1 = new TestData("this is a string", "this is another string", 10);

            TaskDialog.Show("test", test1.Combo);

            return Result.Succeeded;
        }
    }
}
