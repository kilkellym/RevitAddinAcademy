#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;

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
            System.IO.Path.GetFileName(modelPath);
           
            ModelPath mp = doc.GetCloudModelPath();
            

            Employee emp1 = new Employee("Joe", 24, "blue,red,white");
            Employee emp2 = new Employee("Mary", 26, "green,red,brown");
            Employee emp3 = new Employee("Felix", 45, "gray,beige");

            List<Employee> empList = new List<Employee>();
            empList.Add(emp1);
            empList.Add(emp2);
            empList.Add(emp3);

            Employees allEmployees = new Employees(empList);

            Debug.Print("There are " + allEmployees.GetEmployeeCount().ToString() + " employees.");

            Debug.Print(Utilities.GetTextFromClass());

            List<SpatialElement> roomList = Utilities.GetAllRooms(doc);

            using(Transaction t = new Transaction(doc))
            {
                t.Start("Insert furniture");

                FamilySymbol curFS = Utilities.GetFamilySymbolByName(doc, "Desk", @"60"" x 30""");
                curFS.Activate();

                foreach (SpatialElement curRoom in roomList)
                {
                    LocationPoint roomLocation = curRoom.Location as LocationPoint;
                    XYZ roomPoint = roomLocation.Point;

                    FamilyInstance curFI = doc.Create.NewFamilyInstance(roomPoint, curFS, StructuralType.NonStructural);

                    double area = Utilities.GetParamValueAsDouble(curRoom, "Area");

                    Utilities.SetParamValue(curRoom, "Comments", "This is a comment");
                }

                t.Commit();
            }
            
            
            return Result.Succeeded;
        }
    }
}
