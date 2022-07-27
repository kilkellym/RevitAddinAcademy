using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections;
using System.Diagnostics;
using Autodesk.Revit.DB.Architecture;

namespace RevitAddinAcademy
{
    internal class myClass
    {
    }

    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public List<string> FavColors { get; set; }

        public Employee(string name, int age, string favColors)
        {
            Name = name;
            Age = age;
            //FavColors = FormatColorList(favColors);
        }

        //private List<string> FormatColorList(string colorList)
        //{
        //    //List<string> returnList = colorList.Split(',').ToList();
        //    //return returnList;
        //}
    }

    public class Employees
    {
        public List<Employee> EmployeeList { get; set; }

        public Employees(List<Employee> employees)
        {
            EmployeeList = employees;
        }

        public int GetEmployeeCount()
        {
            return EmployeeList.Count;
        }

    }

    public static class Utilities
    {
        public static string GetTextFromClass()
        {
            return "I got this text from a static class";
        }

        public static List<SpatialElement> GetAllRooms(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Rooms);

            List<SpatialElement> roomList = new List<SpatialElement>();

            foreach(Element curElem in collector)
            {
                SpatialElement curRoom = curElem as SpatialElement;
                roomList.Add(curRoom);
            }

            return roomList;
        }

        public static FamilySymbol GetFamilySymbolByName(Document doc, string familyName, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Family));

            foreach(Element element in collector)
            {
                Family curFamily = element as Family;
                 
                if(curFamily.Name == familyName)
                {
                    ISet<ElementId> famSymbolIdList = curFamily.GetFamilySymbolIds();

                    foreach(ElementId famSymbolId in famSymbolIdList)
                    {
                        FamilySymbol curFS = doc.GetElement(famSymbolId) as FamilySymbol;

                        if (curFS.Name == typeName)
                            return curFS;
                    }
                }
            }

            return null;
        }

        public static string GetParamValueAsString(Element curElem, string paramName)
        {
            foreach(Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                    return curParam.AsString();
            }

            return null;
        }

        public static double GetParamValueAsDouble(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                    return curParam.AsDouble();
            }

            return 0;
        }

        public static void SetParamValue(Element curElem, string paramName, string paramValue)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                    curParam.Set(paramValue);
            }
        }

        public static void SetParamValue(Element curElem, string paramName, double paramValue)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                    curParam.Set(paramValue);
            }
        }


    }
}
