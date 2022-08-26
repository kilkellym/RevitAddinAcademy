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
    public class CmdLoadFamily : IExternalCommand
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

            string familyName = "Test Family";
            string familySymbolName = "Type 1";
            string familyPath = @"C:\temp\Test Family.rfa";
            XYZ insPoint = new XYZ(0, 0, 0);

            // family structure in API - family > family symbol > family instance

            using(Transaction t = new Transaction(doc))
            {
                t.Start("Load family");

                // is family loaded - if not then load
                if (IsFamilyLoaded(doc, familyName) == false)
                {
                    if (doc.LoadFamily(familyPath) == true)
                        TaskDialog.Show("Load family", "Loaded family: " + familyPath);
                    else
                        TaskDialog.Show("Load family", "Could not load family");
                }

                // get family object
                Family curFam = GetFamilyByName(doc, familyName);

                if (curFam != null)
                {
                    // does family symbol exist - if not then create
                    if (DoesFamilySymbolExist(doc, curFam, familySymbolName) == false)
                        CreateFamilySymbol(doc, curFam, familySymbolName);

                    // get family symbol
                    FamilySymbol curFamSymbol = GetFamilySymbolByName(doc, curFam, familySymbolName);

                    // insert family symbol
                    if (curFamSymbol != null)
                    {
                        FamilyInstance curFamInstace = InsertFamilySymbol(doc, curFamSymbol, insPoint);
                    }
                }

                t.Commit();
            }
            
            return Result.Succeeded;
        }

        private FamilyInstance InsertFamilySymbol(Document doc, FamilySymbol curFamSymbol, XYZ insPoint)
        {
            FamilyInstance curFI = doc.Create.NewFamilyInstance(insPoint, curFamSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

            return curFI;
        }

        private FamilySymbol GetFamilySymbolByName(Document doc, Family curFam, string familySymbolName)
        {
            List<FamilySymbol> familySymbolList = GetFamilySymbolsFromFamily(doc, curFam);

            foreach (FamilySymbol curFS in familySymbolList)
            {
                if (curFS.Name == familySymbolName)
                    return curFS;
            }

            return null;
        }

        private void CreateFamilySymbol(Document doc, Family curFam, string famSymbolName)
        {
            List<FamilySymbol> familySymbolList = GetFamilySymbolsFromFamily(doc, curFam);

            FamilySymbol curFS = familySymbolList[0];

            ElementType newFS = curFS.Duplicate(famSymbolName);

            TaskDialog.Show("Test", "Created family symbol: " + famSymbolName);

            SetParameterValue(newFS, "Width", 8);
        }

        internal bool DoesFamilySymbolExist(Document doc, Family curFam, string famSymbolName)
        {
            List<FamilySymbol> familySymbolList = GetFamilySymbolsFromFamily(doc, curFam);

            foreach (FamilySymbol curFS in familySymbolList)
            {
                if (curFS.Name == famSymbolName)
                    return true;
            }

            return false;

        }

        internal bool IsFamilyLoaded(Document doc, string familyName)
        {
            List<Family> famList = GetAllFamilies(doc);

            foreach (Family curFam in famList)
            {
                if (curFam.Name == familyName)
                    return true;
            }

            return false;

        }

        internal Family GetFamilyByName(Document doc, string familyName)
        {
            List<Family> famList = GetAllFamilies(doc);

            foreach (Family curFam in famList)
            {
                if (curFam.Name == familyName)
                {
                    return curFam;
                }
            }

            return null;
        }

        internal List<FamilySymbol> GetFamilySymbolsFromFamily(Document doc, Family curFam)
        {
            List<FamilySymbol> returnList = new List<FamilySymbol>();

            ISet<ElementId> fsList = curFam.GetFamilySymbolIds();

            // loop through family symbol ids and look for match
            foreach (ElementId fsID in fsList)
            {
                FamilySymbol fs = doc.GetElement(fsID) as FamilySymbol;
                returnList.Add(fs);
            }

            return returnList;
        }

        internal List<Family> GetAllFamilies(Document doc)
        {
            List<Family> returnList = new List<Family>();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Family));

            foreach (Element curElem in collector)
            {
                Family curFam = curElem as Family;
                returnList.Add(curFam);
            }

            return returnList;
        }

        public static bool SetParameterValue(Element curElem, string paramName, double value)
        {
            Parameter curParam = GetParameterByName(curElem, paramName);

            if (curParam != null)
            {
                curParam.Set(value);
                return true;
            }

            return false;

        }

        public static Parameter GetParameterByName(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name.ToString() == paramName)
                    return curParam;
            }

            return null;
        }
    }
}
