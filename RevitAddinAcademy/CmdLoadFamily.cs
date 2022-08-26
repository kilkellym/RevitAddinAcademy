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

            using(Transaction t = new Transaction(doc))
            {
                t.Start("Load family and insert");

                // is family loaded - if not, load it
                if (IsFamilyLoaded(doc, familyName) == false)
                {
                    if (doc.LoadFamily(familyPath) == true)
                        TaskDialog.Show("Load Family", "Loaded family: " + familyName);
                    else
                    {
                        TaskDialog.Show("Load Family", "Could not load family");
                        return Result.Failed;
                    }
                }

                // get family object
                Family curFam = GetFamilyByName(doc, familyName);

                if (curFam != null)
                {
                    // does family have symbol - if not, create it
                    if (DoesFamilySymbolExist(doc, curFam, familySymbolName) == false)
                    {
                        CreateFamilySymbol(doc, curFam, familySymbolName);
                        TaskDialog.Show("Created family symbol", "Created family symbol:" + familySymbolName);
                    }

                    // insert family
                    FamilySymbol curFS = GetFamilySymbolByName(doc, curFam, familySymbolName);

                    if (curFS != null)
                    {
                        FamilyInstance curFI = InsertFamilySymbol(doc, curFS, insPoint);
                    }
                    else
                    {
                        TaskDialog.Show("Error", "Could not find family symbol");
                    }
                }

                t.Commit();
            }

            return Result.Succeeded;
        }

        private FamilyInstance InsertFamilySymbol(Document doc, FamilySymbol curFS, XYZ insPoint)
        {
            FamilyInstance curFI = doc.Create.NewFamilyInstance(insPoint, curFS, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

            return curFI;
        }

        private FamilySymbol GetFamilySymbolByName(Document doc, Family curFam, string familySymbolName)
        {
            List<FamilySymbol> familySymbols = GetFamilySymbolsFromFamily(doc, curFam);

            foreach (FamilySymbol curFS in familySymbols)
            {
                if (curFS.Name == familySymbolName)
                    return curFS;
            }

            return null;
        }

        private void CreateFamilySymbol(Document doc, Family curFam, string familySymbolName)
        {
            List<FamilySymbol> fsList = GetFamilySymbolsFromFamily(doc, curFam);

            FamilySymbol curFS = fsList[0];

            ElementType newFS = curFS.Duplicate(familySymbolName);

            SetParameterValue(newFS, "Width", 8);
        }

        private void SetParameterValue(ElementType newFS, string paramName, double value)
        {
            foreach(Parameter curParam in newFS.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                    curParam.Set(value);
            }
        }

        private bool DoesFamilySymbolExist(Document doc, Family curFam, string familySymbolName)
        {
            List<FamilySymbol> familySymbols = GetFamilySymbolsFromFamily(doc, curFam);

            foreach(FamilySymbol curFS in familySymbols)
            {
                if (curFS.Name == familySymbolName)
                    return true;
            }

            return false;
        }

        private List<FamilySymbol> GetFamilySymbolsFromFamily(Document doc, Family curFam)
        {
            List<FamilySymbol> returnList = new List<FamilySymbol>();

            ISet<ElementId> fsList = curFam.GetFamilySymbolIds();

            foreach(ElementId curID in fsList)
            {
                FamilySymbol curFS = doc.GetElement(curID) as FamilySymbol;
                returnList.Add(curFS);
            }

            return returnList;
        }

        private Family GetFamilyByName(Document doc, string familyName)
        {
            List<Family> famList = GetAllFamilies(doc);

            foreach (Family curFam in famList)
            {
                if (curFam.Name == familyName)
                    return curFam;
            }

            return null;
        }

        private bool IsFamilyLoaded(Document doc, string familyName)
        {
            List<Family> famList = GetAllFamilies(doc);

            foreach(Family curFam in famList)
            {
                if (curFam.Name == familyName)
                    return true;
            }

            return false;
        }

        private List<Family> GetAllFamilies(Document doc)
        {
            List<Family> returnList = new List<Family>();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Family));

            foreach(Element curElem in collector)
            {
                Family curFam = curElem as Family;
                returnList.Add(curFam);
            }

            return returnList;
        }
    }
}
