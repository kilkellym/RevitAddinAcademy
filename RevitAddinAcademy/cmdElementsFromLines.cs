#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class cmdElementsFromLines : IExternalCommand
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

            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select Elements");

            Level curLevel = GetLevelByName(doc, "Level 1");
            MEPSystemType pipeSystemType = GetMEPSystemTypeByName(doc, "Domestic Hot Water");
            MEPSystemType ductSystemType = GetMEPSystemTypeByName(doc, "Supply Air");
            PipeType pipeType = GetPipeTypeByName(doc, "Default");
            DuctType ductType = GetDuctTypeByName(doc, "Default");
            WallType wallType1 = GetWallTypeByName(doc, @"Generic - 8""");
            WallType wallType2 = GetWallTypeByName(doc, "Storefront");
            int counter = 0;
            int errorCounter = 0;

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create elements from lines");

                foreach (Element element in pickList)
                {
                    if (element is CurveElement)
                    {
                        CurveElement curveElement = (CurveElement)element;
                        GraphicsStyle curGS = curveElement.LineStyle as GraphicsStyle;

                        Curve curCurve = curveElement.GeometryCurve;

                        XYZ startPoint;
                        XYZ endPoint;

                        try
                        {
                            startPoint = curCurve.GetEndPoint(0);
                            endPoint = curCurve.GetEndPoint(1);
                        }
                        catch (Exception)
                        {
                            errorCounter++;
                            startPoint = null;
                            endPoint = null;
                        }

                        try
                        {
                            switch (curGS.Name)
                            {
                                case "A-GLAZ":
                                    Wall.Create(doc, curCurve, wallType2.Id, curLevel.Id, 20, 0, false, false);
                                    counter++;
                                    break;

                                case "A-WALL":
                                    Wall.Create(doc, curCurve, wallType1.Id, curLevel.Id, 20, 0, false, false);
                                    counter++;
                                    break;

                                case "M-DUCT":
                                    if (startPoint != null & endPoint != null)
                                    {
                                        Duct.Create(doc, ductSystemType.Id, ductType.Id, curLevel.Id, startPoint, endPoint);
                                        counter++;
                                    }
                                    break;

                                case "P-PIPE":
                                    if (startPoint != null && endPoint != null)
                                    {
                                        Pipe.Create(doc, pipeSystemType.Id, pipeType.Id, curLevel.Id, startPoint, endPoint);
                                        counter++;
                                    }

                                    break;

                                default:
                                    break;
                            }
                        }
                        catch (Exception)
                        {
                            TaskDialog.Show("Error", "can't create element");
                            throw;
                        }
                        

                    }
                }

                t.Commit();
            }

            TaskDialog.Show("Complete", "Created " + counter.ToString() + " elements.");

            if (errorCounter > 0)
                TaskDialog.Show("Error", "I couldn't create " + errorCounter.ToString() + " elements.");
            
            return Result.Succeeded;
        }

        private WallType GetWallTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            foreach(Element element in collector)
            {
                WallType curType = element as WallType;

                if (curType.Name == typeName)
                    return curType;
            }

            return null;
        }

        private DuctType GetDuctTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(DuctType));

            foreach (Element element in collector)
            {
                DuctType curType = element as DuctType;

                if (curType.Name == typeName)
                    return curType;
            }

            return null;
        }

        private PipeType GetPipeTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(PipeType));

            foreach (Element element in collector)
            {
                PipeType curType = element as PipeType;

                if (curType.Name == typeName)
                    return curType;
            }

            return null;
        }

        private MEPSystemType GetMEPSystemTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(MEPSystemType));

            foreach (Element element in collector)
            {
                MEPSystemType curType = element as MEPSystemType;

                if (curType.Name == typeName)
                    return curType;
            }

            return null;
        }

        private Level GetLevelByName(Document doc, string levelName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Level));
            collector.WhereElementIsNotElementType();

            foreach (Element element in collector)
            {
               Level curLevel = element as Level;

                if (curLevel.Name == levelName)
                    return curLevel;
            }

            return null;
        }
    }
}
