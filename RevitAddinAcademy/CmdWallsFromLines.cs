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
    public class CmdWallsFromLines : IExternalCommand
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

            List<string> wallTypes = GetAllWallTypeNames(doc);
            List<string> lineStyles = GetAllLineStyleNames(doc);

            FrmWallsFromLines curForm = new FrmWallsFromLines(wallTypes, lineStyles);
            curForm.Height = 450;
            curForm.Width = 550;
            curForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

            if(curForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string selectedWallType = curForm.GetSelectedWallType();
                string selectedLineStyle = curForm.GetSelectedLineStyle();
                double wallHeight = curForm.GetWalLHeight();
                bool isStructural = curForm.AreWallsStructural();

                WallType curWT = GetWallTypeByName(doc, selectedWallType);
                List<CurveElement> curveList = GetLinesByStyle(doc, selectedLineStyle);
                Level curLevel = GetLevelFromView(doc);

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Create walls from lines");

                    foreach(CurveElement curve in curveList)
                    {
                        Curve curCurve = curve.GeometryCurve;

                        Wall newWall = Wall.Create(
                            doc,
                            curCurve,
                            curWT.Id,
                            curLevel.Id,
                            wallHeight,
                            0,
                            false,
                            isStructural);
                    }

                    t.Commit();
                }

            }

            return Result.Succeeded;
        }

        private Level GetLevelFromView(Document doc)
        {
            View curView = doc.ActiveView;

            SketchPlane curSP = curView.SketchPlane;

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Level));
            collector.WhereElementIsNotElementType();

            foreach(Level curLevel in collector)
            {
                if (curLevel.Name == curSP.Name)
                    return curLevel;
            }

            return collector.FirstElement() as Level;
        }

        private List<CurveElement> GetLinesByStyle(Document doc, string selectedLineStyle)
        {
            List<CurveElement> results = new List<CurveElement>();

            FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
            collector.OfClass(typeof(CurveElement));

            foreach (CurveElement element in collector)
            {
                GraphicsStyle curGS = element.LineStyle as GraphicsStyle;

                if (curGS.Name == selectedLineStyle)
                {
                    results.Add(element);
                }
            }

            return results;
        }

        private WallType GetWallTypeByName(Document doc, string selectedWallType)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            foreach (WallType wallType in collector)
            {
                if(wallType.Name == selectedWallType)
                    return wallType; 
            }

            return null;
        }

        private List<string> GetAllLineStyleNames(Document doc)
        {
            List<string> results = new List<string>();

            FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
            collector.OfClass(typeof(CurveElement));

            foreach(CurveElement element in collector)
            {
                GraphicsStyle curGS = element.LineStyle as GraphicsStyle;

                if(results.Contains(curGS.Name) == false)
                {
                    results.Add(curGS.Name);
                }
            }

            return results;

        }

        private List<string> GetAllWallTypeNames(Document doc)
        {
            List<string> results = new List<string>();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            foreach(WallType wallType in collector)
            {
                results.Add(wallType.Name);
            }

            return results;
        }
    }
}
