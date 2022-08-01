#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Forms = System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class CmdInsertFurniture : IExternalCommand
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

            int counter = 0;

            string excelFile = "";

            Forms.OpenFileDialog ofd = new Forms.OpenFileDialog();
            ofd.Title = "Select Furniture Excel File";
            ofd.InitialDirectory = @"C:\";
            ofd.Filter = "Excel files (*.xlsx)|*.xlsx";

            if (ofd.ShowDialog() != Forms.DialogResult.OK)
                return Result.Failed;

            excelFile = ofd.FileName;

            List<string[]> excelFurnSetData = GetDataFromExcel(excelFile, "Furniture sets", 3);
            List<string[]> excelFurnData = GetDataFromExcel(excelFile, "Furniture types", 3);

            excelFurnSetData.RemoveAt(0);
            excelFurnData.RemoveAt(0);

            List<FurnSet> furnSetList = new List<FurnSet>();
            List<FurnData> furnDataList = new List<FurnData>();

            foreach (string[] curRow in excelFurnSetData)
            {
                FurnSet tmpFurnSet = new FurnSet(curRow[0].Trim(), curRow[1].Trim(), curRow[2].Trim());
                furnSetList.Add(tmpFurnSet);
            }

            foreach (string[] curRow in excelFurnData)
            {
                FurnData tmpFurnData = new FurnData(doc, curRow[0].Trim(), curRow[1].Trim(), curRow[2].Trim());
                furnDataList.Add(tmpFurnData);
            }

            List<SpatialElement> roomList = GetAllRooms(doc);

            using(Transaction t = new Transaction(doc))
            {
                t.Start("Insert Furniture");

                foreach (SpatialElement room in roomList)
                {
                    string curFurnSet = GetParamValue(room, "Furniture Set");

                    LocationPoint roomPt = room.Location as LocationPoint;
                    XYZ insPoint = roomPt.Point;

                    foreach (FurnSet tmpFurnSet in furnSetList)
                    {
                        if (tmpFurnSet.setType == curFurnSet)
                        {
                            foreach (string curFurn in tmpFurnSet.furnList)
                            {
                                string tmpFurn = curFurn.Trim();
                                FurnData fd = GetFamilyInfo(tmpFurn, furnDataList);

                                if (fd != null)
                                {
                                    fd.familySymbol.Activate();

                                    FamilyInstance newFamInst = doc.Create.NewFamilyInstance(insPoint, fd.familySymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                    counter++;

                                }
                            }
                        }
                        SetParamValueAsInt(room, "Furniture Count", tmpFurnSet.FurnitureCount());
                    }
                }

                t.Commit();
            }

            TaskDialog.Show("Complete", "Inserted " + counter.ToString() + " families.");

            return Result.Succeeded;
        }

        private void SetParamValueAsInt(Element curElem, string paramName, int paramValue)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                {
                    curParam.Set(paramValue);
                }
            }
        }

        private FurnData GetFamilyInfo(string furnName, List<FurnData> furnDataList)
        {
            foreach (FurnData furn in furnDataList)
                if (furn.furnName == furnName)
                    return furn;

            return null;
        }

        private string GetParamValue(Element curElem, string paramName)
        {
            foreach(Parameter curParam in curElem.Parameters)
            {
                Debug.Print(curParam.Definition.Name);
                if (curParam.Definition.Name == paramName)
                {
                    Debug.Print(curParam.AsString());
                    return curParam.AsString();
                }
                    
            }

            return null;
        }

        private List<SpatialElement> GetAllRooms(Document doc)
        {
            List<SpatialElement> returnList = new List<SpatialElement>();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Rooms);
            collector.WhereElementIsNotElementType();

            foreach(Element curElem in collector)
            {
                SpatialElement curRoom = curElem as SpatialElement;
                returnList.Add(curRoom);
            }

            return returnList;
        }

        private List<string[]> GetDataFromExcel(string excelFile, string wsName, int numColumns)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);

            Excel.Worksheet excelWs = GetExcelWorksheetByName(excelWb, wsName);
            Excel.Range excelRng = excelWs.UsedRange as Excel.Range;

            int rowCount = excelRng.Rows.Count;

            List<string[]> data = new List<string[]>();

            for(int i = 1; i <= rowCount; i++)
            {
                string[] rowData = new string[numColumns];

                for(int j = 1; j <= numColumns; j++)
                {
                    Excel.Range cellData = excelWs.Cells[i, j];
                    rowData[j - 1] = cellData.Value.ToString();
                }

                data.Add(rowData);
            }

            excelWb.Close();
            excelApp.Quit();

            return data;
        }

        private Excel.Worksheet GetExcelWorksheetByName(Excel.Workbook excelWb, string wsName)
        {
            foreach(Excel.Worksheet sheet in excelWb.Worksheets)
            {
                if (sheet.Name == wsName)
                    return sheet;
            }

            return null;
        }
    }
}
