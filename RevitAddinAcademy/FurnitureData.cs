using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;

namespace RevitAddinAcademy
{
    public class FurnSet
    {
        public string setType { get; set; }
        public string setName { get; set; }
        public List<string> furnList { get; private set; }

        public FurnSet(string _setType, string _setName, string _furnList)
        {
            setType = _setType;
            setName = _setName;
            furnList = GetFurnListFromString(_furnList);
        }

        private List<string> GetFurnListFromString(string list)
        {
            List<string> returnList = list.Split(',').ToList();
            List<string> returnList2 = new List<string>();

            foreach (string str in returnList)
                returnList2.Add(str.Trim());

            return returnList;
        }

        public int FurnitureCount()
        {
            return furnList.Count;
        }
    }

    public class FurnData
    {
        public string furnName { get; set; }
        public string familyName { get; set; }
        public string typeName { get; set; }
        public FamilySymbol familySymbol { get; private set; }
        public Document doc { get; set; }

        public FurnData(Document _doc, string _furnName, string _familyName, string _typeName)
        {
            furnName = _furnName;
            familyName = _familyName;
            typeName = _typeName;
            doc = _doc;
            familySymbol = GetFamilySymbol();
        }

        private FamilySymbol GetFamilySymbol()
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Family));

            foreach(Family curFam in collector)
            {
                if(curFam.Name == familyName)
                {
                    ISet<ElementId> famSymbolList = curFam.GetFamilySymbolIds();

                    foreach(ElementId curID in famSymbolList)
                    {
                        FamilySymbol curFS = doc.GetElement(curID) as FamilySymbol;

                        if (curFS.Name == typeName)
                            return curFS;
                    }
                }
            }

            return null;
        }
    }
}
