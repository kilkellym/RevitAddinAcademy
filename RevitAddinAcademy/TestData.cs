using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAddinAcademy
{
    internal class TestData
    {
        public string Data1 { get; set; }
        public string Data2 { get; set; }
        public int Data3 { get; set; }
        public string Combo { get; set; }

        public TestData(string d1, string d2, int d3)
        {
            Data1 = d1;
            Data2 = d2;
            Data3 = d3;

            Combo = d1 + " " + d2 + " " + d3.ToString();
        }
    }
}
