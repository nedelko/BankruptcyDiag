using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankruptcyDiagnostics
{
    public class Report
    {
        public double elem_1_1095 { get; set; }
        public double elem_1_1100 { get; set; }
        public double elem_1_1195 { get; set; }
        public double elem_1_1195_first { get; set; }
        public double elem_1_1420 { get; set; }
        public double elem_1_1495 { get; set; }
        public double elem_1_1595 { get; set; }
        public double elem_1_1695 { get; set; }
        public double elem_1_1900 { get; set; }
        public double elem_2_2000 { get; set; }
        public double elem_2_2050 { get; set; }
        public double elem_2_2130 { get; set; }
        public double elem_2_2150 { get; set; }
        public double elem_2_2190 { get; set; }
        public double elem_2_2195 { get; set; }
        public double elem_2_2290 { get; set; }
        public double elem_2_2295 { get; set; }
        public double elem_2_2350 { get; set; }
        public double elem_2_2355 { get; set; }
        public double elem_2_2515 { get; set; }
        public int rep_year { get; set; }
        public Report(double elem1_1095, double elem1_1100, double elem1_1195, double elem1_1195_first, double elem1_1420, double elem1_1495, double elem1_1595, double elem1_1695, double elem1_1900, double elem2_2000, double elem2_2050, double elem2_2130, double elem2_2150, double elem2_2190, double elem2_2195, double elem2_2290, double elem2_2295, double elem2_2350, double elem2_2355, double elem2_2515, int r_year)
        {
            this.elem_1_1095 = elem1_1095;
            this.elem_1_1100 = elem1_1100;
            this.elem_1_1195 = elem1_1195;
            this.elem_1_1195_first = elem1_1195_first;
            this.elem_1_1420 = elem1_1420;
            this.elem_1_1495 = elem1_1495;
            this.elem_1_1595 = elem1_1595;
            this.elem_1_1695 = elem1_1695;
            this.elem_1_1900 = elem1_1900;
            this.elem_2_2000 = elem2_2000;
            this.elem_2_2050 = elem2_2050;
            this.elem_2_2130 = elem2_2130;
            this.elem_2_2150 = elem2_2150;
            this.elem_2_2190 = elem2_2190;
            this.elem_2_2195 = elem2_2195;
            this.elem_2_2290 = elem2_2290;
            this.elem_2_2295 = elem2_2295;
            this.elem_2_2350 = elem2_2350;
            this.elem_2_2355 = elem2_2355;
            this.elem_2_2515 = elem2_2515;
            this.rep_year = r_year;
        }
    }
}
