using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankruptcyDiagnostics
{
    public class ReportCollection : IEnumerable
    {
        private ArrayList repCollection = new ArrayList();
        public Report GetReport(int pos)
        {
            return (Report)repCollection[pos];
        }
        public void AddReport(Report rep)
        {
            int sortingIndex = 0;
            if (repCollection.Count>0)
            {
                for (int i=0; i < repCollection.Count; i++)
                {
                    if(GetReport(i).rep_year < rep.rep_year)
                    {
                        sortingIndex = i+1;
                    } 
                }
                repCollection.Insert(sortingIndex, rep);
            }
            else
            {
                repCollection.Add(rep);
            }
        }
        public void ClearReports()
        {
            repCollection.Clear();
        }
        public int Count
        {
            get { return repCollection.Count; }
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return repCollection.GetEnumerator();
        }
    }
}
