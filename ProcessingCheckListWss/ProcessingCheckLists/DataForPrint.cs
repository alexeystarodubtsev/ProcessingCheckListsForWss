using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessingCheckListWss.ProcessingCheckLists
{
    class DataForPrint
    {
        public string manager;
        public int qty;
        public double AVGPercent;
        public TimeSpan duration;
        public enum Estimate { qty, AVG, duration } 
        public DataForPrint(Stage s1, string manager)
        {
            this.manager = manager;
            this.qty = s1.getCountOfCalls();
            this.AVGPercent = s1.getAVGPersent();
            this.duration = s1.getTotalDuration();
        }
        public DataForPrint(Manager m)
        {
            this.manager = m.Name;
            this.qty = m.getCountOfCalls();
            this.AVGPercent = m.getAVGPersent();
            this.duration = m.getTotalDuration();
        }

        public static List<Estimate> getEstimates()
        {
            List<Estimate> l1 = new List<Estimate>();
            l1.Add(Estimate.AVG);
            l1.Add(Estimate.qty);
            l1.Add(Estimate.duration);
            return l1;
        }
    }
}
