using System;
using System.Collections.Generic;
using System.Globalization;
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
        public enum Estimate { qty, AVG, duration, badPoints }
        public string BadPoints ="";
        public DataForPrint(Stage s1, string manager)
        {
            this.manager = manager;
            this.qty = s1.getCountOfCalls();
            this.AVGPercent = s1.getAVGPersent();
            this.duration = s1.getTotalDuration();
            var dictPoints = s1.getStatisticOfPoints();
            foreach (var p in dictPoints)
            {
                
                
                int qtyRed = p.Value.Key;
                int qtyAll = p.Value.Value;
                double AVGPerCent = (double)(qtyAll - qtyRed) / qtyAll; ;
                if (AVGPerCent < 0.8)
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    
                    this.BadPoints += p.Key +  " (" + AVGPerCent.ToString("P1", CultureInfo.InvariantCulture) + ")" + ";\n";

                }

            }

            this.BadPoints.Trim('\n').Trim(';');

        }
        public DataForPrint(Manager m)
        {
            this.manager = m.Name;
            this.qty = m.getCountOfCalls();
            this.AVGPercent = m.getAVGPersent();
            this.duration = m.getTotalDuration();
            var dictPoints = m.getStatisticOfPoints();
            foreach (var p in dictPoints)
            {

                int qtyRed = p.Value.Key;
                int qtyAll = p.Value.Value;
                double AVGPerCent = (double)(qtyAll - qtyRed) / qtyAll; ;
                if (AVGPerCent < 0.8)
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                    this.BadPoints += p.Key + " (" + AVGPerCent.ToString("P1", CultureInfo.InvariantCulture) + ")" + ";\n";

                }
            }
                this.BadPoints.Trim('\n').Trim(';');
        }

        public static List<Estimate> getEstimates()
        {
            List<Estimate> l1 = new List<Estimate>();
            l1.Add(Estimate.AVG);
            l1.Add(Estimate.qty);
            l1.Add(Estimate.duration);
           // l1.Add(Estimate.badPoints);
            return l1;
        }
    }
}
