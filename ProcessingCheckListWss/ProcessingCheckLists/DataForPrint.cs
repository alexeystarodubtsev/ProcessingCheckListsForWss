using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ProcessingCheckListWss.ProcessingCheckLists
{
    class DataForPrint
    {
        public string manager;
        public int qty;
        public int qtyWithoutIncoming = 0;
        public double AVGPercent;
        public TimeSpan duration;
        public TimeSpan AVGduration;
        public Dictionary <string, int> Objections = new Dictionary<string, int>();
        public enum Estimate { qty, AVG, duration, badPoints, AVGDuration, Objection };
        public string BadPoints ="";
        public List<string> tags = new List<string>();
        public DataForPrint(Stage s1, string manager)
        {
            this.manager = manager;
            this.qty = s1.getCountOfCalls();
            this.AVGPercent = s1.getAVGPersent();
            this.duration = s1.getTotalDuration();
            tags.Add("Цена");
            tags.Add("Сроки");
            tags.Add("Оплата");

            tags.Add("Конкуренты");
            foreach (var tag in tags)
            {
                Objections[tag] = s1.calls.Where(c =>  Regex.Match(c.Objections, tag, RegexOptions.IgnoreCase).Success || (tag == "Цена" && Regex.Match(c.Objections, "Дорого", RegexOptions.IgnoreCase).Success)).Count();

            }
            Objections["Другое"] = s1.calls.Where(c => c.Objections != "" && c.Objections.Trim().ToLower() != "нет" && !tags.Where(tag => Regex.Match(c.Objections, tag, RegexOptions.IgnoreCase).Success || (tag == "Цена" && Regex.Match(c.Objections, "Дорого", RegexOptions.IgnoreCase).Success)).Any()).Count();
            Objections["ИТОГО"] = s1.calls.Where(c => c.Objections != "" && c.Objections.Trim().ToLower() != "нет").Count();
            if (qty != 0)
                this.AVGduration = new TimeSpan((long)duration.TotalSeconds / qty * 10000000);
            else
                this.AVGduration = duration;
            var dictPoints = s1.getStatisticOfPoints();
            foreach (var p in dictPoints)
            {
                
                
                int qtyRed = p.Value.Key;
                int qtyAll = p.Value.Value;
                double AVGPerCent = (double)(qtyAll - qtyRed) / qtyAll; 

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
            this.qtyWithoutIncoming = m.getCountOfCallsWithoutIncoming();

            tags.Add("Цена");
            tags.Add("Сроки");
            tags.Add("Оплата");

            tags.Add("Конкуренты");
            foreach (var tag in tags)
            {
                Objections[tag] = m.GetCalls().Where(c => Regex.Match(c.Objections, tag, RegexOptions.IgnoreCase).Success || (tag == "Цена" && Regex.Match(c.Objections, "Дорого", RegexOptions.IgnoreCase).Success)).Count();

            }
            Objections["Другое"] = m.GetCalls().Where(c => c.Objections != "" && c.Objections.Trim().ToLower() != "нет" && !tags.Where(tag => Regex.Match(c.Objections, tag, RegexOptions.IgnoreCase).Success || (tag == "Цена" && Regex.Match(c.Objections, "Дорого", RegexOptions.IgnoreCase).Success)).Any()).Count();
            Objections["ИТОГО"] = m.GetCalls().Where(c => c.Objections != "" && c.Objections.Trim().ToLower() != "нет").Count();
            if (qty != 0)
                this.AVGduration = new TimeSpan((long)duration.TotalSeconds / qty * 10000000);
            else
                this.AVGduration = duration;

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

        public static List<Estimate> getEstimates(bool totalOpt = false)
        {
            List<Estimate> l1 = new List<Estimate>();
            l1.Add(Estimate.AVG);
            l1.Add(Estimate.qty);
            if (totalOpt)
            {
                l1.Add(Estimate.qty);
                
            }
            l1.Add(Estimate.duration);
            l1.Add(Estimate.AVGDuration);
            l1.Add(Estimate.Objection);
            l1.Add(Estimate.badPoints);
            return l1;
        }
    }
}
