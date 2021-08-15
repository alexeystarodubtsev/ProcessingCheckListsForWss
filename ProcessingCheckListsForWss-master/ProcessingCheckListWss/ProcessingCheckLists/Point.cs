using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessingCheckListWss.ProcessingCheckLists
{
    class Point
    {
        public string name { get; }
        public int mark { get; }
        public bool error { get; }
        public bool noStatistic {get; }
        public string stageForBelfan { get; set; }
        public XLColor ColorForRNR { get; set; }
        public Point(string name, int mark, bool error, string stageForBelfan)
        {
            this.name = name;
            this.mark = mark;
            this.error = error;
            this.stageForBelfan = stageForBelfan;
        }
        public Point(string name, int mark, bool error, bool noStatistic = false)
        {
            this.name = name;
            this.mark = mark;
            this.error = error;
            this.noStatistic = noStatistic;
        }

    }
}
