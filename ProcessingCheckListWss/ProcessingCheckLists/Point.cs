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
        public string stageForBelfan { get; set; }
        public Point(string name, int mark, bool error)
        {
            this.name = name;
            this.mark = mark;
            this.error = error;
        }
        public Point(string name, int mark, bool error, string stageForBelfan)
        {
            this.name = name;
            this.mark = mark;
            this.error = error;
            this.stageForBelfan = stageForBelfan;
        }

    }
}
