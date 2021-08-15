using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ProcessingCheckListWss.ProcessingCheckLists
{
    class Call
    {
        List<Point> points;
        int maxMark;
        TimeSpan duration;
        public XLColor colorCELL { get; }
        public double conversion { get; }
        public string comment { get; }
        public bool redComment { get; }
        public bool greenComment { get; }
        public string client { get; }
        string DealName;
        public DateTime dateOfCall { get; }
        public bool outgoing { get; }

        public bool unconvinienttalk;
        public string Objections { get; set; }
        public string howProcessObj { get; set; }
        public string DealState {get; set;}
        public string ClientLink { get; set; }
        public string DateOfNext { get; set; }
        public string doneObjection { get; set; }

        public Call(
                    XLColor colorCELL, 
                    string client,
                    int maxMark,
                    TimeSpan duration,
                    string comment,                  
                    string DealName, 
                    List<Point> points, bool redComment,
                    DateTime dateOfCall,
                    bool outgoing,
                    bool greencomment = false,
                    string Objections = "",                   
                    string howProcessObj = "",
                    string DealState = "",
                    string ClientLink = "",
                    string DateOfNext = "",
                    string doneObjection = "",
                    double conversion = 0)
        {
            this.maxMark = maxMark;
            this.duration = duration;
            this.conversion = conversion;
            this.comment = comment;
            this.client = client;
            this.DealName = DealName;
            this.points = points;
            this.redComment = redComment;
            this.dateOfCall = dateOfCall;
            this.outgoing = outgoing;
            this.greenComment = greencomment;
            this.Objections = Objections;
            this.howProcessObj = howProcessObj;
            this.DealState = DealState;
            this.ClientLink = ClientLink;
            this.DateOfNext = DateOfNext;
            this.doneObjection = doneObjection;
            
        }
        public XLColor GETcolorCELL()
        {
            return colorCELL;
        }

        public double getAVGConversion()
        {
            return conversion * 1000;
        }
        public double getAVGPersent()
        {
            double mark = 0;
            foreach (Point p in points)
            {
                mark += p.mark;
            }
            return mark / maxMark;
        }
        public TimeSpan getDuration()
        {
            return duration;
        }
        public List <Point> getPoints()
        {          
            return points;
        }
        
    }
}
