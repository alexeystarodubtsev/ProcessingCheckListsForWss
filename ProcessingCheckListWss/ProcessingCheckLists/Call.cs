using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessingCheckListWss.ProcessingCheckLists
{
    class Call
    {
        List<Point> points;
        int maxMark;
        TimeSpan duration;
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
        public Call(string client, 
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
                    string ClientLink = "")
        {
            this.maxMark = maxMark;
            this.duration = duration;
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
