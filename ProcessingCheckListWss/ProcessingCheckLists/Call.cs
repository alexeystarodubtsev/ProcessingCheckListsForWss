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
        public string client { get; }
        string DealName;
        public DateTime dateOfCall { get; }
        public Call(string client, 
                    int maxMark, 
                    TimeSpan duration,
                    string comment,
                    string DealName, 
                    List<Point> points, bool redComment,
                    DateTime dateOfCall)
        {
            this.maxMark = maxMark;
            this.duration = duration;
            this.comment = comment;
            this.client = client;
            this.DealName = DealName;
            this.points = points;
            this.redComment = redComment;
            this.dateOfCall = dateOfCall;
            
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
