using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessingCheckListWss.ProcessingCheckLists
{
    class Stage
    {
        public string name { get; }
        public List<Call> calls { get; }
        public Stage(string name, List<Call> calls)
        {
            this.name = name;
            this.calls = calls;
        }
        public int getCountOfCalls()
        {
            return calls.Count();
        }
        public double getAVGPersent()
        {
            double SumPers = 0;
            foreach (Call call in calls)
            {
                SumPers += call.getAVGPersent();
            }
            return calls.Count > 0 ? SumPers / calls.Count : -1;
        }
        public TimeSpan getTotalDuration()
        {
            TimeSpan t1 = new TimeSpan();
            foreach (var call in calls)
            {
                TimeSpan durationCall = call.getDuration();
                t1 = t1.Add(durationCall);
            }
            return t1;
        }
        public Dictionary<string, KeyValuePair<int, int>> getStatisticOfPoints (bool Belfan = false)
        {
            Dictionary<string, KeyValuePair<int,int>> dict = new Dictionary<string, KeyValuePair<int, int>>(); //Пункт, число красных, число всего
            foreach (var call in calls)
            {
                foreach (var point in call.getPoints())
                {
                    int red = point.error ? 1 : 0;
                    if (!dict.ContainsKey(point.name + (Belfan ? point.stageForBelfan : "")))
                        dict[point.name + (Belfan ? point.stageForBelfan : "")] = new KeyValuePair<int, int>(red, 1);
                    else
                    {
                        KeyValuePair<int, int> old = dict[point.name + (Belfan ? point.stageForBelfan : "")];
                        dict[point.name + (Belfan ? point.stageForBelfan : "")] = new KeyValuePair<int, int>(old.Key + red, old.Value + 1);
                    }

                }
            }

            return dict;
        }
            
    }
}
