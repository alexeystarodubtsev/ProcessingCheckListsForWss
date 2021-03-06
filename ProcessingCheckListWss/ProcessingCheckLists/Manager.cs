﻿using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ProcessingCheckListWss.ProcessingCheckLists
{
    class Manager
    {
        public string Name { get; }
        public string FilePath { get; set; }
        protected List<Stage> stages = new List<Stage>();
        public string month;
        
        public Manager (string filepath, string month)
        {
            var Match = Regex.Match(filepath, @"(\w+).xlsx");
            //this.Name = Match.Groups[1].Value;
            this.Name = Regex.Match(Path.GetFileName(filepath), @"(\w+)").Groups[1].Value;
            this.month = month;
            FilePath = filepath;
        
            //Processing();
        }
        public List<Stage> getStages()
        {
            return stages;
        }
        
        public int getCountOfCalls()
        {
            int calls = 0;
            foreach (var stage in stages)
            {
                calls += stage.getCountOfCalls();
            }
            return calls;
        }
        public int getCountOfCallsWithoutIncoming()
        {
            int qtycalls = 0;
            foreach (var stage in stages.Where(s => !Regex.Match(s.name.ToLower(), "было не удобно говорить").Success))
            {
                qtycalls += stage.calls.Where(c => c.outgoing).Count();
            }
            return qtycalls;
        }
        public int getCountOfCalls(DateTime firstDate)
        {
            return GetCalls().Where(c => c.dateOfCall >= firstDate).Count(); ;
        }
        public int getCountOfCalls(DateTime firstDate, DateTime LastDate)
        {
            return GetCalls().Where(c => c.dateOfCall >= firstDate && c.dateOfCall <= LastDate).Count(); ;
        }
        public List<Call> getBadComments(DateTime firstDate,DateTime lastDate)
        {
            List<Call> BadComments = new List<Call>();
            string comment = "";
            foreach (var call in GetCalls())
            {
                if (call.redComment && call.dateOfCall >= firstDate && call.dateOfCall <= lastDate)
                    //comment += call.comment + " ("+ call.client + " " + call.dateOfCall.ToString("dd.MM") +"); ";
                    BadComments.Add(call);
            }
            //comment = comment.TrimEnd(' ').TrimEnd(';');
            return BadComments;
        }
        public List<Call> getgoodComments(DateTime firstDate, DateTime lastDate)
        {
            List<Call> goodComments = new List<Call>();
            string comment = "";
            foreach (var call in GetCalls())
            {
                if (call.greenComment && call.dateOfCall >= firstDate && call.dateOfCall < lastDate)
                    //comment += call.comment + " (" + call.client + " " + call.dateOfCall.ToString("dd.MM") + "); ";
                    goodComments.Add(call);
            }
            //comment = comment.TrimEnd(' ').TrimEnd(';');
            return goodComments;
        }


        public List<string> getBadPoints(DateTime firstDate,DateTime lastDate)
        {
            List<string> points = new List<string>();
            //string points = "";
            var dictPoints = getStatisticOfPoints(firstDate,lastDate);
            foreach (var point in dictPoints.Keys)
            {
                if ((double)(dictPoints[point].Value - dictPoints[point].Key) / dictPoints[point].Value < 1)
                {
                    HashSet<string> stagesInPoint = new HashSet<string>();

                    foreach (var stage in stages)
                    {
                        foreach(var call in stage.calls)
                        {
                            if (call.getPoints().Any(p => p.name == point))
                            {
                                stagesInPoint.Add(stage.name);
                                break;
                            }
                        }
                    }
                    points.Add(point + " (" + ((double)(dictPoints[point].Value - dictPoints[point].Key) / dictPoints[point].Value).ToString("P1") + ") (этапы: " + String.Join(", ", stagesInPoint.ToArray()) + ")");
                }

            }
            //points = points.TrimEnd(' ').TrimEnd(';');
            return points;
        }

        Dictionary<string, KeyValuePair<int, int>> getStatisticOfPoints(DateTime firstDate, DateTime lastDate)
        {
            Dictionary<string, KeyValuePair<int, int>> dict = new Dictionary<string, KeyValuePair<int, int>>(); //Пункт, число красных, число всего
            foreach (var call in GetCalls())
            {
                if (call.dateOfCall >= firstDate && call.dateOfCall < lastDate)
                {
                    foreach (var point in call.getPoints())
                    {

                        int red = point.error ? 1 : 0;
                        if (!dict.ContainsKey(point.name))
                            dict[point.name] = new KeyValuePair<int, int>(red, 1);
                        else
                        {
                            KeyValuePair<int, int> old = dict[point.name];
                            dict[point.name] = new KeyValuePair<int, int>(old.Key + red, old.Value + 1);
                        }

                    }
                }
            }

            return dict;
        }
        public Dictionary<string, KeyValuePair<int, int>> getStatisticOfPoints()
        {
            Dictionary<string, KeyValuePair<int, int>> dict = new Dictionary<string, KeyValuePair<int, int>>(); //Пункт, число красных, число всего
            foreach (var call in GetCalls())
            {
               
                foreach (var point in call.getPoints())
                {

                    int red = point.error ? 1 : 0;
                    if (!dict.ContainsKey(point.name))
                        dict[point.name] = new KeyValuePair<int, int>(red, 1);
                    else
                    {
                        KeyValuePair<int, int> old = dict[point.name];
                        dict[point.name] = new KeyValuePair<int, int>(old.Key + red, old.Value + 1);
                    }

                }
                
            }

            return dict;
        }

        public TimeSpan getTotalDuration ()
        {
            TimeSpan t1 = new TimeSpan();
            foreach(var s in stages)
            {
                t1 = t1.Add(s.getTotalDuration());
            }
            return t1;
        }
        public Dictionary<string,int> getCountOfCallsByStages()
        {
            Dictionary<string, int>  d1 = new Dictionary<string, int>();
            foreach (var stage in stages)
            {
                d1[stage.name] = stage.getCountOfCalls();
            }
            return d1;
        }
        public double getAVGPersent()
        {
           
             
            double SumPers = 0;
            foreach(var call in GetCalls())
            {
                SumPers += call.getAVGPersent();

            }
            int qty = getCountOfCalls();
            return qty > 0 ? SumPers / qty : -1;
        }
        public double getAVGPersent(DateTime firstDate)
        {
            double SumPers = 0;
            int qty = 0;
            foreach (var call in GetCalls())
            {
                if (call.dateOfCall >= firstDate)
                {
                    SumPers += call.getAVGPersent();
                    qty++;
                }
            }
            
            return qty > 0 ? SumPers / qty : -1;
        }
        public double getAVGPersent(DateTime firstDate, DateTime lastDate)
        {
            double SumPers = 0;
            int qty = 0;
            foreach (var call in GetCalls())
            {
                if (call.dateOfCall >= firstDate && call.dateOfCall <= lastDate)
                {
                    SumPers += call.getAVGPersent();
                    qty++;
                }
            }

            return qty > 0 ? SumPers / qty : -1;
        }

        public Dictionary<string, DataForPrint> getDataByStage()
        {
            Dictionary<string, DataForPrint> pages = new Dictionary<string, DataForPrint>();
            foreach (Stage s1 in stages)
            {
                pages[s1.name] = new DataForPrint(s1, Name);
            }
            return pages;
        }
        public List<Call> GetCalls()
        {
            List<Call> calls = new List<Call>();
            foreach (var s in stages)
            {
                foreach (var call in s.calls)
                {
                    calls.Add(call);
                }
            }
            return calls;
        }
        public string getWorseCall(DateTime firstDate)
        {
            try
            {
                var WorseCall = GetCalls().Where(c => c.dateOfCall >= firstDate).First();
                foreach (var call in GetCalls().Where(c => c.dateOfCall >= firstDate))
                {
                    if (call.getAVGPersent() < WorseCall.getAVGPersent() && (call.redComment || !WorseCall.redComment))
                        WorseCall = call;
                }
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                return WorseCall.getAVGPersent() == 1 ? "" : WorseCall.client + " (" +  WorseCall.getAVGPersent().ToString("P1", CultureInfo.InvariantCulture) + ")";
            }
            catch (System.InvalidOperationException)
            {
                return "";
            }
            
        }

        public void Processing()
        {
            XLWorkbook wb = new XLWorkbook(FilePath);
            
            foreach (var page in wb.Worksheets)
            {
                if (page.Name.ToUpper().Trim() != "СТАТИСТИКА" && page.Name.ToUpper().Trim() != "СВОДНАЯ" && page.Name.ToUpper().Trim() != "СТАТИСТИКИ")
                {
                    const int numColPoint = 4;
                    IXLCell CellDate = page.Cell(1, numColPoint + 1);
                    DateTime curDate;
                    DateTime.TryParse(CellDate.GetValue<string>(), out curDate);
                    Regex rComment = new Regex(@"КОРРЕКЦИИ");
                    int corrRow = 5;
                    Match Mcomment = rComment.Match(page.Cell(corrRow, 1).GetString().ToUpper());
                    while (!Mcomment.Success)
                    {
                        corrRow++;
                        Mcomment = rComment.Match(page.Cell(corrRow, 1).GetString().ToUpper());
                    }
                    List<Call> calls = new List<Call>();
                    while (!(CellDate.CellBelow().IsEmpty() && CellDate.CellBelow().CellRight().IsEmpty() && CellDate.CellBelow().CellBelow().IsEmpty() && CellDate.CellBelow().CellBelow().CellRight().IsEmpty()))
                    {

                        if (CellDate.GetValue<string>() != "")
                        {
                            if (CellDate.DataType == XLDataType.DateTime)
                            {
                                curDate = CellDate.GetDateTime();
                            }
                            else
                            {
                                if (!DateTime.TryParse(CellDate.GetString(), new CultureInfo("ru-RU"), DateTimeStyles.None, out curDate))
                                {
                                    DateTime.TryParse(CellDate.GetString(), new CultureInfo("en-US"), DateTimeStyles.None, out curDate);
                                }
                            }
                        }
                        string phoneNumber = CellDate.CellBelow().GetValue<string>();
                        var phoneCell = CellDate.CellBelow();
                        if (phoneNumber == "")
                        {
                            phoneNumber = CellDate.CellBelow().CellBelow().GetValue<string>();
                            phoneCell = CellDate.CellBelow().CellBelow();
                        }
                        
                        if (phoneNumber != "")
                        {
                            string link = "";
                            if (phoneCell.HasHyperlink)
                                link = phoneCell.Hyperlink.ExternalAddress.AbsoluteUri;

                            TimeSpan duration;
                            
                            
                            TimeSpan wrongtime1 = new TimeSpan(1, 0, 0, 0);
                            TimeSpan wrongtime2 = new TimeSpan();
                            IXLCell CellPoint = CellDate.CellBelow().CellBelow().CellBelow();
                            if (CellPoint.DataType == XLDataType.DateTime)
                                CellPoint.DataType = XLDataType.TimeSpan;
                            
                            TimeSpan.TryParse(CellPoint.GetString(), out duration);
                            IXLCell CellNamePoint;
                            List<Point> points = new List<Point>();
                            Point curPoint;
                            int markOfPoint;
                            if (wrongtime1 <= duration || duration == wrongtime2)
                            {
                                duration = wrongtime2;
                                if (CellPoint.TryGetValue<int>(out markOfPoint))
                                {
                                    CellNamePoint = page.Cell(CellPoint.Address.RowNumber, numColPoint);
                                    bool error = CellPoint.Style.Fill.BackgroundColor == XLColor.Red;
                                    curPoint = new Point(CellNamePoint.GetString(), markOfPoint, error);
                                    points.Add(curPoint);
                                }
                                else
                                {
                                    string answer = CellPoint.GetString().ToLower();
                                    if (answer == "нет" || answer == "да")
                                    {
                                        CellNamePoint = page.Cell(CellPoint.Address.RowNumber, numColPoint);
                                        bool error = CellPoint.Style.Fill.BackgroundColor == XLColor.Red;
                                        curPoint = new Point(CellNamePoint.GetString(), answer == "нет" ? 0 : 1, error, true);
                                        points.Add(curPoint);
                                    }
                                }
                            }
                            CellPoint = CellDate.CellBelow().CellBelow().CellBelow().CellBelow();
                            
                            string DealName = "";
                            
                            
                            string comment = page.Cell(corrRow, CellPoint.Address.ColumnNumber).GetString();
                            bool redComment = page.Cell(corrRow, CellPoint.Address.ColumnNumber).Style.Fill.BackgroundColor 
                                                    == XLColor.Red ? true : false;
                            var Color = page.Cell(corrRow, CellPoint.Address.ColumnNumber).Style.Fill.BackgroundColor;
                            bool greenComment = page.Cell(corrRow, CellPoint.Address.ColumnNumber).Style.Fill.BackgroundColor
                                                    == XLColor.Lime ? true : false;
                            int maxMark;
                            page.Cell(corrRow - 3, CellPoint.Address.ColumnNumber).TryGetValue(out maxMark);
                            if (!CellPoint.TryGetValue<int>(out markOfPoint) )
                            {
                                if (CellPoint.GetString() != "")
                                {
                                    DealName = CellPoint.GetString();
                                }
                            }
                            else
                            {
                                CellNamePoint = page.Cell(CellPoint.Address.RowNumber, numColPoint);
                                bool error = CellPoint.Style.Fill.BackgroundColor == XLColor.Red;
                                curPoint = new Point(CellNamePoint.GetString(), markOfPoint, error);
                                points.Add(curPoint);
                            }
                            CellPoint = CellPoint.CellBelow();
                            while (CellPoint.Address.RowNumber < corrRow - 4)
                            {
                                if (CellPoint.TryGetValue<int>(out markOfPoint))
                                {
                                    CellNamePoint = page.Cell(CellPoint.Address.RowNumber, numColPoint);
                                    
                                    int weightPoint = CellNamePoint.CellLeft().CellLeft().GetValue<int>();
                                    bool error = CellPoint.Style.Fill.BackgroundColor == XLColor.Red;
                                    curPoint = new Point(CellNamePoint.GetString(), markOfPoint, error, CellPoint.Address.RowNumber.ToString());
                                    //if (notTakenPoint(CellNamePoint.GetString()))
                                    //    maxMark -= weightPoint;

                                    //else
                                        points.Add(curPoint);
                                }
                                else
                                {
                                    string answer = CellPoint.GetString().ToLower();
                                    if (answer == "нет" || answer == "да")
                                    {
                                        CellNamePoint = page.Cell(CellPoint.Address.RowNumber, numColPoint);
                                        curPoint = new Point(CellNamePoint.GetString(), answer == "нет" ? 0 : 1, answer == "нет" ? true : false, true);
                                        points.Add(curPoint);
                                    }
                                }
                                CellPoint = CellPoint.CellBelow();
                            }
                            bool outgoing = true;
                            string Objections = "";
                            string howProcessObj = "";
                            string DealState = "";
                            string DateOfNext = "";
                            string doneObj = "";
                            if (curDate > new DateTime(2020,5,6))
                            {
                                Objections = page.Cell(corrRow + 2, CellPoint.Address.ColumnNumber).GetString();
                                howProcessObj = page.Cell(corrRow + 4, CellPoint.Address.ColumnNumber).GetString();
                                DealState = page.Cell(corrRow + 5, CellPoint.Address.ColumnNumber).GetString();
                                DateOfNext = page.Cell(corrRow + 6, CellPoint.Address.ColumnNumber).GetString();
                                DateTime ddateNext;
                                if (DateOfNext != "")
                                {
                                    if (DateTime.TryParse(DateOfNext, out ddateNext))
                                        DateOfNext = ddateNext.ToString("dd.MM.yyyy");
                                }
                                doneObj = page.Cell(corrRow + 3, CellPoint.Address.ColumnNumber).GetString();
                            }
                            if (Regex.Match(page.Name.ToUpper(), "ВХОДЯЩ").Success)
                                outgoing = false;
                            if (points.Count > 0)
                              calls.Add(new Call(phoneNumber, maxMark, duration, comment, DealName, points, redComment, curDate,outgoing, greenComment,Objections,howProcessObj,DealState,link,DateOfNext,doneObj));
                        }
                        CellDate = CellDate.CellRight();
                    }
                    stages.Add(new Stage(page.Name, calls));
                    
                }
            }
            
        }
        public void getInformationPerDay(DateTime firstDate, DateTime lastDate)
        {
            Console.WriteLine(Name);
            while (firstDate < lastDate)
            {
                double val = GetCalls().Where(c => c.dateOfCall == firstDate).Sum(c => c.getAVGPersent());
                var qty =  GetCalls().Where(c => c.dateOfCall == firstDate).Count();
                val /= qty;
                Console.WriteLine(firstDate.ToString("dd.MM") + " " + val.ToString("P2") + " " + qty);
                firstDate = firstDate.AddDays(1);

            }
        }
        public DateTime getLastDate()
        {
            var calls = GetCalls();
            if (calls.Count == 0)
                return new DateTime(1, 1, 1);
            return calls.Max(c => c.dateOfCall).AddDays(1);
        }
        bool notTakenPoint(string point)
        {
            List<string> notCountPoints = new List<string>();
            notCountPoints.Add("Нет слов паразитов");
            notCountPoints.Add("Нет пауз");
            notCountPoints.Add("Нет затяжных, например ээээ и т.п.");
            notCountPoints.Add("Не издает лишних междометий");
            notCountPoints.Add("Нет длительных пауз");
            notCountPoints.Add("Нет слов - паразитов.");
            if (notCountPoints.Contains(point))
                return true;
            else
                return false;

        }
        public void Concat(Manager m)
        {
            if (m != null)
            {
                List<Stage> newStages = new List<Stage>(); 
                foreach (Stage s1 in m.stages)
                {
                    try
                    {
                        var curStage = stages.Where(s => s.name == s1.name).First();
                        foreach (var call in s1.calls)
                            curStage.calls.Add(call);
                    }
                    catch(InvalidOperationException) {
                        if (stages.Where(s => s.name == s1.name).Count() == 0)
                        {
                            newStages.Add(s1);
                        }
                    }
                }
                try
                {
                    stages.AddRange(newStages);
                }
                catch(Exception) { }
            }
        }


    }
}
