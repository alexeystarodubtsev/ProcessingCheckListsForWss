using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ProcessingCheckListWss.ProcessingCheckLists
{
    class ProcessingBelfanCheckList : Manager
    {
        public ProcessingBelfanCheckList(string filepath, string month) : base(filepath, month)
        {

        }
        public new void Processing()
        {
            XLWorkbook wb = new XLWorkbook(FilePath);
            foreach (var page in wb.Worksheets)
            {
                if (page.Name.ToUpper().Trim() != "СТАТИСТИКА" && page.Name.ToUpper().Trim() != "СВОДНЫЕ" && page.Name.ToUpper().Trim() != "СВОДНАЯ" && page.Name.ToUpper().Trim() != "СТАТИСТИКИ")
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
                            DateTime.TryParse(CellDate.GetValue<string>(), out curDate);
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
                            }
                            CellPoint = CellDate.CellBelow().CellBelow().CellBelow().CellBelow();

                            string DealName = "";


                            string comment = page.Cell(corrRow, CellPoint.Address.ColumnNumber).GetString();
                            bool redComment = page.Cell(corrRow, CellPoint.Address.ColumnNumber).Style.Fill.BackgroundColor
                                                    == XLColor.Red ? true : false;
                            int maxMark;
                            page.Cell(corrRow - 3, CellPoint.Address.ColumnNumber).TryGetValue(out maxMark);
                            if (!CellPoint.TryGetValue<int>(out markOfPoint))
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
                                //int i = 0;
                                //while (page.Cell(CellPoint.Address.RowNumber - i, 1).GetString() == "")
                                //{
                                //    i++;
                                //}
                                //curPoint.stageForBelfan = page.Cell(CellPoint.Address.RowNumber - i, 1).GetString();
                                curPoint.stageForBelfan = CellPoint.Address.RowNumber.ToString();
                                points.Add(curPoint);
                            }
                            CellPoint = CellPoint.CellBelow();
                            while (CellPoint.Address.RowNumber < corrRow - 4)
                            {
                                if (CellPoint.TryGetValue<int>(out markOfPoint))
                                {
                                    CellNamePoint = page.Cell(CellPoint.Address.RowNumber, numColPoint);
                                    bool error = CellPoint.Style.Fill.BackgroundColor == XLColor.Red;
                                    curPoint = new Point(CellNamePoint.GetString(), markOfPoint, error);
                                    //int i = 0;
                                    //while (page.Cell(CellPoint.Address.RowNumber - i,1).GetString() == "")
                                    //{
                                    //    i++;
                                    //}
                                    //curPoint.stageForBelfan = page.Cell(CellPoint.Address.RowNumber - i, 1).GetString();
                                    curPoint.stageForBelfan = CellPoint.Address.RowNumber.ToString();
                                    points.Add(curPoint);
                                }
                                else
                                {
                                    string answer = CellPoint.GetString().ToLower();
                                    if (answer == "нет" || answer == "да")
                                    {
                                        CellNamePoint = page.Cell(CellPoint.Address.RowNumber, numColPoint);
                                        curPoint = new Point(CellNamePoint.GetString(), answer == "нет" ? 0 : 1, answer == "нет" ? true : false, true);
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
                            if (curDate > new DateTime(2020, 5, 6))
                            {
                                Objections = page.Cell(corrRow + 2, CellPoint.Address.ColumnNumber).GetString();
                                howProcessObj = page.Cell(corrRow + 4, CellPoint.Address.ColumnNumber).GetString();
                                DealState = page.Cell(corrRow + 5, CellPoint.Address.ColumnNumber).GetString();

                                DateOfNext = page.Cell(corrRow + 6, CellPoint.Address.ColumnNumber).GetString();
                                doneObj = page.Cell(corrRow + 3, CellPoint.Address.ColumnNumber).GetString();
                            }
                            DateTime ddateNext;
                            if (DateOfNext != "")
                            {
                                if (DateTime.TryParse(DateOfNext, out ddateNext)) 
                                    DateOfNext = ddateNext.ToString("dd.MM.yyyy");
                            }
                            if (Regex.Match(phoneNumber.ToUpper(), "ВХОДЯЩ").Success)
                                outgoing = false;
                            bool greenComment = page.Cell(corrRow, CellPoint.Address.ColumnNumber).Style.Fill.BackgroundColor
                                                    == XLColor.Lime ? true : false;
                            if (points.Count > 0)
                                calls.Add(new Call(phoneNumber, maxMark, duration, comment, DealName, points, redComment, curDate, outgoing, greenComment, Objections, howProcessObj, DealState,link, DateOfNext, doneObj));
                        }
                        CellDate = CellDate.CellRight();
                    }
                    stages.Add(new Stage(page.Name.Trim(), calls));

                }
            }
        }

        public new Dictionary<string, KeyValuePair<int, int>> getStatisticOfPoints()
        {
            Dictionary<string, KeyValuePair<int, int>> dict = new Dictionary<string, KeyValuePair<int, int>>(); //Пункт, число красных, число всего
            foreach (var call in GetCalls())
            {

                foreach (var point in call.getPoints())
                {

                    int red = point.error ? 1 : 0;
                    if (!dict.ContainsKey(point.name + point.stageForBelfan))
                        dict[point.name + point.stageForBelfan] = new KeyValuePair<int, int>(red, 1);
                    else
                    {
                        KeyValuePair<int, int> old = dict[point.name + point.stageForBelfan];
                        dict[point.name + point.stageForBelfan] = new KeyValuePair<int, int>(old.Key + red, old.Value + 1);
                    }

                }

            }

            return dict;
        }
    }
}
