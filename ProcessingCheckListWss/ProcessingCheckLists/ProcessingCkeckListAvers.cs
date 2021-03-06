﻿using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ProcessingCheckListWss.ProcessingCheckLists
{
    class ProcessingCkeckListAvers : Manager
    {
        public ProcessingCkeckListAvers(string filepath, string month) : base(filepath, month)
        {

        }
        public new void Processing()
        {
            XLWorkbook wb = new XLWorkbook(FilePath);
            foreach (var page in wb.Worksheets)
            {
                if (page.Name.ToUpper().Trim() != "СТАТИСТИКА" && page.Name.ToUpper().Trim() != "СВОДНАЯ")
                {
                    const int numColPoint = 4;
                    IXLCell CellDate = page.Cell(2, numColPoint + 1);
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
                    while (!(CellDate.CellBelow().IsEmpty() && CellDate.CellBelow().CellRight().IsEmpty()))
                    {
                        if (CellDate.GetValue<string>() != "")
                        {
                            DateTime.TryParse(CellDate.GetValue<string>(), out curDate);
                        }
                        string phoneNumber = CellDate.CellBelow().GetValue<string>();
                        if (phoneNumber != "")
                        {
                            TimeSpan duration;



                            IXLCell CellPoint = CellDate.CellBelow().CellBelow();
                            if (CellPoint.DataType == XLDataType.DateTime)
                                CellPoint.DataType = XLDataType.TimeSpan;

                            TimeSpan.TryParse(CellPoint.GetString(), out duration);
                            IXLCell CellNamePoint;
                            List<Point> points = new List<Point>();
                            Point curPoint;
                            int markOfPoint;
                            CellPoint = CellDate.CellBelow().CellBelow().CellBelow();

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
                                    points.Add(curPoint);
                                }
                                CellPoint = CellPoint.CellBelow();
                            }
                            bool outgoing = true;
                            if (Regex.Match(page.Name.ToUpper(), "ВХОДЯЩ").Success)
                                outgoing = false;
                            if (points.Count > 0)
                                calls.Add(new Call(phoneNumber, maxMark, duration, comment, DealName, points, redComment, curDate, outgoing));
                        }
                        CellDate = CellDate.CellRight();
                    }
                    stages.Add(new Stage(page.Name, calls));

                }
            }
        }
    }
}
