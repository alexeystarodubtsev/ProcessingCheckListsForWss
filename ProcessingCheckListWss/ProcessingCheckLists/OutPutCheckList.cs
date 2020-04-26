using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ProcessingCheckListWss.ProcessingCheckLists
{
    class OutPutCheckList
    {
        
        public  static XLWorkbook FillAnalyticOfPoint(Manager m, Manager PreLastMonthManager)
        {
            XLWorkbook wb = new XLWorkbook(m.FilePath);

            foreach (var stage in m.getStages())
            {
                var dictPoints = stage.getStatisticOfPoints();
                var page = wb.Worksheet(stage.name);
                var table = page.RangeUsed();
                int shift = 2;
                if (table.LastColumn().Cell(6).CellRight().Style.Fill.BackgroundColor != XLColor.NoColor)
                    shift++;
                int curCol = table.LastColumn().ColumnNumber() + shift;
                const int numColPoint = 4;
                IXLCell CellStartCaption = page.Cell(2, curCol);
                if (dictPoints.Count > 0)
                {
                    CellStartCaption.Value = "Количество некорректных пунктов";
                    CellStartCaption.WorksheetColumn().Width = 11;
                    CellStartCaption.CellRight().Value = "Количество прохождений пункта";
                    CellStartCaption.CellRight().WorksheetColumn().Width = 11;
                    CellStartCaption.CellRight().CellRight().Value = "Количество корректных пунктов";
                    CellStartCaption.CellRight().CellRight().WorksheetColumn().Width = 11;
                    CellStartCaption.CellRight().CellRight().CellRight().Value = "% Выполнения";
                    CellStartCaption.CellRight().CellRight().CellRight().WorksheetColumn().Width = 11;
                   
                    var lastCellCaption = CellStartCaption.CellRight().CellRight().CellRight();
                    Dictionary<string, KeyValuePair<int, int>> dictPointsPreLastMonth = new Dictionary<string, KeyValuePair<int, int>>();
                    var prelastMonthCell = CellStartCaption.CellRight().CellRight().CellRight().CellRight();
                    var howChangeCell = prelastMonthCell.CellRight();
                    var totalCompareLastMonthsCell = howChangeCell.CellRight().CellBelow();
                    var totalCompareRange = page.Range(totalCompareLastMonthsCell, totalCompareLastMonthsCell.CellRight().CellRight());
                    var QtyWorseCell = totalCompareLastMonthsCell.CellBelow();
                    var QtyBetterCell = QtyWorseCell.CellRight();
                    var QtyNoChangeCell = QtyBetterCell.CellRight();
                    var totalMonthCell = QtyWorseCell.CellBelow().CellBelow().CellBelow();
                    var totalAVGLastMonthCell = totalMonthCell.CellBelow();
                    var totalAVGPreLastMonthCell = totalAVGLastMonthCell.CellBelow();
                    if (PreLastMonthManager != null && PreLastMonthManager.Name == m.Name && PreLastMonthManager.getStages().Exists(s => s.name == stage.name))
                    {
                        
                        prelastMonthCell.Value = PreLastMonthManager.month;
                        var preLastMonthCurStage = PreLastMonthManager.getStages().Where(s => s.name == stage.name).First();
                        dictPointsPreLastMonth = preLastMonthCurStage.getStatisticOfPoints();
                        
                        howChangeCell.Value = "Как изменилось по сравнению с прошлым месяцем";
                        howChangeCell.WorksheetColumn().Width = 18;
                        lastCellCaption = howChangeCell;
                        
                        totalCompareLastMonthsCell.Value = "Итоги сравнения с прошлым месяцем";
                        
                        totalCompareRange.Merge();
                        
                        QtyWorseCell.Value = "Всего Ухудшилось";
                        
                        QtyBetterCell.Value = "Всего Улучшилось";
                        
                        QtyNoChangeCell.Value = "Без изменения";
                        QtyWorseCell = QtyWorseCell.CellBelow();
                        QtyBetterCell = QtyBetterCell.CellBelow();
                        QtyNoChangeCell = QtyNoChangeCell.CellBelow();
                        totalCompareRange = page.Range(totalCompareLastMonthsCell, QtyNoChangeCell);
                        Range(ref totalCompareRange);
                        totalMonthCell.Value = "Месяц";
                        totalMonthCell.CellRight().Value = "Средний процент выполнения пунктов";
                        totalAVGLastMonthCell.Value = m.month;
                        totalAVGPreLastMonthCell.Value = PreLastMonthManager.month;
                        totalAVGLastMonthCell = totalAVGLastMonthCell.CellRight();
                        totalAVGPreLastMonthCell = totalAVGPreLastMonthCell.CellRight();
                        var totalrng = page.Range(totalMonthCell, totalAVGPreLastMonthCell);
                        Range(ref totalrng);
                        page.Columns(QtyWorseCell.Address.ColumnNumber, QtyNoChangeCell.Address.ColumnNumber).AdjustToContents();

                    }
                    var Caption = page.Range(CellStartCaption, lastCellCaption);
                    
                    Caption.Style.Alignment.WrapText = true;
                    if (page.Cell(3, 4).IsMerged())
                    {
                        Caption = page.Range(CellStartCaption, lastCellCaption.CellBelow());
                        MergeRange(ref Caption);
                    }
                    Regex rComment = new Regex(@"КОРРЕКЦИИ");
                    int corrRow = 5;
                    Match Mcomment = rComment.Match(page.Cell(corrRow, 1).GetString().ToUpper());
                    while (!Mcomment.Success)
                    {
                        corrRow++;
                        Mcomment = rComment.Match(page.Cell(corrRow, 1).GetString().ToUpper());
                    }
                    var CellPoint = CellStartCaption.CellBelow();
                    IXLCell lastCell = CellPoint;
                    double totalSumLast = 0;
                    double totalSumPreLast = 0;
                    int totalQtyPointsLast = 0;
                    int totalQtyPointsPreLast = 0;
                    int qtyNoChange = 0;
                    int qtyBetter = 0;
                    int qtyWorse = 0;
                    while (CellPoint.Address.RowNumber < corrRow - 4)
                    {
                        var CellNamePoint = page.Cell(CellPoint.Address.RowNumber, numColPoint);
                        if (dictPoints.ContainsKey(CellNamePoint.GetString()))
                        {
                            int qtyRed = dictPoints[CellNamePoint.GetString()].Key;
                            int qtyAll = dictPoints[CellNamePoint.GetString()].Value;
                            CellPoint.CellLeft().Style.Fill.BackgroundColor = XLColor.Red;
                            CellPoint.Value = qtyRed;
                            CellPoint.Style.NumberFormat.NumberFormatId = OutPutDoc.getFormatData(DataForPrint.Estimate.qty);
                            CellPoint.CellRight().Value = qtyAll;
                            CellPoint.CellRight().Style.NumberFormat.NumberFormatId = OutPutDoc.getFormatData(DataForPrint.Estimate.qty);
                            CellPoint.CellRight().CellRight().Value = qtyAll - qtyRed;
                            CellPoint.CellRight().CellRight().Style.NumberFormat.NumberFormatId = OutPutDoc.getFormatData(DataForPrint.Estimate.qty);
                            double AVGLast = (double)(qtyAll - qtyRed) / qtyAll;
                            CellPoint.CellRight().CellRight().CellRight().Value = AVGLast;
                            
                            CellPoint.CellRight().CellRight().CellRight().Style.NumberFormat.NumberFormatId = OutPutDoc.getFormatData(DataForPrint.Estimate.AVG);
                            if (AVGLast < 0.8)
                            {
                                CellPoint.CellRight().CellRight().CellRight().Style.Fill.BackgroundColor = XLColor.Red;
                            }
                            totalSumLast += AVGLast;
                            totalQtyPointsLast++;
                            lastCell = CellPoint.CellRight().CellRight().CellRight();
                            if (PreLastMonthManager != null && PreLastMonthManager.Name == m.Name && PreLastMonthManager.getStages().Exists(s => s.name == stage.name))
                            {
                                howChangeCell = page.Cell(CellPoint.Address.RowNumber, howChangeCell.Address.ColumnNumber);
                                
                                lastCell = howChangeCell;
                                if (dictPointsPreLastMonth.ContainsKey(CellNamePoint.GetString()))
                                {
                                    prelastMonthCell = page.Cell(CellPoint.Address.RowNumber, prelastMonthCell.Address.ColumnNumber);
                                    int qtyAllPreLast = dictPointsPreLastMonth[CellNamePoint.GetString()].Value;
                                    int qtyRedPreLast = dictPointsPreLastMonth[CellNamePoint.GetString()].Key;
                                    double AvgPreLast = (double)(qtyAllPreLast - qtyRedPreLast) / qtyAllPreLast;
                                    prelastMonthCell.Value = AvgPreLast;
                                    prelastMonthCell.Style.NumberFormat.NumberFormatId = OutPutDoc.getFormatData(DataForPrint.Estimate.AVG);
                                    totalSumPreLast += AvgPreLast;
                                    totalQtyPointsPreLast++;
                                    if (AvgPreLast < 0.8)
                                    {
                                        prelastMonthCell.Style.Fill.BackgroundColor = XLColor.Red;
                                    }
                                    
                                    if (AVGLast < AvgPreLast)
                                    {
                                        howChangeCell.Value = "Ухудшилось";
                                        howChangeCell.Style.Fill.BackgroundColor = XLColor.Red;
                                        qtyWorse++;
                                    }
                                    if (AVGLast > AvgPreLast)
                                    {
                                        howChangeCell.Value = "Улучшилось";
                                        howChangeCell.Style.Fill.BackgroundColor = XLColor.BrightGreen;
                                        qtyBetter++;
                                    }
                                    if (AVGLast == AvgPreLast)
                                    {
                                        howChangeCell.Value = "Не изменилось";
                                        qtyNoChange++;
                                    }

                                }
                                else
                                {
                                    howChangeCell.Value = "Критерий изменился";

                                }
                            }
                        }
                        CellPoint = CellPoint.CellBelow();
                    }
                    QtyNoChangeCell.Value = qtyNoChange;
                    QtyNoChangeCell.Style.NumberFormat.NumberFormatId = OutPutDoc.getFormatData(DataForPrint.Estimate.qty);
                    QtyWorseCell.Value = qtyWorse;
                    QtyWorseCell.Style.NumberFormat.NumberFormatId = OutPutDoc.getFormatData(DataForPrint.Estimate.qty);
                    QtyBetterCell.Value = qtyBetter;
                    QtyBetterCell.Style.NumberFormat.NumberFormatId = OutPutDoc.getFormatData(DataForPrint.Estimate.qty);
                    totalAVGLastMonthCell.Value = totalSumLast / totalQtyPointsLast;
                    totalAVGLastMonthCell.Style.NumberFormat.NumberFormatId = OutPutDoc.getFormatData(DataForPrint.Estimate.AVG);
                    totalAVGPreLastMonthCell.Value = totalSumPreLast / totalQtyPointsPreLast;
                    totalAVGPreLastMonthCell.Style.NumberFormat.NumberFormatId = OutPutDoc.getFormatData(DataForPrint.Estimate.AVG);
                    var rngTable = page.Range(CellStartCaption.CellLeft(), lastCell);
                    //rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    rngTable.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    rngTable.Style.Border.OutsideBorder = XLBorderStyleValues.Thin; ;
                    rngTable.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }
            }

            return wb;
        }
        static void Range(ref IXLRange rng)
        {
            rng.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rng.Style.Border.OutsideBorder = XLBorderStyleValues.Thin; ;
            rng.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            
        }
        static void MergeRange(ref IXLRange rng)
        {
            var curCell = rng.FirstCell();
            for (int i = rng.FirstColumn().ColumnNumber(); i<= rng.LastColumn().ColumnNumber(); i++)
            {
                var newrngmrg = rng.Range(curCell,curCell.CellBelow());
                    newrngmrg.Merge();
                curCell = curCell.CellRight();
            }
        }
        public static XLWorkbook getStatistic(List<Manager> lm, DateTime firstDate)
        {
            XLWorkbook wbAnalytic = new XLWorkbook();
            var page = wbAnalytic.AddWorksheet("Еженедельная сводка");
            var CaptionTable = page.Cell("A1");
            CaptionTable.Value = "Касание с компанией за период с " + firstDate.ToString("dd.MM") + " по " + DateTime.Now.AddDays(-2).ToString("dd.MM");
            var ManagerCell = page.Cell("A2");
            var BadPointCell = ManagerCell.CellRight();
            var BadCommentCell = BadPointCell.CellRight(); 
            var GoodCorrectionCell = BadCommentCell.CellRight();
            var WorseCallCell = GoodCorrectionCell.CellRight();
            var BestCallCell = WorseCallCell.CellRight();
            var qtyCell = BestCallCell.CellRight();
            var AVGCell = qtyCell.CellRight();
            var rngCaption = page.Range(ManagerCell, AVGCell);
            rngCaption.Style.Font.Bold = true;
            ManagerCell.Value = "Менеджер";
            BadPointCell.Value = "Систематически невыполняемые пункты";
            BadCommentCell.Value = "Коррекции, отклоняющиеся от нормы";
            WorseCallCell.Value = "Худший звонок";
            BestCallCell.Value = "Лучший звонок";
            GoodCorrectionCell.Value = "Положительные коррекции";
            qtyCell.Value = "Всего звонков за период";
            AVGCell.Value = "Средний % по звонкам";
            ManagerCell.WorksheetColumn().Width = 15;
            BadPointCell.WorksheetColumn().Width = 30;
            BadCommentCell.WorksheetColumn().Width = 30;
            qtyCell.WorksheetColumn().Width = 10;
            AVGCell.WorksheetColumn().Width = 10;
            WorseCallCell.WorksheetColumn().Width = 15;
            BestCallCell.WorksheetColumn().Width = 15;
            GoodCorrectionCell.WorksheetColumn().Width = 30;
            foreach (var m in lm)
            {
                string BadPoints = m.getBadPoints(firstDate);
                string BadComments = m.getBadComments(firstDate);
                int qty = m.getCountOfCalls(firstDate);
                double AVGPerCent = m.getAVGPersent(firstDate);
                ManagerCell = ManagerCell.CellBelow();
                ManagerCell.Value = m.Name;
                ManagerCell.Style.Font.Bold = true;
                BadPointCell = BadPointCell.CellBelow();
                BadPointCell.Value = BadPoints;
                BadCommentCell = BadCommentCell.CellBelow();
                BadCommentCell.Value = BadComments;
                WorseCallCell = WorseCallCell.CellBelow();
                WorseCallCell.Value = m.getWorseCall(firstDate);
                qtyCell = qtyCell.CellBelow();
                qtyCell.Value = qty;
                qtyCell.Style.NumberFormat.NumberFormatId = OutPutDoc.getFormatData(DataForPrint.Estimate.qty);
                AVGCell = AVGCell.CellBelow();
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                AVGCell.Value = AVGPerCent == -1 ? "" : String.Format("{0:0.####}", AVGPerCent); 
                AVGCell.Style.NumberFormat.NumberFormatId = OutPutDoc.getFormatData(DataForPrint.Estimate.AVG);
            }
            var Rng = page.RangeUsed();
            var Caption = page.Range(1, 1, 1, Rng.LastColumn().ColumnNumber());
            Caption.Style.Font.Bold = true;
            Caption.Merge();
            Rng.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            Rng.Style.Border.OutsideBorder = XLBorderStyleValues.Thin; ;
            Rng.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            Rng.Style.Alignment.WrapText = true;
            return wbAnalytic;
        }

    }
}
