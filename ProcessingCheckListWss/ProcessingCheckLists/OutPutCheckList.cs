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
        
        public  static XLWorkbook FillAnalyticOfPoint(Manager m)
        {
            XLWorkbook wb = new XLWorkbook(m.FilePath);

            foreach (var stage in m.getStages())
            {
                var dictPoints = stage.getStatisticOfPoints();
                var page = wb.Worksheet(stage.name);
                var table = page.RangeUsed();
                int curCol = table.LastColumn().ColumnNumber() + 2;
                const int numColPoint = 4;
                IXLCell CellStartCaption = page.Cell(2, curCol);
                CellStartCaption.Value = "Количество некорректных пунктов";
                CellStartCaption.WorksheetColumn().Width = 11;
                CellStartCaption.CellRight().Value = "Количество прохождений пункта";
                CellStartCaption.CellRight().WorksheetColumn().Width = 11;
                CellStartCaption.CellRight().CellRight().Value = "Количество корректных пунктов";
                CellStartCaption.CellRight().CellRight().WorksheetColumn().Width = 11;
                CellStartCaption.CellRight().CellRight().CellRight().Value = "% Выполнения";
                CellStartCaption.CellRight().CellRight().CellRight().WorksheetColumn().Width = 11;
                var Caption = page.Range(CellStartCaption, CellStartCaption.CellRight().CellRight().CellRight());
                Caption.Style.Alignment.WrapText = true;
                Regex rComment = new Regex(@"КОРРЕКЦИИ");
                int corrRow = 5;
                Match Mcomment = rComment.Match(page.Cell(corrRow, 1).GetString().ToUpper());
                while (!Mcomment.Success)
                {
                    corrRow++;
                    Mcomment = rComment.Match(page.Cell(corrRow, 1).GetString().ToUpper());
                }
                var CellPoint = CellStartCaption.CellBelow();
                while (CellPoint.Address.RowNumber < corrRow - 4)
                {
                    var CellNamePoint = page.Cell(CellPoint.Address.RowNumber, numColPoint);
                    if (dictPoints.ContainsKey(CellNamePoint.GetString()))
                    {
                        int qtyRed = dictPoints[CellNamePoint.GetString()].Key;
                        int qtyAll = dictPoints[CellNamePoint.GetString()].Value;
                        CellPoint.Value = qtyRed;
                        CellPoint.Style.NumberFormat.NumberFormatId = OutPutDoc.getFormatData(DataForPrint.Estimate.qty);
                        CellPoint.CellRight().Value = qtyAll;
                        CellPoint.CellRight().Style.NumberFormat.NumberFormatId = qtyRed;
                        CellPoint.CellRight().CellRight().Value = qtyAll - qtyRed;
                        CellPoint.CellRight().CellRight().Style.NumberFormat.NumberFormatId = OutPutDoc.getFormatData(DataForPrint.Estimate.qty);
                        CellPoint.CellRight().CellRight().CellRight().Value = (double)(qtyAll - qtyRed) / qtyAll;
                        CellPoint.CellRight().CellRight().CellRight().Style.NumberFormat.NumberFormatId = OutPutDoc.getFormatData(DataForPrint.Estimate.AVG);
                        if ((double)(qtyAll - qtyRed) / qtyAll < 0.8)
                        {
                            CellPoint.CellRight().CellRight().CellRight().Style.Fill.BackgroundColor = XLColor.Red;
                        }
                    }
                    CellPoint = CellPoint.CellBelow();
                }
                var rngTable = page.Range(CellStartCaption, CellPoint.CellAbove().CellRight().CellRight().CellRight());
                //rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                rngTable.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                rngTable.Style.Border.OutsideBorder = XLBorderStyleValues.Thin; ;
                rngTable.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            }

            return wb;
        }
        public static XLWorkbook getStatistic(List<Manager> lm, DateTime firstDate)
        {
            XLWorkbook wbAnalytic = new XLWorkbook();
            var page = wbAnalytic.AddWorksheet("Еженедельная сводка");
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
            Rng.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            Rng.Style.Border.OutsideBorder = XLBorderStyleValues.Thin; ;
            Rng.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            Rng.Style.Alignment.WrapText = true;
            return wbAnalytic;
        }

    }
}
