using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessingCheckListWss.ProcessingCheckLists
{
    class OutPutDoc
    {
        XLWorkbook wbout = new XLWorkbook();
        
        public OutPutDoc(Dictionary<string, Dictionary<string, List<DataForPrint>>> printPagesByMonth, bool totalopt = false)
        {
            string LastMonth = printPagesByMonth.Keys.Last();
            
            foreach (var stage in printPagesByMonth[LastMonth].Keys)
            {
                if (stage == "ТЗ отправлено")
                {

                }
                IXLWorksheet worksheet;
                bool qtyFull = true;
                IXLCell Cell;
                int firstCol = 1;
                int lastCol = firstCol;
                if (!totalopt)
                {
                    worksheet = wbout.AddWorksheet(stage);
                    Cell = worksheet.Cell("A1");
                    
                }
                else
                {
                    worksheet = wbout.AddWorksheet("Лишняя");
                    Cell = worksheet.Cell("A1");
                }
                string NameList = "";
                foreach (var opt in DataForPrint.getEstimates(totalopt))
                {
                    if (totalopt)
                    {
                        NameList = getpageCaption(opt);//.Substring(0,30);
                        if (!wbout.Worksheets.Contains(NameList))
                        {
                            worksheet = wbout.AddWorksheet(NameList);
                            Cell = worksheet.Cell("A1");
                            firstCol = 1;
                            lastCol = 1;
                        }
                        else
                        {
                            worksheet = wbout.Worksheet(NameList);
                            firstCol = worksheet.RangeUsed().LastColumn().ColumnNumber() + 2;
                            lastCol = firstCol;
                            Cell = worksheet.Cell(1, firstCol);
                        }
                    }
                    Cell.Value = gettableCaption(opt, qtyFull);
                    Cell = Cell.CellBelow();
                    Cell.Value = "Менеджер \\ Месяц";
                    
                    int firstRow = 1;
                    int lastRow = 2;
                    int lastRowCaption = 2;
                    var CellManager = Cell.CellBelow();
                    if (opt != DataForPrint.Estimate.badPoints && opt!=DataForPrint.Estimate.Objection)
                    {
                        foreach (var manager in printPagesByMonth[LastMonth][stage])
                        {
                            lastRow++;
                            CellManager.Value = manager.manager;
                            CellManager = CellManager.CellBelow();
                        }


                        if (opt == DataForPrint.Estimate.duration || opt == DataForPrint.Estimate.qty)
                        {
                            lastRow++;
                            worksheet.Cell(lastRow, CellManager.Address.ColumnNumber).Value = "ИТОГО";
                        }

                        var CellMonth = Cell;
                        foreach (var month in printPagesByMonth.Keys)
                        {
                            if (!(!printPagesByMonth[month].ContainsKey(stage) || !haveCalls(printPagesByMonth[month][stage])))
                            {
                                lastCol++;
                                CellMonth = CellMonth.CellRight();
                                CellManager = Cell.CellBelow();
                                if (printPagesByMonth[month].ContainsKey(stage))
                                {

                                    CellMonth.Value = month;
                                    var CellPrintValue = worksheet.Cell(CellManager.Address.RowNumber, CellMonth.Address.ColumnNumber);
                                    if (opt == DataForPrint.Estimate.badPoints)
                                    {
                                        CellMonth.WorksheetColumn().Width = 30;
                                        CellMonth.WorksheetColumn().Style.Alignment.WrapText = true;
                                    }
                                    int sumQty = 0;
                                    TimeSpan sumDuaration = new TimeSpan(0);
                                    foreach (var manager in printPagesByMonth[LastMonth][stage])
                                    {
                                        
                                        string val1 = getValueOfPointOfManager(printPagesByMonth[month][stage], CellManager.GetString(), opt, qtyFull);
                                        if (val1 != "")
                                        {
                                            CellPrintValue.Value = val1;
                                            if (opt == DataForPrint.Estimate.qty)
                                                sumQty += int.Parse(val1);
                                            if (opt == DataForPrint.Estimate.duration)
                                                sumDuaration += TimeSpan.Parse(val1);
                                        }
                                        CellManager = CellManager.CellBelow();
                                        CellPrintValue = worksheet.Cell(CellManager.Address.RowNumber, CellMonth.Address.ColumnNumber);
                                    }
                                    
                                    if (opt == DataForPrint.Estimate.qty)
                                        CellPrintValue.Value = sumQty;
                                    if (opt == DataForPrint.Estimate.duration)
                                        CellPrintValue.Value = sumDuaration;
                                }
                            }

                        }
                    }
                    else
                    {
                        if (opt == DataForPrint.Estimate.badPoints)
                        {
                            foreach (var mm in printPagesByMonth[LastMonth][stage])
                            {
                                lastRow++;
                                CellManager.Value = mm.manager;

                                int lastRowinMonth = lastRow;
                                var CellMonth = Cell;
                                foreach (var month in printPagesByMonth.Keys)
                                {
                                    lastRowinMonth = CellManager.Address.RowNumber;
                                    if ((printPagesByMonth[month].ContainsKey(stage) && haveCalls(printPagesByMonth[month][stage])))
                                    {
                                        CellMonth = CellMonth.CellRight();
                                        if (printPagesByMonth[month].ContainsKey(stage))
                                        {

                                            CellMonth.Value = month;




                                            CellMonth.WorksheetColumn().Width = 30;
                                            CellMonth.WorksheetColumn().Style.Alignment.WrapText = true;

                                            string val1 = getValueOfPointOfManager(printPagesByMonth[month][stage], mm.manager, opt, qtyFull);
                                            string[] points = val1.Split(';');
                                            foreach (string p in points)
                                            {
                                                p.Trim('\n');
                                                var CellPrintValue = worksheet.Cell(lastRowinMonth, CellMonth.Address.ColumnNumber);
                                                if (p.Trim('\n') != "")
                                                {
                                                    if (CellPrintValue.GetString() != "")
                                                        lastRowinMonth++;
                                                    CellPrintValue = worksheet.Cell(lastRowinMonth, CellMonth.Address.ColumnNumber);
                                                    CellPrintValue.SetValue<string>(p.Trim('\n'));
                                                }
                                            }


                                        }
                                    }

                                    if (lastRow < lastRowinMonth + 1)
                                    {
                                        lastRow = lastRowinMonth + 1;
                                    }

                                }
                                if (lastCol < CellMonth.Address.ColumnNumber)
                                    lastCol = CellMonth.Address.ColumnNumber;
                                worksheet.Range(CellManager, worksheet.Cell(lastRow - 1, CellManager.Address.ColumnNumber)).Merge();
                                worksheet.Range(lastRow, CellManager.WorksheetColumn().ColumnNumber(), lastRow, lastCol).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 0, 112, 192);

                                int qtyRows = lastRow - CellManager.WorksheetRow().RowNumber();
                                if (qtyRows > 6)
                                {
                                    CellManager.Style.Alignment.TopToBottom = true;
                                    CellManager.Style.Font.FontSize = Math.Max(8,Math.Min(qtyRows, 24));
                                }
                                CellManager.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                CellManager = worksheet.Cell(lastRow + 1, CellManager.Address.ColumnNumber);
                            }
                            worksheet.Range(lastRow, CellManager.WorksheetColumn().ColumnNumber(), lastRow, lastCol).Style.Fill.BackgroundColor = XLColor.NoColor;
                        }
                        else
                        {
                            Cell.Value = "Месяц";
                            List<string> tags = new List<string>();
                            


                            CellManager = Cell.CellBelow();
                            CellManager.Value = "Менеджер";
                            
                            lastRow++;
                            foreach (var manager in printPagesByMonth[LastMonth][stage])
                            {
                                CellManager = CellManager.CellBelow();
                                lastRow++;
                                CellManager.Value = manager.manager;
                                
                            }

                            //worksheet.Range(1, 12, 1, tags.Count + 13).Merge();
                            var CellMonth = Cell;
                            bool hasData = false;
                            foreach (var pair in printPagesByMonth.Where(p => p.Value.Count > 0))
                            {
                                var month = pair.Key;
                                Dictionary<string, int> val = new Dictionary<string, int>();
                                try
                                {
                                     val = getObjections(printPagesByMonth[month][stage.Trim()], CellManager.GetString());
                                }
                                catch (Exception)
                                {
                                    //var sh = printPagesByMonth[month][stage];
                                    getObjections(printPagesByMonth[month][stage.Trim()], CellManager.GetString());
                                }
                                tags = val.Keys.ToList();
                                if (printPagesByMonth[month].ContainsKey(stage) && haveObj(printPagesByMonth[month][stage]))
                                {
                                    hasData = true;
                                    CellMonth = worksheet.Cell(2, lastCol + 1);
                                    CellMonth.Value = month;
                                    lastCol += tags.Count;
                                    worksheet.Range(CellMonth, worksheet.Cell(2, lastCol)).Merge();
                                    worksheet.Range(CellMonth, worksheet.Cell(2, lastCol)).Value = month;
                                    CellMonth.Value = month;
                                    for (int i = 0; i < tags.Count; i++)
                                    {
                                        worksheet.Cell(3, CellMonth.Address.ColumnNumber + i).Value = tags[i];
                                        
                                    }
                                    
                                    for (int i = 4; i <= lastRow; i++)
                                    {
                                        val = getObjections(printPagesByMonth[month][stage], worksheet.Cell(i, CellManager.Address.ColumnNumber).GetString());
                                        try
                                        {
                                            for (int j = 0; j < tags.Count; j++)
                                            {
                                                worksheet.Cell(i, CellMonth.Address.ColumnNumber + j).Value = val.Where(p => p.Key == worksheet.Cell(3, CellMonth.Address.ColumnNumber + j).GetString()).First().Value;
                                            }
                                            worksheet.Cell(i, CellMonth.Address.ColumnNumber + tags.Count - 1).Style.Fill.BackgroundColor = XLColor.Yellow;

                                        }
                                        catch (InvalidOperationException)
                                        {
                                            
                                        }

                                    }
                                    worksheet.Cell(lastRow + 1, CellMonth.Address.ColumnNumber + tags.Count - 1).Style.Fill.BackgroundColor = XLColor.Red;

                                }
                                
                            }

                            lastRow++;
                            worksheet.Cell(lastRow, CellManager.Address.ColumnNumber).Value = "ИТОГО";
                            

                           for (int j = Cell.Address.ColumnNumber + 1; j <= lastCol; j++)
                           {
                                int sum = 0;
                                
                                    for (int i = 4; i < lastRow; i++)
                                    {
                                        try
                                        {
                                            sum += worksheet.Cell(i, j).GetValue<int>();
                                        }
                                        catch (FormatException)
                                        {
                                            
                                        }
                                    }
                                worksheet.Cell(lastRow, j).Value = sum;
                                if (worksheet.Cell(lastRow, j).Style.Fill.BackgroundColor != XLColor.Red)
                                    worksheet.Cell(lastRow, j).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 0, 112, 192);
                                if (sum == 0)
                                {
                                    worksheet.Column(j).Delete();
                                    j--;
                                    lastCol--;
                                }
                                
                                

                            }
                            if (!hasData)
                            {
                                worksheet.Column(CellManager.Address.ColumnNumber).Delete();
                                lastCol -= 2;
                            }
                            else
                            {
                                worksheet.Range(worksheet.Cell(lastRow, CellManager.Address.ColumnNumber), worksheet.Cell(lastRow, lastCol)).Style.Font.Bold = true;
                                worksheet.Range(worksheet.Cell(3, CellManager.Address.ColumnNumber), worksheet.Cell(3, lastCol)).Style.Font.Bold = true;
                                
                            }

                            //        {
                            //            it++;
                            //            wsheet.Cell("L" + it).Value = m.Name;
                            //            var listcalls = m.GetCalls().Where(c => c.dateOfCall >= firstDate && c.dateOfCall <= LastDate && c.Objections != "" && c.Objections.Trim().ToLower() != "нет");

                            //            for (int i = 0; i < tags.Count; i++)
                            //            {
                            //                string tag = wsheet.Cell(2, i + 13).GetString();
                            //                wsheet.Cell(it, i + 13).SetValue<int>(listcalls.Where(c => Regex.Match(c.Objections, tag, RegexOptions.IgnoreCase).Success).Count());

                            //            }
                            //            wsheet.Cell(it, 13 + tags.Count).SeValue<int>(listcalls.Where(c => !tags.Where(tag => Regex.Match(c.Objections, tag, RegexOptions.IgnoreCase).Success).Any()).Count());

                            //        }
                            //        it++;
                            //        wsheet.Cell("L" + it).Value = "Итого";
                            //        for (int i = 0; i < tags.Count + 1; i++)
                            //        {
                            //            int qty = 0;
                            //            for (int j = 3; j < it; j++)
                            //            {
                            //                qty += wsheet.Cell(j, 13 + i).GetValue<int>();
                            //            }
                            //            wsheet.Cell(it, 13 + i).Value = qty;

                            //        }
                            //        Rng = wsheet.Range(1, 12, it, 13 + tags.Count);
                            //        Rng.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                            //        Rng.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            //        Rng.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


                            //        wsheet.Column(12).Width = 20;
                            //        wsheet.Columns(13, 13 + tags.Count).Width = 20;
                            //        Rng.Style.Alignment.WrapText = true;
                            //        wsheet.Range("L1", "L" + it).Style.Font.Bold = true;
                            //        wsheet.Range(2, 12, 2, 13 + tags.Count).Style.Font.Bold = true;
                            //        wsheet.Range(it, 12, it, 13 + tags.Count).Style.Font.Bold = true;
                        }
                    }
                    if (firstCol <= lastCol)
                    {
                        if (opt == DataForPrint.Estimate.badPoints)
                            lastRow--;
                        var rngTable = worksheet.Range(firstRow, firstCol, lastRow, lastCol);
                        rngTable.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        rngTable.Style.Border.OutsideBorder = XLBorderStyleValues.Thin; ;
                        rngTable.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        if (opt == DataForPrint.Estimate.duration || opt == DataForPrint.Estimate.qty)
                            rngTable.LastRow().Style.Font.Bold = true;
                        var Caption = worksheet.Range(firstRow, firstCol, firstRow, lastCol);
                        Caption.Merge();
                        Caption = worksheet.Range(firstRow, firstCol, lastRowCaption, lastCol);
                        Caption.Style.Font.Bold = true;
                        var TableData = worksheet.Range(lastRowCaption + 1, firstCol + 1, lastRow, lastCol);
                        if (opt != DataForPrint.Estimate.badPoints)
                            TableData.Style.NumberFormat.NumberFormatId = getFormatData(opt);
                    }
                    //rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin;

                    var rng = worksheet.RangeUsed();
                    try
                    {
                        firstCol = rng.LastColumn().ColumnNumber() + 2;
                    }
                    catch(NullReferenceException)
                    {
                        firstCol = 1;
                    }
                    lastCol = firstCol;
                    Cell = worksheet.Cell(1,firstCol);
                    if (opt == DataForPrint.Estimate.qty)
                        qtyFull = false;
                    worksheet.Columns().AdjustToContents(); //ширина столбца
                }
                if (totalopt)
                  wbout.Worksheets.Delete(1);
                
            }
        }
        string getValueOfPointOfManager(List<DataForPrint> managers, string manager, DataForPrint.Estimate opt, bool qtyFull = true)
        {
            string returnValue = "";
            foreach (var man in managers)
            {
                if (man.manager == manager)
                {
                    if (opt == DataForPrint.Estimate.AVG)
                    {
                        System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                        returnValue = man.AVGPercent == -1 ? "" : String.Format("{0:0.####}", man.AVGPercent);
                    }
                    if (opt == DataForPrint.Estimate.qty && qtyFull)
                    {
                        returnValue = man.qty.ToString();
                    }
                    if (opt == DataForPrint.Estimate.qty && !qtyFull)
                    {
                        returnValue = man.qtyWithoutIncoming.ToString();
                    }
                    if (opt == DataForPrint.Estimate.duration)
                    {
                        returnValue = man.duration.ToString();
                    }
                    if (opt == DataForPrint.Estimate.badPoints)
                    {
                        returnValue = man.BadPoints.ToString();
                    }
                    if (opt == DataForPrint.Estimate.AVGDuration)
                    {
                        returnValue = man.AVGduration.ToString();
                    }
                    break;
                }
            }
            
            return returnValue;
        }
        public Dictionary<string, int> getObjections (List<DataForPrint> managers, string manager)
        {
            try
            {
                var man = managers.Where(m => m.manager == manager).First();
                return man.Objections;
            }
            catch(System.InvalidOperationException)
            {
                var man = managers.First();
                var Obj = man.Objections;
                Dictionary<string, int> Obj2 = new Dictionary<string, int>();
                foreach (var o1 in Obj.Keys)
                {
                    Obj2[o1] = 0; 
                }
                return Obj2;
            }
            
        }
        public static int getFormatData(DataForPrint.Estimate opt)
        {
            if (opt == DataForPrint.Estimate.AVG)
                return 10;
            if (opt == DataForPrint.Estimate.qty)
                return 1;
            if (opt == DataForPrint.Estimate.duration)
                return 46;
            if (opt == DataForPrint.Estimate.AVGDuration)
                return 46;
            if (opt == DataForPrint.Estimate.Objection)
                return 1;

            return 0;
        }
        string gettableCaption(DataForPrint.Estimate opt, bool qtyFull = true)
        {
            if (opt == DataForPrint.Estimate.AVG)
                return "Средний %";
            if (opt == DataForPrint.Estimate.qty && qtyFull)
                return "Количество";
            if (opt == DataForPrint.Estimate.qty && !qtyFull)
                return "Количество без входящих и было не удобно разговаривать";
            if (opt == DataForPrint.Estimate.duration)
                return "Продолжительность";
            if (opt == DataForPrint.Estimate.AVGDuration)
                return "Средняя продолжительность";
            if (opt == DataForPrint.Estimate.Objection)
                return "Статистика возражений";
            if (opt == DataForPrint.Estimate.badPoints)
            {
                return "Плохо выполняемые пункты";
            }
            return "";
        }

        string getpageCaption(DataForPrint.Estimate opt)
        {
            if (opt == DataForPrint.Estimate.AVG)
                return "Средний %";
            if (opt == DataForPrint.Estimate.qty)
                return "Количество и продолжительность";
            if (opt == DataForPrint.Estimate.duration)
                return "Количество и продолжительность";
            if (opt == DataForPrint.Estimate.AVGDuration)
                return "Количество и продолжительность";
            if (opt == DataForPrint.Estimate.Objection)
                return "Статистика возражений";
            if (opt == DataForPrint.Estimate.badPoints)
            {
                return "Плохо выполняемые пункты";
            }
            return "";
        }
        bool haveCalls (List<DataForPrint> managers)
        {
            foreach (var man in managers)
            {
                if (man.qty > 0)
                    return true;
            }
            return false;
        }
        bool haveObj(List<DataForPrint> managers)
        {
            foreach (var man in managers)
            {
                foreach (var obj in man.Objections)
                {
                    if (obj.Value > 0)
                        return true;
                }
                    
            }
            return false;
        }
        public XLWorkbook getWb()
        {
            return wbout;
        }
    }
}

