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
                        NameList = gettableCaption(opt);//.Substring(0,30);
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
                    foreach (var manager in printPagesByMonth[LastMonth][stage])
                    {
                        lastRow++;
                        CellManager.Value = manager.manager;
                        CellManager = CellManager.CellBelow();
                    }

                    var CellMonth = Cell;
                    foreach (var month in printPagesByMonth.Keys)
                    {
                        if (!(opt == DataForPrint.Estimate.duration && month == "Февраль" || !printPagesByMonth[month].ContainsKey(stage) || !haveCalls(printPagesByMonth[month][stage])))
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
                                foreach (var manager in printPagesByMonth[LastMonth][stage])
                                {
                                    string val1 = getValueOfPointOfManager(printPagesByMonth[month][stage], CellManager.GetString(), opt, qtyFull);
                                    if (val1 != "")
                                        CellPrintValue.Value = val1;
                                    CellManager = CellManager.CellBelow();
                                    CellPrintValue = worksheet.Cell(CellManager.Address.RowNumber, CellMonth.Address.ColumnNumber);
                                }
                            }
                        }

                    }
                    var rngTable = worksheet.Range(firstRow, firstCol, lastRow, lastCol);
                    //rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    rngTable.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    rngTable.Style.Border.OutsideBorder = XLBorderStyleValues.Thin; ;
                    rngTable.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    var Caption = worksheet.Range(firstRow, firstCol, firstRow, lastCol);
                    Caption.Merge();
                    Caption = worksheet.Range(firstRow, firstCol, lastRowCaption, lastCol);
                    Caption.Style.Font.Bold = true;
                    var TableData = worksheet.Range(lastRowCaption + 1, firstCol + 1, lastRow, lastCol);
                    if (opt != DataForPrint.Estimate.badPoints)
                      TableData.Style.NumberFormat.NumberFormatId = getFormatData(opt);
                    firstCol = worksheet.RangeUsed().LastColumn().ColumnNumber() + 2;
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
                    break;
                }
            }
            
            return returnValue;
        }
        public static int getFormatData(DataForPrint.Estimate opt)
        {
            if (opt == DataForPrint.Estimate.AVG)
                return 10;
            if (opt == DataForPrint.Estimate.qty)
                return 1;
            if (opt == DataForPrint.Estimate.duration)
                return 46;
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
        public XLWorkbook getWb()
        {
            return wbout;
        }
    }
}

