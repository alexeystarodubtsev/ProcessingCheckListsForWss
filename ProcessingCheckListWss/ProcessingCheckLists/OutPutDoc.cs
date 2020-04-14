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
        
        
        public OutPutDoc(Dictionary<string, Dictionary<string, List<DataForPrint>>> printPagesByMonth)
        {
            string LastMonth = printPagesByMonth.Keys.Last();
            
            foreach (var stage in printPagesByMonth[LastMonth].Keys)
            {
                var worksheet = wbout.AddWorksheet(stage);

                var Cell = worksheet.Cell("A1");
                int firstCol = 1;
                int lastCol = firstCol;
                foreach (var opt in DataForPrint.getEstimates())
                {
                    
                    Cell.Value = gettableCaption(opt);
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
                        if (!(opt == DataForPrint.Estimate.duration && month == "Февраль" || !haveCalls(printPagesByMonth[month][stage])))
                        {
                            lastCol++;
                            CellMonth = CellMonth.CellRight();
                            CellManager = Cell.CellBelow();
                            if (printPagesByMonth[month].ContainsKey(stage))
                            {

                                CellMonth.Value = month;
                                var CellPrintValue = worksheet.Cell(CellManager.Address.RowNumber, CellMonth.Address.ColumnNumber);
                                foreach (var manager in printPagesByMonth[LastMonth][stage])
                                {
                                    string val1 = getValueOfPointOfManager(printPagesByMonth[month][stage], CellManager.GetString(), opt);
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
                    TableData.Style.NumberFormat.NumberFormatId = getFormatData(opt);
                    firstCol = worksheet.RangeUsed().LastColumn().ColumnNumber() + 2;
                    lastCol = firstCol;
                    Cell = worksheet.Cell(1,firstCol);
                }
                worksheet.Columns().AdjustToContents(); //ширина столбца
            }
        }
        string getValueOfPointOfManager(List<DataForPrint> managers, string manager, DataForPrint.Estimate opt)
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
                    if (opt == DataForPrint.Estimate.qty)
                    {
                        returnValue = man.qty.ToString();
                    }
                    if (opt == DataForPrint.Estimate.duration)
                    {
                        returnValue = man.duration.ToString();
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
        string gettableCaption(DataForPrint.Estimate opt)
        {
            if (opt == DataForPrint.Estimate.AVG)
                return "Средний %";
            if (opt == DataForPrint.Estimate.qty)
                return "Количество";
            if (opt == DataForPrint.Estimate.duration)
                return "Продолжительность";
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

