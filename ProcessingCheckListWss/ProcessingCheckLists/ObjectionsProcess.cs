using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProcessingCheckListWss.ProcessingCheckLists
{
    class ObjectionsProcess
    {
        public static XLWorkbook GetXLWorkbook(List<Manager> lm, DateTime firstDate, DateTime LastDate)
        {
            var wb =  new XLWorkbook();
            var wsheet = wb.AddWorksheet("Возражения");
            List<Call> calls = new List<Call>();
            wsheet.Cell("A1").Value = "Возражения по звонкам с " + firstDate.ToString("dd.MM") + " по " + LastDate.ToString("dd.MM");
            wsheet.Range("A1", "E1").Merge();
            wsheet.Cell("A2").Value = "Менеджер";
            wsheet.Cell("B2").Value = "Клиент";
            wsheet.Cell("C2").Value = "Возражение";
            wsheet.Cell("D2").Value = "Как отработал";
            wsheet.Cell("E2").Value = "Результат";
            int firstrow = 3;
            int lastrow = 2;
            foreach (var m in lm)
            {
                var listcalls = m.GetCalls().Where(c => c.dateOfCall >= firstDate && c.dateOfCall <= LastDate && c.Objections != "");
                foreach (var call in listcalls)
                {
                    lastrow++;
                    wsheet.Cell("B" + lastrow).Value = call.client;
                    if (call.ClientLink != "")
                    {
                        wsheet.Cell("B" + lastrow).Hyperlink = new XLHyperlink( call.ClientLink);
                    }
                    wsheet.Cell("C" + lastrow).Value = call.Objections;
                    wsheet.Cell("D" + lastrow).Value = call.howProcessObj;
                    wsheet.Cell("E" + lastrow).Value = call.DealState;
                }
                if (lastrow >= firstrow)
                {
                    wsheet.Cell("A" + firstrow).Value = m.Name;
                    wsheet.Range("A" + firstrow, "A" + lastrow).Merge();
                }
                firstrow = lastrow + 1;

            }
            var Rng = wsheet.Range("A1", "E" + lastrow);
            Rng.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            Rng.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            Rng.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            wsheet.Column(1).Width = 15;
            wsheet.Column(2).Width = 15;
            wsheet.Column(3).Width = 30;
            wsheet.Column(4).Width = 30;
            wsheet.Column(5).Width = 10;
            Rng.Style.Alignment.WrapText = true;
            wsheet.Range("A1", "E2").Style.Font.Bold = true;
            wsheet.Range("A3", "A" + lastrow).Style.Font.Bold = true;
            return wb;
        }
    }
}
