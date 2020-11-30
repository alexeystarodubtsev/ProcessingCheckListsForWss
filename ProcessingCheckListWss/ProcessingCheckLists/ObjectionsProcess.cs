using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ProcessingCheckListWss.ProcessingCheckLists
{
    class ObjectionsProcess
    {
        public static XLWorkbook GetXLWorkbook(List<Manager> lm, DateTime firstDate, DateTime LastDate, XLWorkbook oldFile)
        {
            var wb = oldFile;
            var wsheet = wb.Worksheets.Contains("Возражения") ?  wb.Worksheet("Возражения") : wb.AddWorksheet("Возражения");
            List<Call> calls = new List<Call>();
            wsheet.Cell("A1").Value = "Возражения по звонкам по " + LastDate.ToString("dd.MM");
            wsheet.Range("A1", "I1").Merge();
            wsheet.Cell("A2").Value = "Менеджер";
            wsheet.Cell("B2").Value = "Этап";
            wsheet.Cell("C2").Value = "Клиент";
            wsheet.Cell("D2").Value = "Дата звонка";
            wsheet.Cell("E2").Value = "Тэг";
            wsheet.Cell("F2").Value = "Возражение";
            wsheet.Cell("G2").Value = "Отработал ли менеджер";
            wsheet.Cell("H2").Value = "Как отработал";
            wsheet.Cell("I2").Value = "Результат";
            wsheet.Cell("J2").Value = "Дата назначенного контакта";
            int lastrow = wsheet.RangeUsed().LastRow().RowNumber();
            var LastDateFact = firstDate.AddDays(1);
            foreach (var m in lm)
            {
                if (m.getLastDate() > LastDateFact && LastDate != firstDate)
                    LastDateFact = m.getLastDate();
                foreach (var stage in m.getStages())
                {
                    foreach (var call in stage.calls.Where(c => c.dateOfCall >= firstDate && c.dateOfCall <= LastDate && c.Objections != "" && c.Objections.Trim().ToLower() != "нет"))
                    {
                        
                        var tagsMatch = Regex.Match(call.Objections, @"([\s,\w]*)\.", RegexOptions.IgnoreCase);
                        var tags = tagsMatch.Groups[0].Value.Trim('.').Split(',');
                        
                        foreach (var tag in tags.Where(t => t != ""))
                        {
                            lastrow++;
                            wsheet.Cell("A" + lastrow).Value = m.Name;
                            wsheet.Cell("C" + lastrow).Value = call.client;

                            if (call.ClientLink != "")
                            {
                                wsheet.Cell("C" + lastrow).Hyperlink = new XLHyperlink(call.ClientLink);
                            }
                            wsheet.Cell("D" + lastrow).SetValue<string>(call.dateOfCall.ToString("dd.MM.yyyy"));
                            wsheet.Cell("E" + lastrow).Value = (tag.Trim(' ').Substring(0, 1).ToUpper() + tag.Trim(' ').Substring(1, tag.Trim(' ').Length - 1).ToLower());
                            wsheet.Cell("F" + lastrow).Value = call.Objections;
                            wsheet.Cell("G" + lastrow).Value = call.doneObjection;
                            wsheet.Cell("H" + lastrow).Value = call.howProcessObj;
                            wsheet.Cell("I" + lastrow).Value = call.DealState;
                            wsheet.Cell("B" + lastrow).Value = stage.name;
                            wsheet.Cell("J" + lastrow).SetValue<string>(call.DateOfNext);
                        }
                        
                    }
                }

            }
            wsheet.Cell("A1").Value = "Возражения по звонкам по " + LastDateFact.AddDays(-1).ToString("dd.MM");
            var Rng = wsheet.Range("A1", "J" + lastrow);
            Rng.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            Rng.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            Rng.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            wsheet.Column(1).Width = 15;
            wsheet.Column(2).Width = 15;
            wsheet.Column(3).Width = 15;
            wsheet.Column(4).Width = 10;
            wsheet.Column(5).Width = 20;
            wsheet.Column(6).Width = 30;
            wsheet.Column(7).Width = 15;
            wsheet.Column(8).Width = 30;
            wsheet.Column(9).Width = 10;
            wsheet.Column(10).Width = 10;
            Rng.Style.Alignment.WrapText = true;
            wsheet.Range("A1", "J2").Style.Font.Bold = true;
            wsheet.Range("A3", "A" + lastrow).Style.Font.Bold = true;

            return wb;
        }
    }
}
