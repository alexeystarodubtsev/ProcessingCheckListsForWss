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
        public static XLWorkbook GetXLWorkbook(List<Manager> lm, DateTime firstDate, DateTime LastDate)
        {
            var wb =  new XLWorkbook();
            var wsheet = wb.AddWorksheet("Возражения");
            List<Call> calls = new List<Call>();
            //wsheet.Cell("A1").Value = "Возражения по звонкам с " + firstDate.ToString("dd.MM") + " по " + LastDate.ToString("dd.MM");
            wsheet.Range("A1", "F1").Merge();
            wsheet.Cell("A2").Value = "Менеджер";
            wsheet.Cell("B2").Value = "Этап";
            wsheet.Cell("C2").Value = "Клиент";
            wsheet.Cell("D2").Value = "Дата звонка";
            wsheet.Cell("E2").Value = "Возражение";
            wsheet.Cell("F2").Value = "Отработал ли менеджер";
            wsheet.Cell("G2").Value = "Как отработал";
            wsheet.Cell("H2").Value = "Результат";
            wsheet.Cell("I2").Value = "Дата назначенного контакта";
            int firstrow = 3;
            int lastrow = 2;

            var LastDateFact = firstDate.AddDays(1);
            foreach (var m in lm)
            {
                if (m.getLastDate() > LastDateFact && LastDate != firstDate)
                    LastDateFact = m.getLastDate();
                foreach (var stage in m.getStages())
                {
                    foreach (var call in stage.calls.Where(c => c.dateOfCall >= firstDate && c.dateOfCall <= LastDate && c.Objections != "" && c.Objections.Trim().ToLower() != "нет"))
                    {
                        lastrow++;
                        wsheet.Cell("C" + lastrow).Value = call.client;
                        if (call.ClientLink != "")
                        {
                            wsheet.Cell("C" + lastrow).Hyperlink = new XLHyperlink(call.ClientLink);
                        }
                        wsheet.Cell("D" + lastrow).SetValue<string>(call.dateOfCall.ToString("dd.MM.yyyy"));
                        wsheet.Cell("E" + lastrow).Value = call.Objections;
                        wsheet.Cell("F" + lastrow).Value = call.doneObjection;
                        wsheet.Cell("G" + lastrow).Value = call.howProcessObj;
                        wsheet.Cell("H" + lastrow).Value = call.DealState;
                        wsheet.Cell("B" + lastrow).Value = stage.name;
                        wsheet.Cell("I" + lastrow).SetValue<string>(call.DateOfNext); 
                    }
                }
                //var listcalls = m.GetCalls().Where(c => c.dateOfCall >= firstDate && c.dateOfCall <= LastDate && c.Objections != "" && c.Objections.Trim().ToLower() != "нет");
                //foreach (var call in listcalls)
                //{
                //    lastrow++;
                //    wsheet.Cell("C" + lastrow).Value = call.client;
                //    if (call.ClientLink != "")
                //    {
                //        wsheet.Cell("C" + lastrow).Hyperlink = new XLHyperlink( call.ClientLink);
                //    }
                //    wsheet.Cell("D" + lastrow).SetValue<string>(call.dateOfCall.ToString("dd.MM.yyyy"));
                //    wsheet.Cell("E" + lastrow).Value = call.Objections;
                //    wsheet.Cell("F" + lastrow).Value = call.howProcessObj;
                //    wsheet.Cell("G" + lastrow).Value = call.DealState;
                //}
                if (lastrow >= firstrow)
                {
                    wsheet.Cell("A" + firstrow).Value = m.Name;
                    wsheet.Range("A" + firstrow, "A" + lastrow).Merge();
                }
                firstrow = lastrow + 1;

            }
            wsheet.Cell("A1").Value = "Возражения по звонкам с " + firstDate.ToString("dd.MM") + " по " + LastDateFact.AddDays(-1).ToString("dd.MM");
            var Rng = wsheet.Range("A1", "I" + lastrow);
            Rng.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            Rng.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            Rng.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            wsheet.Column(1).Width = 15;
            wsheet.Column(2).Width = 15;
            wsheet.Column(3).Width = 15;
            wsheet.Column(4).Width = 10;
            wsheet.Column(5).Width = 30;
            wsheet.Column(6).Width = 15;
            wsheet.Column(7).Width = 30;
            wsheet.Column(8).Width = 10;
            wsheet.Column(9).Width = 10;
            Rng.Style.Alignment.WrapText = true;
            wsheet.Range("A1", "I2").Style.Font.Bold = true;
            wsheet.Range("A3", "A" + lastrow).Style.Font.Bold = true;

            List<string> tags = new List<string>();
            tags.Add("Цена");
            tags.Add("Сроки");
            tags.Add("Оплата");
            tags.Add("Конкуренты");
            wsheet.Cell("L1").Value = "Статистика";

            wsheet.Range(1,12,1,tags.Count + 13).Merge();
            for (int i = 0; i < tags.Count; i++)
            {
                wsheet.Cell(2,i + 13).Value = tags[i];
            }
            wsheet.Cell(2, 13 + tags.Count).Value = "Другое";

            int it = 2;
            wsheet.Cell("L" + it).Value = "Менеджер";
            foreach (var m in lm)
            {
                it++;
                wsheet.Cell("L" + it).Value = m.Name;
                var listcalls = m.GetCalls().Where(c => c.dateOfCall >= firstDate && c.dateOfCall <= LastDate && c.Objections != "" && c.Objections.Trim().ToLower() != "нет");
                
                for (int i = 0; i < tags.Count; i++)
                {
                    string tag = wsheet.Cell(2, i + 13).GetString();
                    wsheet.Cell(it,i+13).SetValue<int>(listcalls.Where(c => Regex.Match(c.Objections, tag, RegexOptions.IgnoreCase).Success || (tag=="Цена" && Regex.Match(c.Objections, "Дорого", RegexOptions.IgnoreCase).Success)).Count());
                    
                }
                wsheet.Cell(it, 13 + tags.Count).SetValue<int>(listcalls.Where(c => !tags.Where(tag => Regex.Match(c.Objections,tag, RegexOptions.IgnoreCase).Success || (tag == "Цена" && Regex.Match(c.Objections, "Дорого", RegexOptions.IgnoreCase).Success)).Any()).Count());

            }
            it++;
            wsheet.Cell("L" + it).Value = "Итого";
            for (int i = 0; i < tags.Count + 1; i++)
            {
                int qty = 0;
                for (int j = 3; j < it; j++)
                {
                    qty += wsheet.Cell(j, 13 + i).GetValue<int>();
                }
                wsheet.Cell(it, 13 + i).Value = qty;

            }
            Rng = wsheet.Range(1,12,it, 13 + tags.Count);
            Rng.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            Rng.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            Rng.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


            wsheet.Column(12).Width = 20;
            wsheet.Columns(13, 13 + tags.Count).Width = 20;
            Rng.Style.Alignment.WrapText = true;
            wsheet.Range("L1", "L" + it).Style.Font.Bold = true;
            wsheet.Range(2,12,2, 13 + tags.Count).Style.Font.Bold = true;
            wsheet.Range(it, 12, it, 13 + tags.Count).Style.Font.Bold = true;

            return wb;
        }
    }
}
