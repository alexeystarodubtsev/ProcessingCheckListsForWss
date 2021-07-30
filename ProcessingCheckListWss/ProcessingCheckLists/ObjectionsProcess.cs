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
        public static XLWorkbook GetXLWorkbook(List<Manager> lm, DateTime firstDate, DateTime LastDate, XLWorkbook oldFile, string company = "default")
        {
            var wb = oldFile;
            var wsheet = wb.Worksheets.Contains("Возражения") ? wb.Worksheet("Возражения") : wb.AddWorksheet("Возражения");
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

            int TargetQDays = -1;

            if (company == "Высоцкий")
            {
                LastDateFact = firstDate;

                LastDate = firstDate.AddDays(6);

                LastDateFact = LastDate;

                TargetQDays = 0;
            }

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
                            if (Regex.Match(call.DealState, "работ", RegexOptions.IgnoreCase).Success)
                                wsheet.Cell("J" + lastrow).SetValue<string>(call.DateOfNext);
                        }

                    }
                }

            }

            wsheet.Cell("A1").Value = "Возражения по звонкам по " + LastDateFact.AddDays(TargetQDays).ToString("dd.MM");
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

            foreach (var page in wb.Worksheets)
            {
                var rg = page.RangeUsed();

                List<string> nameManager = new List<string>();
                List<string> tag = new List<string>();

                int idxstart = -1;

                for (int i = 1; i < rg.LastColumn().ColumnNumber(); i++)
                {
                    if (idxstart == -1)
                    {
                        if (page.Cell(2, i).GetString() == "Менеджер")
                        {
                            if (nameManager.Count() == 0) nameManager.Add(page.Cell(3, i).GetString());

                            for (int j = 3; j < rg.LastRow().RowNumber(); j++)
                            {
                                for (int k = 0; k < nameManager.Count(); k++)
                                {
                                    if (nameManager[k] == page.Cell(j, i).GetString()) break;
                                    else if (k == nameManager.Count() - 1) nameManager.Add(page.Cell(j, i).GetString());
                                }
                            }
                        }
                    }

                    idxstart = -2;

                    if (page.Cell(2, i).GetString() == "Тэг")
                    {
                        if (tag.Count() == 0) tag.Add(page.Cell(3, i).GetString());

                        for (int j = 3; j < rg.LastRow().RowNumber(); j++)
                        {
                            for (int k = 0; k < tag.Count(); k++)
                            {
                                if (tag[k] == page.Cell(j, i).GetString()) break;
                                else if (k == tag.Count() - 1) tag.Add(page.Cell(j, i).GetString());
                            }
                        }
                    }
                }

                for (int i = 1; i < rg.LastColumn().ColumnNumber(); i++)
                {
                    idxstart = i;

                    if (page.Cell(1, i).GetString() == "Статистика")
                    {
                        page.Cell(2, i).SetValue<string>("Менеджер");

                        i++;

                        List<string> calcFormul = new List<string>();

                        for (int j = 0; j < tag.Count(); j++)
                        {
                            for (int k = 0; k < nameManager.Count(); k++)
                            {                 
                                calcFormul.Add((string.Format("=COUNTIFS($E$3:$E$1000,\"{0}\",$A$3:$A$1000,{1}{2})",        
                                    tag[j],
                                    page.Cell(j + 3, idxstart).WorksheetColumn().ColumnLetter(),
                                    page.Cell(k + 3, idxstart).WorksheetRow().RowNumber())));
                            }            
                        }

                        page.Cell(3 + nameManager.Count(), idxstart).SetValue<string>("ИТОГО");

                        for (int j = 0; j < tag.Count(); j++)
                        {
                            page.Column(j + idxstart + 1).Width = 20;
                            page.Cell(2, j + idxstart + 1).SetValue<string>(tag[j]);
                        }

                        for (int j = 0; j < tag.Count(); j++)
                        {
                            page.Cell(nameManager.Count() + 3, j + idxstart + 1).SetFormulaA1((
                                string.Format(@"=SUM({0}{2}:{0}{1})", 
                                page.Cell(3, j + idxstart + 1).WorksheetColumn().ColumnLetter(), page.Cell(nameManager.Count() + 2, 
                                j + idxstart + 1).WorksheetRow().RowNumber(), page.Cell(3, j + idxstart + 1).WorksheetRow().RowNumber())));
                        }

                        for (int j = 0; j < nameManager.Count(); j++)
                        {
                            page.Cell(j + 3, idxstart).SetValue<string>(nameManager[j]);
                        }

                        for (int j = 0, q = 0; j < tag.Count(); j++)
                        {
                            for (int k = 0; k < nameManager.Count(); k++, q++) 
                            {
                                page.Cell(k + 3, j + idxstart + 1).FormulaA1 = calcFormul[q];
                            }
                        }
                        
                    }  
               
                }

                bool TF = false;

                if (TF == false)
                {
                    string buffer = (string.Format("{0}",
                                    page.Cell(1, idxstart + tag.Count() - 4).WorksheetColumn().ColumnLetter()));

                    rg = page.Range("L2", buffer + (nameManager.Count() + 3));

                    rg.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    rg.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    rg.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rg.Style.Alignment.WrapText = true;

                    page.Range("L2", "AS" + (nameManager.Count() + 3)).Style.Font.Bold = true;

                    TF = true;
                }
            }

            return wb;
        }
    }
}
