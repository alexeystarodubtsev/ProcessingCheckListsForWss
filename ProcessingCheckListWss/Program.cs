﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Text.RegularExpressions;
using ProcessingCheckListWss.ProcessingCheckLists;
using System.Globalization;

namespace ProcessingCheckListWss
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello");
            Directory.CreateDirectory("Result");
            Console.WriteLine("Введите номер последнего месяца, обрабатываемых файлов");
            int month;
            Int32.TryParse(Console.ReadLine(), out month);
            
            Console.WriteLine("Введите название компании");
            string Company = Console.ReadLine();
            Console.WriteLine("Посчитать: 1. Ежемесячную аналитику 2. Еженедельную. Введите цифру:");
            string opt = Console.ReadLine();
            ParseAllFiles(@"FilesToProccessing", month, opt,  Company);
        }
        private static void ParseAllFiles(string folder, int numMonth, string opt, string Company)
        {
            DateTimeFormatInfo info = CultureInfo.GetCultureInfo("ru-RU").DateTimeFormat;

            Dictionary<string, string> folders = new Dictionary<string, string>();
            folders[info.MonthNames[(numMonth + 9) % 12]] = "PrePreLastMonth";
            folders[info.MonthNames[(numMonth + 10) % 12]] = "PreLastMonth";
            folders[info.MonthNames[(numMonth - 1) % 12]] = "LastMonth";
            Dictionary<string, Dictionary<string, List<DataForPrint>>> printPagesByMonth = new Dictionary<string, Dictionary<string, List<DataForPrint>>>();
            Dictionary<string, Dictionary<string, List<DataForPrint>>> printTotalManagers = new Dictionary<string, Dictionary<string, List<DataForPrint>>>();
            List<Manager> allMonthManagers = new List<Manager>();
            foreach (var Month in folders.Keys)
            {
                List<Manager> managers = new List<Manager>();

                printPagesByMonth[Month] = new Dictionary<string, List<DataForPrint>>();
                printTotalManagers[Month] = new Dictionary<string, List<DataForPrint>>();
                foreach (var file in Directory.GetFiles(folder + "\\" + folders[Month]))
                {
                    Manager m1;
                    if (Company == "РНР")
                    {
                        ProcessingCheckListRNRHause m2;
                        m2 = new ProcessingCheckListRNRHause(file,Month);
                        m2.Processing();
                        allMonthManagers.Add(m2);
                        managers.Add(m2);
                        m1 = m2;
                    }
                    else
                    {
                        if (Company == "Аверс")
                        {
                            ProcessingCkeckListAvers m2;
                            m2 = new ProcessingCkeckListAvers(file, Month);
                            m2.Processing();
                            allMonthManagers.Add(m2);
                            managers.Add(m2);
                            m1 = m2;
                        }
                        else
                        {
                            if (Company == "Белфан")
                            {
                                ProcessingBelfanCheckList m2;
                                m2 = new ProcessingBelfanCheckList(file, Month);
                                m2.Processing();
                                allMonthManagers.Add(m2);
                                managers.Add(m2);
                                m1 = m2;
                            }
                            else
                            {
                                Manager m2;
                                m2 = new Manager(file, Month);
                                m2.Processing();
                                allMonthManagers.Add(m2);
                                managers.Add(m2);
                                m1 = m2;
                            }
                        }
                    }
                    
                    if (opt == "1")
                    {
                        
                        Dictionary<string, DataForPrint> AnalyticManager = m1.getDataByStage();
                        Dictionary<string, DataForPrint> AnalyticManagerTotal = new Dictionary<string, DataForPrint>();
                        AnalyticManagerTotal["ИТОГО"] = new DataForPrint(m1);
                        //Add(new DataForPrint(m1));
                        foreach (var page in AnalyticManager.Keys)
                        {
                            if (!printPagesByMonth[Month].ContainsKey(page))
                            {
                                printPagesByMonth[Month][page] = new List<DataForPrint>();
                            }
                            printPagesByMonth[Month][page].Add(AnalyticManager[page]);
                        }
                        if (!printTotalManagers[Month].ContainsKey("ИТОГО"))
                        {
                            printTotalManagers[Month]["ИТОГО"] = new List<DataForPrint>();
                        }

                        printTotalManagers[Month]["ИТОГО"].Add(AnalyticManagerTotal["ИТОГО"]);
                        if (folders[Month] == "LastMonth")
                        {

                            var wb = OutPutCheckList.FillAnalyticOfPoint(m1,allMonthManagers.Where(m => m.Name == m1.Name && folders[m.month] == "PreLastMonth").FirstOrDefault());
                            wb.SaveAs(@"Result\" + m1.Name + ".xlsx");
                        }
                    }
                } 
                if (folders[Month] == "LastMonth" && opt == "2")
                {
                    Console.WriteLine("Введите дату начала счета статистики");
                    string inputstr = Console.ReadLine();
                    DateTime firstDate;
                    CultureInfo.GetCultureInfo("ru-RU");
                    DateTime.TryParse(inputstr, out firstDate);
                    var wb = OutPutCheckList.getStatistic(managers, firstDate);
                    wb.SaveAs(@"Result\Еженедельная статистика.xlsx");
                }

            }
            if (opt == "1")
            { 
                OutPutDoc doc = new OutPutDoc(printPagesByMonth);
                doc.getWb().SaveAs(@"Result\По этапам.xlsx");
                doc = new OutPutDoc(printTotalManagers);
                doc.getWb().SaveAs(@"Result\Итоговая.xlsx");
            }
        }

    }
}
