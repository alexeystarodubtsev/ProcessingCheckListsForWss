using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Text.RegularExpressions;
using ProcessingCheckListWss.ProcessingCheckLists;

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
            Dictionary<string, string> months = new Dictionary<string, string>();
            months[new DateTime(2020, month, 1).AddMonths(-2).ToString("MMMM")] = "PrePreLastMonth";
            months[new DateTime(2020, month, 1).AddMonths(-1).ToString("MMMM")] = "PreLastMonth";
            months[new DateTime(2020, month, 1).ToString("MMMM")] = "LastMonth";
            Console.WriteLine("Посчитать: 1. Ежемесячную аналитику 2. Еженедельную. Введите цифру:");
            string opt = Console.ReadLine();
            
            ParseAllFiles(@"FilesToProccessing", months, opt);
        }
        private static void ParseAllFiles(string folder, Dictionary<string, string> folders, string opt)
        {
            Dictionary<string, Dictionary<string, List<DataForPrint>>> printPagesByMonth = new Dictionary<string, Dictionary<string, List<DataForPrint>>>();
            Dictionary<string, Dictionary<string, List<DataForPrint>>> printTotalManagers = new Dictionary<string, Dictionary<string, List<DataForPrint>>>();
            foreach (var Month in folders.Keys)
            {
                List<Manager> managers = new List<Manager>();

                printPagesByMonth[Month] = new Dictionary<string, List<DataForPrint>>();
                printTotalManagers[Month] = new Dictionary<string, List<DataForPrint>>();
                foreach (var file in Directory.GetFiles(folder + "\\" + folders[Month]))
                {
                    Manager m1 = new Manager(file);
                    managers.Add(m1);
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
                        //if (folders[Month] == "LastMonth")
                        //{

                        //    var wb = OutPutCheckList.FillAnalyticOfPoint(m1);
                        //    wb.SaveAs(@"Result\" + m1.Name + ".xlsx");
                        //}
                    }
                }
                if (folders[Month] == "LastMonth" && opt == "2")
                {
                    Console.WriteLine("Введите дату начала счета статистики");
                    string inputstr = Console.ReadLine();
                    DateTime firstDate;
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
