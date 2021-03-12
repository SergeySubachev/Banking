using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

//https://razilov-code.ru/2017/12/13/microsoft-office-interop-excel/
//https://docs.microsoft.com/ru-ru/previous-versions/office/troubleshoot/office-developer/automate-excel-from-visual-c

namespace Banking
{
    class Program
    {
        internal static int FILE_1_DATA_BEGINING = 3;
        internal static int FILE_2_DATA_BEGINING = 2;

        [STAThread]
        static void Main(string[] args)
        {
            string filename1, filename2;
            var ofd = new OpenFileDialog { Title = "Открыть основной файл" };
            if (ofd.ShowDialog() != DialogResult.OK)
                return;
            filename1 = ofd.FileName;

            ofd.Title = "Открыть файл поступлений";
            if (ofd.ShowDialog() != DialogResult.OK)
                return;
            filename2 = ofd.FileName;

            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {
                var persons = new Dictionary<string, Person>();

                excelApp = new Microsoft.Office.Interop.Excel.Application();
                var wb1 = excelApp.Workbooks.Open(filename1, false, false);
                Console.WriteLine($"Количество вкладок основного файла: {wb1.Sheets.Count}.");
                for (int i = 1; i <= wb1.Sheets.Count; i++)
                {
                    Console.WriteLine($"Чтение вкладки {i}...");
                    var sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)wb1.Sheets.Item[i];
                    int row1 = FILE_1_DATA_BEGINING;
                    string id = (sheet1.Cells[row1, 4] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString();
                    while (!string.IsNullOrWhiteSpace(id))
                    {
                        if (persons.ContainsKey(id))
                            Console.Out.WriteLine($"Ошибка. Повторение ID в основном файле: '{id}'. Строка {row1} пропускается.");
                        else
                        {
                            string balance = (sheet1.Cells[row1, 5] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString();
                            double? initbalance = convertToDouble(balance);
                            if (initbalance.HasValue)
                                persons.Add(id, new Person(id, initbalance.Value, i, row1));
                            else
                                Console.Out.WriteLine($"Ошибка. Не удалось прочитать остаток долга по ID '{id}'. Строка {row1}.");
                        }
                        row1++;
                        id = (sheet1.Cells[row1, 4] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString();
                    }

                    Console.Out.WriteLine($"Прочитано записей вкладки {i}: {row1 - FILE_1_DATA_BEGINING}\n");
                }

                var wb2 = excelApp.Workbooks.Open(filename2, false, false);
                var sheet2 = (Microsoft.Office.Interop.Excel.Worksheet)wb2.Sheets.Item[1];
                int row2 = FILE_2_DATA_BEGINING;
                Console.WriteLine($"Чтение файла поступлений...");
                string id2 = (sheet2.Cells[row2, 3] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                while (!string.IsNullOrWhiteSpace(id2))
                {
                    if (persons.ContainsKey(id2))
                    {
                        DateTime date = DateTime.FromOADate((double)(sheet2.Cells[row2, 4] as Microsoft.Office.Interop.Excel.Range).Value2);
                        double? cost = convertToDouble((sheet2.Cells[row2, 5] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString());
                        if (cost.HasValue)
                        {
                            var person = persons[id2];
                            person.Costs.Add((date.Month, cost.Value));
                            person.Balance -= cost.Value;
                            (sheet2.Cells[row2, 6] as Microsoft.Office.Interop.Excel.Range).Value2 = person.Balance;
                            (wb1.Sheets.Item[person.SheetNumber].Cells[person.Row, 24 + date.Month] as Microsoft.Office.Interop.Excel.Range).Value2 = cost.Value;
                        }
                        else
                            Console.Out.WriteLine($"Ошибка. Не удалось прочитать сумму поступления по ID '{id2}'. Строка {row2}.");
                    }
                    else
                        Console.Out.WriteLine($"Ошибка. В основном файле нет записи с ID '{id2}'. Строка {row2}.");

                    row2++;
                    id2 = (sheet2.Cells[row2, 3] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString();
                }
                var allCosts = persons.Values.SelectMany(p => p.Costs);
                Console.Out.WriteLine($"Всего прочитано поступлений: {allCosts.Count()} на сумму {allCosts.Select(c => c.sum).Sum():0.## р.}\n");

                foreach (var person in persons.Values)
                    (wb1.Sheets.Item[person.SheetNumber].Cells[person.Row, 6] as Microsoft.Office.Interop.Excel.Range).Value2 = person.Balance;

                Console.Out.WriteLine($"Сохранение...");
                wb1.Save(); wb1.Close();
                wb2.Save(); wb2.Close();
                Console.Out.WriteLine($"Данные о поступлениях сохранены.");
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.Message);
                Console.Out.WriteLine("Работа программы завершена.");
            }
            finally
            {
                excelApp.Quit();
            }

            Console.Out.WriteLine("Нажмите любую клавишу");
            Console.ReadKey();
        }

        private static double? convertToDouble(string balance)
        {
            if (string.IsNullOrEmpty(balance))
                return null;

            bool hasComma = balance.Contains(",");
            bool hasPoint = balance.Contains(".");
            if (hasPoint)
            {
                if (hasComma)
                    balance = balance.Remove(balance.IndexOf('.'), 1);
                else
                    balance = balance.Replace('.', ',');
            }

            if (double.TryParse(balance, out double res))
                return res;
            else
                return null;
        }
    }
}
