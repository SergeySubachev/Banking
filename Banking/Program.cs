﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

//https://razilov-code.ru/2017/12/13/microsoft-office-interop-excel/
//https://docs.microsoft.com/ru-ru/previous-versions/office/troubleshoot/office-developer/automate-excel-from-visual-c

namespace Banking
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string filename1, filename2;
            var ofd = new OpenFileDialog();
            ofd.Title = "Открыть основной файл";
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
                var wb1 = excelApp.Workbooks.Open(filename1);
                var sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)wb1.Sheets.Item[1];
                int row1 = 2;
                string id = (sheet1.Cells[row1, 4] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                while (!string.IsNullOrWhiteSpace(id))
                {
                    //Console.Out.WriteLine(row.ToString());
                    if (persons.ContainsKey(id))
                        Console.Out.WriteLine($"Ошибка. Повторение ID в основном файле: '{id}'");
                    else
                    {
                        string balance = (sheet1.Cells[row1, 5] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString();
                        double? initbalance = convertToDouble(balance);
                        if (initbalance.HasValue)
                            persons.Add(id, new Person(id, initbalance.Value, row1));
                        else
                            Console.Out.WriteLine($"Ошибка. Не удалось прочитать баланс по ID '{id}'. Строка {row1}.");
                    }
                    row1++;
                    id = (sheet1.Cells[row1, 4] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString();
                }

                Console.Out.WriteLine($"Прочитано записей основного файла: {row1 - 1}\n");

                var wb2 = excelApp.Workbooks.Open(filename2);
                var sheet2 = (Microsoft.Office.Interop.Excel.Worksheet)wb2.Sheets.Item[1];
                int row2 = 2;
                id = (sheet2.Cells[row2, 3] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                while (!string.IsNullOrWhiteSpace(id))
                {
                    if (persons.ContainsKey(id))
                    {
                        DateTime date = DateTime.FromOADate((double)(sheet2.Cells[row2, 4] as Microsoft.Office.Interop.Excel.Range).Value2);
                        double? cost = convertToDouble((sheet2.Cells[row2, 5] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString());
                        if (cost.HasValue)
                            persons[id].Costs.Add((date.Month, cost.Value));
                        else
                            Console.Out.WriteLine($"Ошибка. Не удалось прочитать сумму поступления по ID '{id}'. Строка {row2}.");
                    }
                    else
                        Console.Out.WriteLine($"Ошибка. В основном файле нет записи с ID '{id}'. Строка {row2}.");

                    row2++;
                    id = (sheet2.Cells[row2, 3] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString();
                }
                Console.Out.WriteLine($"Прочитано записей файла поступлений: {row2 - 1}\n");

                foreach (var person in persons.Values)
                {
                    row1 = person.Row;
                    foreach (var (month, sum) in person.Costs)
                        (sheet1.Cells[row1, 24 + month] as Microsoft.Office.Interop.Excel.Range).Value = sum;
                }

                wb1.Close(true);
                wb2.Close();
                Console.Out.WriteLine($"Данные о поступлениях сохранены");
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.Message);
                Console.Out.WriteLine("Работа программы завершена");
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
