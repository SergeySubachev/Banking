using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace APIClient
{
    class Program
    {
        internal static int FILE_DATA_BEGINING = 2;

        static void Main(string[] args)
        {
            string exePath = Assembly.GetEntryAssembly().Location;
#if DEBUG
            string keyFileName = Path.Combine(Directory.GetParent(exePath).FullName, @"..\..\..\..\api-key.txt");
#else
            string keyFileName = Path.Combine(Directory.GetParent(exePath).FullName, "api-key.txt");
#endif
            if (!File.Exists(keyFileName))
                Console.Out.WriteLine($"Ошибка. Не найден ключ доступа к серверу '{keyFileName}'.");

            var ofd = new OpenFileDialog { Title = "Открыть файл адресов" };
            if (ofd.ShowDialog() != DialogResult.OK)
                return;
            var excelFileName = ofd.FileName;

            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                var wb = excelApp.Workbooks.Open(excelFileName, false, false);
                var sheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[1];
                int row = FILE_DATA_BEGINING;
                string id = (sheet.Cells[row, 1] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString();
                while (!string.IsNullOrWhiteSpace(id))
                {
                    string addressStr = (sheet.Cells[row, 2] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString();
                    if (string.IsNullOrWhiteSpace(addressStr))
                    {
                        Console.Out.WriteLine($"Пустая ячейка адреса. Строка {row}. ID {id}.");
                    }
                    else
                    {

                    }

                    row++;
                }
            }
            catch (Exception)
            {

                throw;
            }

            Console.Out.WriteLine("Нажмите любую клавишу");
            Console.ReadKey();
        }
    }
}
