using RestSharp;
using System;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Windows.Forms;

namespace APIClient
{
    class Program
    {
        internal static int FILE_DATA_BEGINING = 2;

        [STAThread]
        static void Main(string[] args)
        {
            string exePath = Assembly.GetEntryAssembly().Location;
#if DEBUG
            string keyFileName = Path.Combine(Directory.GetParent(exePath).FullName, @"..\..\api-key.txt");
#else
            string keyFileName = Path.Combine(Directory.GetParent(exePath).FullName, "api-key.txt");
#endif
            if (!File.Exists(keyFileName))
                Console.Out.WriteLine($"Ошибка. Не найден ключ доступа к серверу '{keyFileName}'.");

            string apiKey = GetApiKey(keyFileName);

            var ofd = new OpenFileDialog { Title = "Открыть файл адресов" };
            if (ofd.ShowDialog() != DialogResult.OK)
                return;
            var excelFileName = ofd.FileName;

            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {
                var client = new RestClient("https://api.explorer.debex.ru/production/jurisdiction") { Timeout = -1 };

                excelApp = new Microsoft.Office.Interop.Excel.Application();
                var wb = excelApp.Workbooks.Open(excelFileName, false, false);
                var sheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[1];

                var dtoProps = typeof(CourtDto).GetProperties();
                for (int i = 0; i < dtoProps.Length; i++)
                    (sheet.Cells[1, 3 + i] as Microsoft.Office.Interop.Excel.Range).Value2 = dtoProps[i].Name;

                int row = FILE_DATA_BEGINING;
                string id = (sheet.Cells[row, 1] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString();
                while (!string.IsNullOrWhiteSpace(id))
                {
                    string address = (sheet.Cells[row, 2] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString();
                    if (string.IsNullOrWhiteSpace(address))
                    {
                        Console.Out.WriteLine($"Пустая ячейка адреса. Строка {row}. ID {id}.");
                    }
                    else
                    {
                        var request = new RestRequest(Method.POST);
                        request.AddHeader("x-api-key", apiKey);
                        request.AddParameter("address", address);
                        IRestResponse response = client.Execute(request);
                        if (response.StatusCode == System.Net.HttpStatusCode.OK)
                        {
                            var resultObj = JsonSerializer.Deserialize<ResponseDto>(response.Content);
                            for (int i = 0; i < dtoProps.Length; i++)
                            {
                                var val = dtoProps[i].GetValue(resultObj.Result.Court);
                                if (val != null)
                                    (sheet.Cells[row, 3 + i] as Microsoft.Office.Interop.Excel.Range).Value2 = val.ToString();
                            }
                            Console.Out.WriteLine($"OK");
                        }
                        else if (response.StatusCode == System.Net.HttpStatusCode.BadRequest)
                        {

                        }
                        else if (response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                        {

                        }
                        else
                        {

                        }
                    }

                    row++;
                    id = (sheet.Cells[row, 1] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString();
                }
                Console.Out.WriteLine($"Сохранение...");
                wb.Save(); wb.Close();
                Console.Out.WriteLine($"Данные сохранены.");
            }
            catch (Exception)
            {

                throw;
            }

            Console.Out.WriteLine("Нажмите любую клавишу");
            Console.ReadKey();
        }

        private static string GetApiKey(string keyFileName)
        {
            string res = File.ReadAllText(keyFileName);
            return res.Trim();
        }
    }
}
