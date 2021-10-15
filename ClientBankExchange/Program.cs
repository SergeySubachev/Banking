using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ClientBankExchange
{
    class Program
    {
        internal static int FILE_DATA_BEGINING = 2;

        [STAThread]
        static void Main(string[] args)
        {
            string xmlFilename;
            var ofd = new OpenFileDialog { Title = "Открыть файл Excel" };
            if (ofd.ShowDialog() != DialogResult.OK)
                return;
            xmlFilename = ofd.FileName;

            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                var wb = excelApp.Workbooks.Open(xmlFilename, false, false);
                Console.WriteLine($"Чтение данных...");
                var sheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets.Item[1];
                int row = FILE_DATA_BEGINING;
                var document = new ExchangeDocument();
                while (TryConvertToString(sheet, row, 1, out string id))
                {
                    var docSection = new ExchangeDocumentSection { Id = id };
                    if (!TryConvertToString(sheet, row, 2, out docSection.Name) ||
                        !TryConvertToDate(sheet, row, 3, out docSection.Birthday) ||
                        !TryConvertToString(sheet, row, 4, out docSection.Court) ||
                        !TryConvertToString(sheet, row, 5, out docSection.CourtAddress) ||
                        !TryConvertToString(sheet, row, 6, out docSection.Recipient) ||
                        !TryConvertToString(sheet, row, 7, out docSection.RecipientKPP) ||
                        !TryConvertToString(sheet, row, 8, out docSection.RecipientINN) ||
                        !TryConvertToString(sheet, row, 9, out docSection.RecipientOKTMO) ||
                        !TryConvertToString(sheet, row, 10, out docSection.RecipientAccount) ||
                        !TryConvertToString(sheet, row, 11, out docSection.TreasuryAccount) ||
                        !TryConvertToString(sheet, row, 12, out docSection.Bank) ||
                        !TryConvertToString(sheet, row, 13, out docSection.City) ||
                        !TryConvertToString(sheet, row, 14, out docSection.PaymentType) ||
                        !TryConvertToString(sheet, row, 15, out docSection.SenderStatus) ||
                        !TryConvertToString(sheet, row, 16, out docSection.PaymentPurpose) ||
                        !TryConvertToString(sheet, row, 17, out docSection.BIK) ||
                        !TryConvertToString(sheet, row, 18, out docSection.KBK) ||
                        !TryConvertToString(sheet, row, 19, out docSection.AgreementNumber) ||
                        !TryConvertToDouble(sheet, row, 20, out docSection.ClaimAmount) ||
                        !TryConvertToDouble(sheet, row, 21, out docSection.StateDutyAmount))
                        Console.Out.WriteLine($"Ошибка чтения данных. Строка {row} пропускается.");
                    else
                        document.Sections.Add(docSection);
                    row++;
                }
                wb.Close();
                Console.Out.WriteLine("");

                Console.Out.WriteLine($"Экспорт...");
                string txtFilename = Path.ChangeExtension(xmlFilename, "txt");
                File.WriteAllText(txtFilename, document.ToString(), Encoding.GetEncoding(1251));
                Console.Out.WriteLine($"Экпорт выполнен. Имя файла: \"{Path.GetFileName(txtFilename)}\".");
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
            Console.Out.WriteLine("Дождитесь закрытия программы...");
        }

        private static bool TryConvertToString(Microsoft.Office.Interop.Excel.Worksheet sheet, int row, int col, out string value)
        {
            value = (sheet.Cells[row, col] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString();
            if (!string.IsNullOrEmpty(value))
                value = value.Trim(new[] { ' ', '.' });
            return !string.IsNullOrEmpty(value);
        }

        private static bool TryConvertToDouble(Microsoft.Office.Interop.Excel.Worksheet sheet, int row, int col, out double value)
        {
            value = default;
            if (!TryConvertToString(sheet, row, col, out string stringValue))
            {
                Console.Out.Write($"Пустая ячейка (строка {row}, столбец {col}. ");
                return false;
            }

            bool hasComma = stringValue.Contains(",");
            bool hasPoint = stringValue.Contains(".");
            if (hasPoint)
            {
                if (hasComma)
                    stringValue = stringValue.Remove(stringValue.IndexOf('.'), 1);
                else
                    stringValue = stringValue.Replace('.', ',');
            }

            if (!double.TryParse(stringValue, out value))
            {
                Console.Out.Write($"Не удалось прочитать число \"{stringValue}\". Cтрока {row}, столбец {col}. ");
                return false;
            }
            return true;
        }

        private static bool TryConvertToDate(Microsoft.Office.Interop.Excel.Worksheet sheet, int row, int col, out DateTime value)
        {
            value = default;
            if (!TryConvertToString(sheet, row, col, out string stringValue))
            {
                Console.Out.Write($"Не удалось прочитать дату \"{stringValue}\". Cтрока {row}, столбец {col}. ");
                return false;
            }

            try
            {
                value = DateTime.FromOADate((double)(sheet.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Value2);
            }
            catch (ArgumentException)
            {
                Console.Out.Write($"Не удалось прочитать дату \"{stringValue}\". Cтрока {row}, столбец {col}. ");
                return false;
            }
            return true;
        }
    }
}
