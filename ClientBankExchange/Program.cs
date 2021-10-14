using System;
using System.Windows.Forms;

namespace ClientBankExchange
{
    class Program
    {
        internal static int FILE_DATA_BEGINING = 1;

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
                while (tryConvertToString(sheet, row, 1, out string id))
                {
                    var docSection = new ExchangeDocumentSection { Id = id };
                    if (!tryConvertToString(sheet, row, 2, out docSection.Name) ||
                        !tryConvertToDate(sheet, row, 2, out docSection.Birthday) ||
                        !tryConvertToString(sheet, row, 2, out docSection.Court) ||
                        !tryConvertToString(sheet, row, 2, out docSection.CourtAddress) ||
                        !tryConvertToString(sheet, row, 2, out docSection.Recipient) ||
                        !tryConvertToString(sheet, row, 2, out docSection.KPP) ||
                        !tryConvertToString(sheet, row, 2, out docSection.INN) ||
                        !tryConvertToString(sheet, row, 2, out docSection.OKTMO) ||
                        !tryConvertToString(sheet, row, 2, out docSection.RecipientAccount) ||
                        !tryConvertToString(sheet, row, 2, out docSection.TreasuryAccount) ||
                        !tryConvertToString(sheet, row, 2, out docSection.Bank) ||
                        !tryConvertToString(sheet, row, 2, out docSection.City) ||
                        !tryConvertToString(sheet, row, 2, out docSection.PaymentType) ||
                        !tryConvertToString(sheet, row, 2, out docSection.SenderStatus) ||
                        !tryConvertToString(sheet, row, 2, out docSection.PaymentPurpose) ||
                        !tryConvertToString(sheet, row, 2, out docSection.BIK) ||
                        !tryConvertToString(sheet, row, 2, out docSection.KBK) ||
                        !tryConvertToString(sheet, row, 2, out docSection.AgreementNumber) ||
                        !tryConvertToCurrency(sheet, row, 2, out docSection.ClaimAmount) ||
                        !tryConvertToCurrency(sheet, row, 2, out docSection.StateDutyAmount))
                        Console.Out.WriteLine($"Ошибка. Повторение ID в основном файле: '{id}'. Строка {row1} пропускается.");
                    else
                        document.Sections.Add(docSection);
                    row++;
                }
                Console.Out.WriteLine($"Прочитано записей таблицы: {row - FILE_DATA_BEGINING}\n");

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
            Console.Out.WriteLine(" Дождитесь закрытия программы...");
        }

        private static bool tryConvertToString(Microsoft.Office.Interop.Excel.Worksheet sheet, int row, int col, out string value)
        {
            value = (sheet.Cells[row, col] as Microsoft.Office.Interop.Excel.Range)?.Value2?.ToString();
            return !string.IsNullOrWhiteSpace(value);
        }

        private static bool tryConvertToDouble(string str, out double value)
        {
            value = double.NaN;
            if (string.IsNullOrEmpty(str)) return false;

            bool hasComma = str.Contains(",");
            bool hasPoint = str.Contains(".");
            if (hasPoint)
            {
                if (hasComma)
                    str = str.Remove(str.IndexOf('.'), 1);
                else
                    str = str.Replace('.', ',');
            }

            if (double.TryParse(str, out value))
                return true;
            return false;
        }
    }
}
