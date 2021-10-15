using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ClientBankExchange
{
    public class ExchangeDocument
    {
        public List<ExchangeDocumentSection> Sections = new List<ExchangeDocumentSection>();

        public override string ToString()
        {
            var strBuilder = new StringBuilder();
            strBuilder.AppendLine("1CClientBankExchange");
            strBuilder.AppendLine("ВерсияФормата=1.03");
            strBuilder.AppendLine("Кодировка=Windows");
            strBuilder.AppendLine("Отправитель=Бухгалтерия предприятия, редакция 3.0");
            strBuilder.AppendLine("Получатель=");
            strBuilder.AppendLine($"ДатаСоздания={DateTime.Today:dd.MM.yyyy}");
            strBuilder.AppendLine($"ВремяСоздания={DateTime.Now:T}");
            strBuilder.AppendLine($"ДатаНачала={DateTime.Today:dd/MM/yyyy}");
            strBuilder.AppendLine($"ДатаКонца={DateTime.Today:dd/MM/yyyy}");
            strBuilder.AppendLine("РасчСчет=40702810562100001617");
            strBuilder.AppendLine("Документ=Платежное поручение");
            foreach (var section in Sections)
            {
                strBuilder.AppendLine("СекцияДокумент=Платежное поручение");
                strBuilder.AppendLine($"Номер={section.Id}");
                strBuilder.AppendLine($"Дата={DateTime.Today:dd/MM/yyyy}");
                strBuilder.AppendLine($"Сумма={section.StateDutyAmount.ToString("0.00", CultureInfo.CreateSpecificCulture("en-US"))}");
                strBuilder.AppendLine("ПлательщикСчет=40702810562100001617");
                strBuilder.AppendLine("Плательщик=ИНН 6671021289 ООО КОЛЛЕКТОРСКОЕ АГЕНТСТВО \"ОРПЗ\"");
                strBuilder.AppendLine("ПлательщикИНН=6671021289");
                strBuilder.AppendLine("Плательщик1=ООО КОЛЛЕКТОРСКОЕ АГЕНТСТВО \"ОРПЗ\"");
                strBuilder.AppendLine("ПлательщикРасчСчет=40702810562100001617");
                strBuilder.AppendLine("ПлательщикБанк1=ПАО КБ \"УБРИР\"");
                strBuilder.AppendLine("ПлательщикБанк2=г Екатеринбург");
                strBuilder.AppendLine("ПлательщикБИК=046577795");
                strBuilder.AppendLine("ПлательщикКорсчет=30101810900000000795");
                strBuilder.AppendLine($"ПолучательСчет={section.TreasuryAccount}");
                strBuilder.AppendLine($"Получатель=ИНН {section.RecipientINN} {section.Recipient}");
                strBuilder.AppendLine($"ПолучательИНН={section.RecipientINN}");
                strBuilder.AppendLine($"Получатель1={section.Recipient}");
                strBuilder.AppendLine($"ПолучательРасчСчет={section.TreasuryAccount}");
                strBuilder.AppendLine($"ПолучательБанк1={section.Bank}");
                strBuilder.AppendLine($"ПолучательБанк2={section.City}");
                strBuilder.AppendLine($"ПолучательБИК={section.BIK}");
                strBuilder.AppendLine($"ПолучательКорсчет={section.RecipientAccount}");
                strBuilder.AppendLine($"ВидОплаты={section.PaymentType}");
                strBuilder.AppendLine($"СтатусСоставителя={section.SenderStatus}");
                strBuilder.AppendLine("ПлательщикКПП=667901001");
                strBuilder.AppendLine($"ПолучательКПП={section.RecipientKPP}");
                strBuilder.AppendLine($"ПоказательКБК={section.KBK}");
                strBuilder.AppendLine($"ОКАТО={section.RecipientOKTMO}");
                strBuilder.AppendLine("ПоказательОснования=ТП");
                strBuilder.AppendLine("ПоказательПериода=0");
                strBuilder.AppendLine("ПоказательНомера=0");
                strBuilder.AppendLine("ПоказательДаты=0");
                strBuilder.AppendLine("ПоказательТипа=");
                strBuilder.AppendLine("Очередность=5");
                strBuilder.AppendLine($"НазначениеПлатежа={section.PaymentPurpose} {section.Name}");
                strBuilder.AppendLine($"НазначениеПлатежа1={section.PaymentPurpose} {section.Name}");
                strBuilder.AppendLine("Код = 0");
                strBuilder.AppendLine("КонецДокумента");
            }
            strBuilder.AppendLine("КонецФайла");
            return strBuilder.ToString();
        }
    }
}
