using System;

namespace ClientBankExchange
{
    /// <summary>
    /// СекцияДокумент=Платежное поручение
    /// </summary>
    public class ExchangeDocumentSection
    {
        public string
            Id,
            Name,
            Court, //суд
            CourtAddress, //адрес суда
            Recipient, //Наименование получателя платежа
            RecipientKPP, //КПП
            RecipientINN, //ИНН
            RecipientOKTMO, //ОКТМО
            RecipientAccount, //Номер счета получателя платежа
            TreasuryAccount, //номер казначейского счета
            Bank, //Наименование банка
            City, //Город
            PaymentType, //Вид оплаты
            SenderStatus, //Статус составителя
            PaymentPurpose, //Назначение платежа
            BIK, //БИК
            KBK, //Код бюджетной классификации(КБК)
            AgreementNumber; //Номер кредитного договора

        public DateTime 
            Birthday;

        public double
            ClaimAmount, //Сумма иска
            StateDutyAmount; //Сумма ГП
    }
}
