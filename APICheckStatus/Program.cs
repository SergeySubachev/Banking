using RestSharp;
using System;

namespace APICheckStatus
{
    class Program
    {
        static void Main(string[] _)
        {
            Console.Out.WriteLine("== Статистика использования сервисов за последние 3 месяца ==\n");
            Console.Out.Write("Введите ключ ОРПЗ (Production): ");
            var key = Console.In.ReadLine().Trim();
            Console.Out.WriteLine(" ");

            //https://documenter.getpostman.com/view/407152/T1Dv7ZXg#8c193790-3ace-41b6-a409-6e18ae37eba1
            var client = new RestClient("https://api.explorer.debex.ru/production/usage") { Timeout = -1 };
            var request = new RestRequest(Method.GET);
            request.AddHeader("x-api-key", key);
            request.AddParameter("text/plain", "", ParameterType.RequestBody);
            var response = client.Execute(request);
            Console.Out.WriteLine(response.Content);
            Console.Out.WriteLine("\nНажмите любую клавишу");
            Console.ReadKey();
        }
    }
}
