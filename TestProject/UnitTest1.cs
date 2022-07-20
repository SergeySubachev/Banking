using ApprovalTests;
using ApprovalTests.Reporters;
using FluentAssertions;
using NUnit.Framework;
using RestSharp;
using System.Text.Json;

namespace TestProject
{
    public class Tests
    {
        private string apiKey = "";

        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void GitHubZenTest()
        {
            //Without RestSharp:
            //https://docs.microsoft.com/ru-ru/aspnet/web-api/overview/advanced/calling-a-web-api-from-a-net-client
            //https://github.com/SergeySubachev/web-game.git

            var client = new RestClient("https://api.github.com/zen");
            var request = new RestRequest(Method.GET);

            var response = client.Execute(request);
            response.StatusCode.Should().Be(System.Net.HttpStatusCode.OK);
        }

        [Test]
        [UseReporter(typeof(DiffReporter))]
        public void GitHubUserTest()
        {
            var client = new RestClient("https://api.github.com/users/SergeySubachev");
            var request = new RestRequest(Method.GET);

            var response = client.Execute(request);
            response.StatusCode.Should().Be(System.Net.HttpStatusCode.OK);

            Approvals.VerifyJson(response.Content);

            var user = JsonSerializer.Deserialize<GitHubUserDto>(response.Content);
            user.Login.Should().Be("SergeySubachev");
            user.Id.Should().Be(56560526);
            user.Url.Should().Be("https://api.github.com/users/SergeySubachev");
        }

        [Test]
        public void UsageTest()
        {
            //https://documenter.getpostman.com/view/407152/T1Dv7ZXg#8c193790-3ace-41b6-a409-6e18ae37eba1
            var client = new RestClient("https://api.explorer.debex.ru/production/usage") { Timeout = -1 };
            var request = new RestRequest(Method.GET);
            request.AddHeader("x-api-key", apiKey);
            request.AddParameter("text/plain", "", ParameterType.RequestBody);
            var response = client.Execute(request);
            response.StatusCode.Should().Be(System.Net.HttpStatusCode.OK);
        }

        [Test]
        public void DemoSampleTest()
        {
            //https://documenter.getpostman.com/view/407152/T1Dv7ZXg#8c193790-3ace-41b6-a409-6e18ae37eba1
            var client = new RestClient("https://api.explorer.debex.ru/production/jurisdiction") { Timeout = -1 };
            var request = new RestRequest(Method.POST);
            request.AddHeader("x-api-key", "productionKJdsnvkjnlekfnsdokfnj32n23523");
            request.AddParameter("address", "Екатеринбург, Комсомольская, 67");
            var response = client.Execute(request);
            response.StatusCode.Should().Be(System.Net.HttpStatusCode.OK);
        }

        [Test]
        public void DeserializeTest()
        {
            var client = new RestClient("https://api.explorer.debex.ru/production/jurisdiction") { Timeout = -1 };
            var request = new RestRequest(Method.POST);
            request.AddHeader("x-api-key", apiKey);
            request.AddParameter("address", "Екатеринбург, Комсомольская, 67");
            var response = client.Execute(request);

            var result = JsonSerializer.Deserialize<APIClient.ResponseDto>(response.Content);
            result.Should().NotBeNull();
        }
    }
}