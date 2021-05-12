using System.Text.Json.Serialization;

namespace TestProject
{
    class GitHubUserDto
    {
        [JsonPropertyName("login")]
        public string Login { get; set; }

        [JsonPropertyName("id")]
        //[JsonConverter(typeof(JsonConverter<string>))]
        public int Id { get; set; }

        [JsonPropertyName("url")]
        public string Url { get; set; }
    }
}
