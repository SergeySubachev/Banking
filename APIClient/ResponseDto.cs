using System.Text.Json.Serialization;

namespace APIClient
{
    public class ResponseDto
    {
        [JsonPropertyName("result")]
        public ResponseResultDto Result { get; set; }
    }

    public class ResponseResultDto
    {
        [JsonPropertyName("court")]
        public CourtDto Court { get; set; }

        [JsonPropertyName("higher_court")]
        public CourtDto HigherCourt { get; set; }
    }

    public class CourtDto
    {
        [JsonPropertyName("code")]
        public string Code { get; set; }

        [JsonPropertyName("type")]
        public string CourtType { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("address")]
        public string Address { get; set; }

        [JsonPropertyName("phone")]
        public string Phone { get; set; }

        [JsonPropertyName("email")]
        public string Email { get; set; }

        [JsonPropertyName("website")]
        public string Website { get; set; }

        [JsonPropertyName("ufk.code")]
        public string UfkCode { get; set; }

        [JsonPropertyName("ufk.name")]
        public string UfkName { get; set; }

        [JsonPropertyName("ufk.inn")]
        public string UfkInn { get; set; }

        [JsonPropertyName("ufk.kpp")]
        public string UfkKpp { get; set; }

        [JsonPropertyName("ufk.kbk")]
        public string UfkKbk { get; set; }

        [JsonPropertyName("ufk.oktmo")]
        public string UfkOktmo { get; set; }

        [JsonPropertyName("ufk.bank")]
        public string UfkBank { get; set; }

        [JsonPropertyName("ufk.bik")]
        public string UfkBik { get; set; }

        [JsonPropertyName("ufk.account")]
        public string UfkAccount { get; set; }

        [JsonPropertyName("ufk.correspondent_account")]
        public string UfkCorrespondentAccount { get; set; }

        [JsonPropertyName("ufk.phone")]
        public string UfkPhone { get; set; }

        [JsonPropertyName("ufk.website")]
        public string UfkWebsite { get; set; }

        [JsonPropertyName("ufk.address")]
        public string UfkAddress { get; set; }

        [JsonPropertyName("fssp.name")]
        public string FsspName { get; set; }

        [JsonPropertyName("fssp.address")]
        public string FsspAddress { get; set; }

        [JsonPropertyName("fssp.phone")]
        public string FsspPhone { get; set; }
    }
}
