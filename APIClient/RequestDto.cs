using System.Text.Json.Serialization;

namespace APIClient
{
    public class RequestDto
    {
        [JsonPropertyName("address")]
        public string Address { get; set; }

        [JsonPropertyName("get_fssp_geo")]
        public int Fssp { get; set; } = 0;

        [JsonPropertyName("debt_sum")]
        public int DebtSum { get; set; } = 500;

        //[JsonPropertyName("legal_type")]
        //public int LegalType { get; set; } = 0;
    }
}
