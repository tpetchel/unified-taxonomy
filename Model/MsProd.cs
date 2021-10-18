using System;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace unified_taxonomy
{
    public class MsProd
    {
        private static int nextId = 0; 

        public int Id { get; init; }

        public MsProd()
        {
            Id = nextId++;
        }

        [JsonPropertyName("uid")]
        public Uri Uid { get; set; }

        [JsonPropertyName("pillar")]
        public string Pillar { get; set; }

        [JsonPropertyName("product")]
        public string Product { get; set; }

        [JsonPropertyName("msProduct")]
        public string MsProduct { get; set; }

        [JsonPropertyName("isDashboard")]
        public bool IsDashboard { get; set; }

        [JsonPropertyName("active")]
        public bool Active { get; set; }

        [JsonPropertyName("createdAt")]
        public DateTimeOffset CreatedAt { get; set; }

        [JsonPropertyName("updatedAt")]
        public DateTimeOffset UpdatedAt { get; set; }

        [JsonPropertyName("updatedBy")]
        public Uri UpdatedBy { get; set; }

        [JsonPropertyName("technology")]
        public string Technology { get; set; }

        [JsonPropertyName("msTechnology")]
        public string MsTechnology { get; set; }

        [JsonPropertyName("manager")]
        public string Manager { get; set; }

        public static MsProd[] FromJson(string json) => JsonSerializer.Deserialize<MsProd[]>(json);
        //public static string ToJson(this MsProd[] self) => JsonConvert.SerializeObject(self, QuickType.Converter.Settings);
    }
}