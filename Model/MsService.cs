using System;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace unified_taxonomy
{
    public class MsService
    {
        private static int nextId = 0; 

        public int Id { get; init; }

        public MsService()
        {
            Id = nextId++;
        }

        [JsonPropertyName("uid")]
        public Uri Uid { get; set; }

        [JsonPropertyName("pillar")]
        public string Pillar { get; set; }

        [JsonPropertyName("serviceArea")]
        public string ServiceArea { get; set; }

        [JsonPropertyName("service")]
        public string Service { get; set; }

        [JsonPropertyName("msService")]
        public string MsService_ { get; set; }

        [JsonPropertyName("subService")]
        public string SubService { get; set; }

        [JsonPropertyName("msSubService")]
        public string MsSubService { get; set; }

        [JsonPropertyName("manager")]
        public string Manager { get; set; }

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
    
        public static MsService[] FromJson(string json) => JsonSerializer.Deserialize<MsService[]>(json);
        //public static string ToJson(this MsService[] self) => JsonConvert.SerializeObject(self, QuickType.Converter.Settings);
    }
}