using System;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace unified_taxonomy
{
    public class Product
    {
        private static int nextId = 0; 

        public int Id { get; init; }

        public Product()
        {
            Id = nextId++;
        }

        [JsonPropertyName("uid")]
        public Uri Uid { get; set; }

        [JsonPropertyName("level")]
        public long Level { get; set; }

        [JsonPropertyName("label")]
        public string Label { get; set; }

        [JsonPropertyName("slug")]
        public string Slug { get; set; }

        [JsonPropertyName("parentSlug")]
        public string ParentSlug { get; set; }

        [JsonPropertyName("createdAt")]
        public DateTimeOffset CreatedAt { get; set; }

        [JsonPropertyName("updatedAt")]
        public DateTimeOffset UpdatedAt { get; set; }

        [JsonPropertyName("updatedBy")]
        public Uri UpdatedBy { get; set; }
    
        public static Product[] FromJson(string json) => JsonSerializer.Deserialize<Product[]>(json);
        //public static string ToJson(this Product[] self) => JsonConvert.SerializeObject(self, QuickType.Converter.Settings);
    }
}
