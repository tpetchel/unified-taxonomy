using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace unified_taxonomy
{
    public class Unified
    {
        public Product Product { get; set; }

        public IEnumerable<MsService> MsServices { get; set; }

        public IEnumerable<MsProd> MsProds { get; set; }
    }
}