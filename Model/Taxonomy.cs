using System.Collections.Generic;

namespace unified_taxonomy
{
    public class Taxonomy
    {
        public List<Product> Products { get; set; } = new();

        public List<MsService> MsServices { get; set; } = new();

        public List<MsProd> MsProds { get; set; } = new();

        public Taxonomy(IEnumerable<Product> products, IEnumerable<MsService> msServices, IEnumerable<MsProd> msProds)
        {
            Products.AddRange(products);
            MsServices.AddRange(msServices);
            MsProds.AddRange(msProds);
        }

        public Taxonomy(Taxonomy other) : this(other.Products, other.MsServices, other.MsProds)
        {
        }
    }
}