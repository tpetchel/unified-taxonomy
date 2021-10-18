using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using ClosedXML.Excel;

namespace unified_taxonomy
{
    class Program
    {
        static void Main(bool pull)
        {
            if (pull)
            {
                PullLatestTaxonomy();
            }

            Taxonomy everything = new(
                Product.FromJson(File.ReadAllText("data/product.json")),
                MsService.FromJson(File.ReadAllText("data/ms.service.json")),
                MsProd.FromJson(File.ReadAllText("data/ms.prod.json"))
            );

            Taxonomy unmapped = new(everything);

            Console.WriteLine(unmapped.Products.Count());
            Console.WriteLine(unmapped.MsServices.Count());
            Console.WriteLine(unmapped.MsProds.Count());

            var unified = MapProducts(everything, unmapped);

            Console.WriteLine(unmapped.Products.Count());
            Console.WriteLine(unmapped.MsServices.Count());
            Console.WriteLine(unmapped.MsProds.Count());

            BuildReport(everything, unified, unmapped);
        }

        static IEnumerable<Unified> MapProducts(Taxonomy everything, Taxonomy unmapped)
        {
            IEnumerable<Product> products = everything.Products;
            IEnumerable<MsService> msServices = everything.MsServices;
            IEnumerable<MsProd> msProds = everything.MsProds;

            List<Unified> results = new();

            foreach (var product in products)
            {
                Unified result = new() { Product = product };

                List<MsService> services = new();
                services.AddRange(MapProductToMsServices(product, msServices));
                services.AddRange(MapProductToMsSubServices(product, msServices));
                result.MsServices = services.GroupBy(s => s.Id).Select(g => g.First());

                result.MsProds = MapProductToMsProds(product, msProds).GroupBy(p => p.Id).Select(g => g.First());

                foreach (var service in result.MsServices)
                {
                    int r = unmapped.MsServices.RemoveAll(s => s.Id == service.Id);
                    //System.Diagnostics.Debug.Assert(r == 1);
                }
                foreach (var prod in result.MsProds)
                {
                    int r = unmapped.MsProds.RemoveAll(p => p.Id == prod.Id);
                    //System.Diagnostics.Debug.Assert(r == 1);
                }
                if (result.MsServices.Count() > 0 || result.MsProds.Count() > 0)
                {
                    int r = unmapped.Products.RemoveAll(p => p.Id == product.Id);
                    System.Diagnostics.Debug.Assert(r == 1);
                }

                results.Add(result);
                //Console.WriteLine(product.Label);
            }

            return results;
        }

        static IEnumerable<MsProd> MapProductToMsProds(Product product, IEnumerable<MsProd> msProds)
        {
            List<MsProd> results = new();

            // Try mapping product slug to ms.product.
            string productSlug = product.Slug;
            results.AddRange(msProds.Where(prod => string.Equals(prod.MsProduct, productSlug, StringComparison.CurrentCultureIgnoreCase)));
            var productSlugNoAzure = productSlug.Replace("azure-", string.Empty);
            if (productSlugNoAzure != productSlug)
            {
                results.AddRange(msProds.Where(prod => string.Equals(prod.MsProduct, productSlugNoAzure, StringComparison.CurrentCultureIgnoreCase)));
            }

            // Try matching labels.
            string label = product.Label;
            results.AddRange(msProds.Where(prod => string.Equals(prod.Product, label, StringComparison.CurrentCultureIgnoreCase)));
            var labelNoAzure = label.Replace("Azure ", string.Empty);
            if (labelNoAzure != label)
            {
                results.AddRange(msProds.Where(prod => string.Equals(prod.Product, labelNoAzure, StringComparison.CurrentCultureIgnoreCase)));
            }

            return results;
        }

        static IEnumerable<MsService> MapProductToMsServices(Product product, IEnumerable<MsService> msServices)
        {
            List<MsService> results = new();

            // Try mapping product slug to ms.product.
            string productSlug = product.Slug;
            results.AddRange(msServices.Where(service => string.Equals(service.MsService_, productSlug, StringComparison.CurrentCultureIgnoreCase)));
            var productSlugNoAzure = productSlug.Replace("azure-", string.Empty);
            if (productSlugNoAzure != productSlug)
            {
                results.AddRange(msServices.Where(service => string.Equals(service.MsService_, productSlugNoAzure, StringComparison.CurrentCultureIgnoreCase)));
            }

            // Try matching labels.
            string label = product.Label;
            results.AddRange(msServices.Where(service => string.Equals(service.Service, label, StringComparison.CurrentCultureIgnoreCase)));
            var labelNoAzure = label.Replace("Azure ", string.Empty);
            if (labelNoAzure != label)
            {
                results.AddRange(msServices.Where(service => string.Equals(service.Service, labelNoAzure, StringComparison.CurrentCultureIgnoreCase)));
            }

            return results;
        }

        static IEnumerable<MsService> MapProductToMsSubServices(Product product, IEnumerable<MsService> msServices)
        {
            List<MsService> results = new();

            // Try mapping product slug to ms.product.
            string productSlug = product.Slug;
            results.AddRange(msServices.Where(service => string.Equals(service.MsSubService, productSlug, StringComparison.CurrentCultureIgnoreCase)));
            var productSlugNoAzure = productSlug.Replace("azure-", string.Empty);
            if (productSlugNoAzure != productSlug)
            {
                results.AddRange(msServices.Where(service => string.Equals(service.MsSubService, productSlugNoAzure, StringComparison.CurrentCultureIgnoreCase)));
            }

            // Try matching labels.
            string label = product.Label;
            results.AddRange(msServices.Where(service => string.Equals(service.MsSubService, label, StringComparison.CurrentCultureIgnoreCase)));
            var labelNoAzure = label.Replace("Azure ", string.Empty);
            if (labelNoAzure != label)
            {
                results.AddRange(msServices.Where(service => string.Equals(service.MsSubService, labelNoAzure, StringComparison.CurrentCultureIgnoreCase)));
            }

            return results;
        }

        static void PullLatestTaxonomy()
        {
            Uri[] uris = new[] {
                new Uri("https://taxonomy.docs.microsoft.com/taxonomies/product"),
                new Uri("https://taxonomy.docs.microsoft.com/taxonomies/ms.prod"),
                new Uri("https://taxonomy.docs.microsoft.com/taxonomies/ms.service"),
                new Uri("https://taxonomy.docs.microsoft.com/taxonomies/ms.topic"),
            };

            HttpClient client = new();
            foreach (var uri in uris)
            {
                var request = new HttpRequestMessage(HttpMethod.Get, uri);
                request.Headers.Add("User-Agent", "Mozilla/5.0");
                var response = client.Send(request);
                Stream receiveStream = response.Content.ReadAsStream();
                StreamReader readStream = new StreamReader(receiveStream, System.Text.Encoding.UTF8);
                var content = readStream.ReadToEnd();
                File.WriteAllText(Path.Join("data", uri.Segments.Last() + ".json"), content);
            }
        }

        static void BuildReport(Taxonomy everything, IEnumerable<Unified> unified, Taxonomy unmapped)
        {
            using (var workbook = new XLWorkbook())
            {
                BuildProductSheet(workbook, everything.Products);
                BuildMsServiceSheet(workbook, everything.MsServices);
                BuildMsProdSheet(workbook, everything.MsProds);
                BuildProductToMsProdSheet(workbook, unified);
                BuildProductToMsServiceSheet(workbook, unified);
                BuildUnifiedSimpleSheet(workbook, unified);
                BuildUnifiedUnmappedSheet(workbook, unmapped);

                workbook.SaveAs($"UnifiedTaxonomy-{String.Format("{0:yyyy-MM-dd}", DateTime.Now)}.xlsx");
            }
        }

        static void BuildProductSheet(XLWorkbook workbook, List<Product> products)
        {
            var worksheet = workbook.Worksheets.Add("All Product");

            string[] columns = { "Slug", "Label", "Level", "Parent" };
            AddHeaderRow(worksheet, columns);

            products.Sort(CompareProducts);

            int index = 2;
            foreach (var product in products)
            {
                worksheet.Cell(index, 1).Value = product.Slug;
                worksheet.Cell(index, 2).Value = product.Label;
                worksheet.Cell(index, 3).Value = product.Level.ToString();
                var parentSlug = product.ParentSlug;
                worksheet.Cell(index, 4).Value = parentSlug;
                if (! string.IsNullOrEmpty(parentSlug))
                {
                    int row = products.FindIndex(p => p.Slug == parentSlug) + 2;
                    worksheet.Cell(index, 4).Hyperlink = new XLHyperlink($"A{row}");
                }
                index++;
            }

            FormatWorksheet(worksheet);
            // Protect worksheet.
            worksheet.Protect().AllowElement(XLSheetProtectionElements.AutoFilter);
            worksheet.RangeUsed().SetAutoFilter(true);
        }

        static void BuildMsServiceSheet(XLWorkbook workbook, List<MsService> msServices)
        {
            var worksheet = workbook.Worksheets.Add("All ms.service");

            string[] columns = { "ms.service", "Service", "ms.subservice", "SubService", "Manager", "Service Area", "Pillar", "Active?" };
            AddHeaderRow(worksheet, columns);

            msServices.Sort(CompareMsServices);

            int index = 2;
            foreach (var service in msServices)
            {
                worksheet.Cell(index, 1).Value = service.MsService_;
                worksheet.Cell(index, 2).Value = service.Service;
                worksheet.Cell(index, 3).Value = service.MsSubService;
                worksheet.Cell(index, 4).Value = service.SubService;
                worksheet.Cell(index, 5).Value = service.Manager;
                worksheet.Cell(index, 6).Value = service.ServiceArea;
                worksheet.Cell(index, 7).Value = service.Pillar;
                worksheet.Cell(index, 8).Value = service.Active;
                index++;
            }

            FormatWorksheet(worksheet);
        }

        static void BuildMsProdSheet(XLWorkbook workbook, List<MsProd> msProds)
        {
            var worksheet = workbook.Worksheets.Add("All ms.prod");

            string[] columns = { "ms.product", "Product", "Manager", "ms.technology", "Technology", "Pillar", "Active?" };
            AddHeaderRow(worksheet, columns);

            msProds.Sort(CompareMsProds);

            int index = 2;
            foreach (var prod in msProds)
            {
                worksheet.Cell(index, 1).Value = prod.MsProduct;
                worksheet.Cell(index, 2).Value = prod.Product;
                worksheet.Cell(index, 3).Value = prod.Manager;
                worksheet.Cell(index, 4).Value = prod.MsTechnology;
                worksheet.Cell(index, 5).Value = prod.Technology;
                worksheet.Cell(index, 6).Value = prod.Pillar;
                worksheet.Cell(index, 7).Value = prod.Active;
                index++;
            }

            FormatWorksheet(worksheet);
        }

        static void BuildProductToMsProdSheet(XLWorkbook workbook, IEnumerable<Unified> unified_)
        {
            var worksheet = workbook.Worksheets.Add("Product to ms.prod");

            string[] columns = { "Product slug", "Product label", "Level", "ms.product", "ms.product label", "Manager", "ms.technology", "Technology", "Pillar", "Active?" };
            AddHeaderRow(worksheet, columns);

            var unified = unified_.ToList();
            unified.Sort(CompareUnified);

            int index = 2;
            foreach (var item in unified)
            {
                var product = item.Product;
                var msProds = item.MsProds.ToList();
                msProds.Sort(CompareMsProds);
                foreach (var prod in msProds)
                {
                    worksheet.Cell(index, 1).Value = product.Slug;
                    worksheet.Cell(index, 2).Value = product.Label;
                    worksheet.Cell(index, 3).Value = product.Level.ToString();
                    worksheet.Cell(index, 4).Value = prod.MsProduct;
                    worksheet.Cell(index, 5).Value = prod.Product;
                    worksheet.Cell(index, 6).Value = prod.Manager;
                    worksheet.Cell(index, 7).Value = prod.MsTechnology;
                    worksheet.Cell(index, 8).Value = prod.Technology;
                    worksheet.Cell(index, 9).Value = prod.Pillar;
                    worksheet.Cell(index, 10).Value = prod.Active;
                    index++;
                }
            }

            FormatWorksheet(worksheet);
        }

        static void BuildProductToMsServiceSheet(XLWorkbook workbook, IEnumerable<Unified> unified_)
        {
            var worksheet = workbook.Worksheets.Add("Product to ms.service");

            string[] columns = { "Product slug", "Product label", "Level", "ms.service", "Service", "ms.subservice", "SubService", "Manager", "Service Area", "Pillar", "Active?" };
            AddHeaderRow(worksheet, columns);

            var unified = unified_.ToList();
            unified.Sort(CompareUnified);

            int index = 2;
            foreach (var item in unified)
            {
                var product = item.Product;
                var msServices = item.MsServices.ToList();
                msServices.Sort(CompareMsServices);
                foreach (var service in msServices)
                {
                    worksheet.Cell(index, 1).Value = product.Slug;
                    worksheet.Cell(index, 2).Value = product.Label;
                    worksheet.Cell(index, 3).Value = product.Level.ToString();
                    worksheet.Cell(index, 4).Value = service.MsService_;
                    worksheet.Cell(index, 5).Value = service.Service;
                    worksheet.Cell(index, 6).Value = service.MsSubService;
                    worksheet.Cell(index, 7).Value = service.SubService;
                    worksheet.Cell(index, 8).Value = service.Manager;
                    worksheet.Cell(index, 9).Value = service.ServiceArea;
                    worksheet.Cell(index, 10).Value = service.Pillar;
                    worksheet.Cell(index, 11).Value = service.Active;
                    index++;
                }
            }

            FormatWorksheet(worksheet);
        }

        static void BuildUnifiedSimpleSheet(XLWorkbook workbook, IEnumerable<Unified> unified_)
        {
            var worksheet = workbook.Worksheets.Add("Unified (Simple)");

            string[] columns = { "Product slug", "Level", "ms.prod", "ms.service", "# ms.subservices", "Managers" };
            AddHeaderRow(worksheet, columns);

            var unified = unified_.ToList();
            unified.Sort(CompareUnified);

            int index = 2;
            foreach (var item in unified)
            {
                var msProd = item.MsProds
                    .Where(p => !string.IsNullOrEmpty(p.MsProduct))
                    .FirstOrDefault()?.MsProduct ?? string.Empty;

                var msService = item.MsServices
                    .Where(s => !string.IsNullOrEmpty(s.MsService_))
                    .FirstOrDefault()?.MsService_ ?? string.Empty;

                int msSubServices = item.MsServices.Where(s => !string.IsNullOrEmpty(s.MsSubService)).Count();

                List<string> managerList = new();
                managerList.AddRange(item.MsProds
                    .Where(p => !string.IsNullOrEmpty(p.Manager))
                    .Select(p => p.Manager));
                managerList.AddRange(item.MsServices
                    .Where(s => !string.IsNullOrEmpty(s.Manager))
                    .Select(s => s.Manager));
                managerList.Sort();
                var managers = string.Join(',', managerList.Distinct());

                worksheet.Cell(index, 1).Value = item.Product.Slug;
                worksheet.Cell(index, 2).Value = item.Product.Level.ToString();
                worksheet.Cell(index, 3).Value = msProd;
                worksheet.Cell(index, 4).Value = msService;
                worksheet.Cell(index, 5).Value = msSubServices.ToString();
                worksheet.Cell(index, 6).Value = managers;

                if (string.IsNullOrEmpty(msProd) && string.IsNullOrEmpty(msService) &&
                    msSubServices == 0)
                {
                    worksheet.Row(index).Cells(1, 6).Style.Fill.BackgroundColor = XLColor.PaleGoldenrod;
                }
                index++;
            }

            FormatWorksheet(worksheet);
        }

        struct UnmappedEntry
        {
            public IEnumerable<string> Items { get; init; }
            public int StartColumn { get; init; }

            public UnmappedEntry(string item, int column)
            {
                Items = new List<string>() { item };
                StartColumn = column;
            }

            public UnmappedEntry(string[] items, int startColumn)
            {
                Items = items.ToList();
                StartColumn = startColumn;
            }

            public static int Compare(UnmappedEntry e1, UnmappedEntry e2) => CompareSequences(e1.Items.GetEnumerator(), e2.Items.GetEnumerator());

            static int CompareSequences(IEnumerator<string> e1, IEnumerator<string> e2)
            {
                int c = 0;
                if (e1.MoveNext() && e2.MoveNext())
                {
                    c = e1.Current.CompareTo(e2.Current);
                    if (c == 0)
                    {
                        return CompareSequences(e1, e2);
                    }
                }
                return c;
            }
        }

        // Unmapped, Active
        static void BuildUnifiedUnmappedSheet(XLWorkbook workbook, Taxonomy unmapped)
        {
            var worksheet = workbook.Worksheets.Add("Unmapped (Active)");

            string[] columns = { "Product slug", "ms.prod", "ms.service", "ms.subservice" };
            AddHeaderRow(worksheet, columns);

            // Mark
            List<UnmappedEntry> entries = new();

            entries.AddRange(unmapped.Products
                .Select(p => p.Slug)
                .Distinct()
                .Select(slug => new UnmappedEntry(slug, 1)));

            entries.AddRange(unmapped.MsProds
                .Where(p => p.Active)
                .Select(p => p.MsProduct)
                .Distinct()
                .Select(msProduct => new UnmappedEntry(msProduct, 2)));

            entries.AddRange(unmapped.MsServices
                .Where(s => s.Active)
                .Select(s => new string[] { s.MsService_, s.MsSubService ?? string.Empty })
                .Distinct()
                .Select(items => new UnmappedEntry(items, 3)));

            entries.Sort(UnmappedEntry.Compare);

            // Sweep

            int index = 2;
            foreach (var entry in entries)
            {
                int column = entry.StartColumn;
                foreach (var item in entry.Items)
                {
                    worksheet.Cell(index, column).Value = item;
                    column++;
                }
                index++;
            }

            FormatWorksheet(worksheet);
        }

        static void AddHeaderRow(IXLWorksheet worksheet, params string[] values)
        {
            for (int i = 1; i <= values.Length; i++)
            {
                worksheet.Cell(1, i).Value = values[i-1];
            }

            var row = worksheet.Row(1);
            row.Style.Fill.BackgroundColor = XLColor.DarkCerulean;
            row.Style.Font.Bold = true;
            row.Style.Font.FontColor = XLColor.White;
            worksheet.SheetView.FreezeRows(1);
        }

        static void FormatWorksheet(IXLWorksheet worksheet)
        {
            foreach (var column in worksheet.ColumnsUsed())
            {
                column.AdjustToContents();
            }
            worksheet.RangeUsed().SetAutoFilter();
        }

        static int CompareUnified(Unified u1, Unified u2) => CompareProducts(u1.Product, u2.Product);

        static int CompareProducts(Product p1, Product p2)
        {
            int c = p1.Level.CompareTo(p2.Level);
            if (c == 0)
            {
                c = p1.Slug.CompareTo(p2.Slug);
            }
            return c;
        }

        static int CompareMsServices(MsService s1, MsService s2)
        {
            int c = s1.MsService_.CompareTo(s2.MsService_);
            if (c == 0)
            {
                if (s1.MsSubService == null)
                {
                    c = -1;
                }
                else if (s2.MsSubService == null)
                {
                    c = 1;
                }
                else
                {
                    c = s1.MsSubService.CompareTo(s2.MsSubService);
                }
            }
            return c;
        }

        static int CompareMsProds(MsProd p1, MsProd p2)
        {
            int c = p1.MsProduct.CompareTo(p2.MsProduct);
            if (c == 0)
            {
                c = p1.Product.CompareTo(p2.Product);
            }
            if (c == 0)
            {
                if (p1.MsTechnology == null)
                {
                    c = -1;
                }
                else if (p2.MsTechnology == null)
                {
                    c = 1;
                }
                else
                {
                    c = p1.MsTechnology.CompareTo(p2.MsTechnology);
                }
            }
            return c;
        }
    }
}
