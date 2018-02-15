using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using CSV.Models;

namespace ImportProducts.Services
{
    public class ImportCsvJob
    {
        private readonly LogWriter _logger;
        private readonly ExcelMapper _mapper;

        public ImportCsvJob()
        {
            this._logger = new LogWriter();
            this._mapper = new ExcelMapper();
        }

        public StringBuilder DoJob(string refff, IEnumerable<string> t2TreFs)
        {
            var csvLines = new StringBuilder();
            _logger.LogWrite("Generating stock.csv: This will take a few minutes, please wait....");
            var descriptions = _mapper.MapToDescriptions();

            using (var connectionHandler = new OleDbConnection(System.Configuration.ConfigurationManager.AppSettings["AccessConnectionString"]))
            {
                connectionHandler.Open();

                var reff = refff.Substring(0, 6);

                var data = new DataSet();
                var myAccessCommand = new OleDbCommand(SqlQuery.ImportProductsQuery, connectionHandler);
                myAccessCommand.Parameters.AddWithValue("?", reff);

                var myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                myDataAdapter.Fill(data);

                var actualStock = "0";
                var inStockFlag = false;
                var groupSkus = "";

                foreach (DataRow dr in data.Tables[0].Rows)
                {
                    _logger.LogWrite("Working....");
                    var isStock = 0;
                    var simpleSkusList = new List<string>();
                    for (var i = 1; i < 14; i++)
                    {
                        if (!string.IsNullOrEmpty(dr["QTY" + i].ToString()))
                        {
                            if (dr["QTY" + i].ToString() != "")
                            {
                                if (Convert.ToInt32(dr["QTY" + i]) > 0)
                                {
                                    if (String.IsNullOrEmpty(dr["LY" + i].ToString()))
                                    {
                                        actualStock = dr["QTY" + i].ToString();
                                    }
                                    else
                                    {
                                        actualStock =
                                            (Convert.ToInt32(dr["QTY" + i]) - Convert.ToInt32(dr["LY" + i]))
                                            .ToString();
                                    }

                                    isStock = 1;
                                    inStockFlag = true;
                                }
                                else
                                {
                                    isStock = 0;
                                }
                                var append = (1000 + i).ToString();
                                groupSkus = dr["NewStyle"].ToString();
                                var groupSkus2 = dr["NewStyle"] + append.Substring(1, 3);
                                var shortDescription = BuildShortDescription(descriptions.FirstOrDefault(x => x.T2TRef == reff));
                                var descripto = descriptions.Where(x => x.T2TRef == reff)
                                    .Select(y => y.Descriptio).FirstOrDefault();

                                var size = "";
                                size = i < 10 ? dr["S0" + i].ToString() : dr["S" + i].ToString();

                                if (size.Contains("½"))
                                    size = size.Replace("½", ".5");

                                if (!string.IsNullOrEmpty(size))
                                {
                                    simpleSkusList.Add(groupSkus2);
                                    var newLine = BuildChildImportProduct(groupSkus2, dr, descriptions, reff,
                                        shortDescription, actualStock, descripto, size, isStock, refff, t2TreFs);
                                    csvLines.AppendLine(newLine);
                                }

                            }
                            actualStock = "0";
                        }
                    }

                    isStock = inStockFlag ? 1 : 0;
                    if (!string.IsNullOrEmpty(dr["NewStyle"].ToString()))
                    {
                        var newLine = ParentImportProduct(groupSkus, descriptions, reff, dr, simpleSkusList,
                            isStock, refff, t2TreFs);
                        csvLines.AppendLine(newLine);
                    }
                    inStockFlag = false;
                    if (data.Tables[0].Rows.Count > 1)
                    {
                        break;
                    }

                }
            }
            
            _logger.LogWrite("Finished for: " + refff);
            return csvLines;
        }

        private static string ParentImportProduct(string groupSkus, List<Descriptions> descriptions, string reff, DataRow dr, List<string> simpleSkusList,
            int isStock, string reffColour, IEnumerable<string> t2TreFs)
        {
            var store = "\"admin\"";
            var websites = Websites()?.TrimEnd();
            var attribut_set = "\"Default\"";
            var type = "\"configurable\"";
            var sku = "\"" + groupSkus?.TrimEnd() + "\"";
            var hasOption = "\"1\"";
            var name = "\"" + descriptions.Where(x => x.T2TRef == reff).Select(y => y.Descriptio).FirstOrDefault() + " in " +
                       dr["MasterColour"] + "\"";
            var pageLayout = "\"No layout updates.\"";
            var optionsContainer = "\"Product Info Column\"";
            var price = "\"" + dr["BASESELL"].ToString().TrimEnd() + "\"";
            var weight = "\"0.01\"";
            var status = "\"Enabled\"";
            var visibility = Visibility()?.TrimEnd();
            var shortDescription = "\"" + BuildShortDescription(descriptions.FirstOrDefault(x => x.T2TRef == reff)) + "\"";
            var gty = "\"0\"";
            var productName = "\"" + descriptions.Where(x => x.T2TRef == reff).Select(y => y.Descriptio).FirstOrDefault() + "\"";
            var color = "\"" + dr["MasterColour"].ToString().TrimEnd() + "\"";
            var sizeRange = "\"\"";
            var vat = dr["VAT"].ToString() == "A" ? "TAX" : "None";
            var taxClass = "\"" + vat + "\"";
            var configurableAttribute = "\"size\"";
            var simpleSku = BuildSimpleSku(simpleSkusList, reff);
            var manufactor = "\"" + dr["MasterSupplier"] + "\"";
            var isInStock = "\"" + isStock + "\"";
            var category = "\"" + Category(dr) + "\"";
            var season = "\"\"";
            var stockType = "\"" + dr["MasterStocktype"] + "\"";
            var image = "\"+/" + reffColour + ".jpg\"";
            var smallImage = "\"/" + reffColour + ".jpg\"";
            var thumbnail = "\"/" + reffColour + ".jpg\"";
            var gallery = "\"" + BuildGalleryImages(t2TreFs, reff) + "\"";
            var condition = "\"new\"";
            var ean = "\"\"";
            var description = "\"" + descriptions.Where(x => x.T2TRef == reff).Select(y => y.Description).FirstOrDefault()?.TrimEnd() +
                              "\"";
            var model = "\"" + dr["SHORT"] + "\"";

            var newLine = $"{store}," +
                          $"{websites},{attribut_set},{type},{sku},{hasOption},{name.TrimEnd()},{pageLayout},{optionsContainer},{price},{weight},{status}," +
                          $"{visibility}," +
                          $"{shortDescription},{gty},{productName},{color}," +
                          $"{sizeRange},{taxClass},{configurableAttribute},{simpleSku},{manufactor},{isInStock}," +
                          $"{category},{season},{stockType},{image},{smallImage},{thumbnail},{gallery},{condition},{ean}," +
                          $"{description},{model}";
            return newLine;
        }

        private static string Category(DataRow dr)
        {
            var category = dr["MasterStocktype"] + "/Shop By Department/" + dr["MasterDept"] + ";;";

            if (dr["MasterSubDept"] != "ANY" || dr["MasterSubDept"] != "")
            {
                category = category + dr["MasterStocktype"] + "/Shop By Department/" +
                           dr["MasterDept"] + "/" + dr["MasterSubDept"] + "::1::1::0;;";
            }

            category = category + dr["MasterStocktype"] + "/Shop By Brand/" +
                       dr["MasterSupplier"] + ";;";
            category = category + dr["MasterStocktype"] + "/Shop By Brand/" +
                       dr["MasterSupplier"] + "/" + dr["MasterDept"] + "::1::1::0;;";

            if (dr["MasterSubDept"] != "ANY" || dr["MasterSubDept"] != "")
            {
                category = category + dr["MasterStocktype"] + "/Shop By Brand/" +
                           dr["MasterSupplier"] + "/" + dr["MasterDept"] +
                           "/" + dr["MasterSubDept"] + "::1::1::0;;";
            }
            category = category + "Brands/" + dr["MasterSupplier"];
            return category;
        }

        private static string BuildChildImportProduct(string groupSkus2, DataRow dr, List<Descriptions> descriptions, string reff,
            string short_description, string actualStock, string descripto, string size, int isStock, string reffColour,
            IEnumerable<string> t2TreFs)
        {
            const string store = "\"admin\"";
            var websites = Websites()?.TrimEnd();
            const string attribut_set = "\"Default\"";
            const string type = "\"simple\"";
            var sku = "\"" + groupSkus2?.TrimEnd() + "\"";
            const string hasOption = "\"1\"";
            var name = "\"" + dr["MasterSupplier"] + " " + descriptions.Where(x => x.T2TRef == reff).Select(y => y.Descriptio).FirstOrDefault() + " in " + dr["MasterColour"] + "\"";
            const string pageLayout = "\"No layout updates.\"";
            const string optionsContainer = "\"Product Info Column\"";
            var price = "\"" + dr["BASESELL"].ToString().TrimEnd() + "\"";
            const string weight = "\"0.01\"";
            const string status = "\"Enabled\"";
            const string visibility = "\"Not Visible Individually\"";
            var shortDescription = "\"" + short_description?.TrimEnd() + "\"";
            var gty = "\"" + actualStock + "\"";
            var productName = "\"" + descripto?.TrimEnd() + "\"";
            var color = "\"" + dr["MasterColour"].ToString().TrimEnd() + "\"";
            var sizeRange = "\"" + dr["SIZERANGE"] + size + "\"";
            var vat = dr["VAT"].ToString() == "A" ? "TAX" : "None";
            var taxClass = "\"" + vat + "\"";
            const string configurableAttribute = "\"\"";
            const string simpleSku = "\"\"";
            var manufactor = "\"" + dr["MasterSupplier"] + "\"";
            var isInStock = "\"" + isStock + "\"";
            const string category = "\"\"";
            const string season = "\"\"";
            var stockType = "\"" + dr["MasterStocktype"] + "\"";
            var image = "\"+/" + reffColour + ".jpg\"";
            var smallImage = "\"/" + reffColour + ".jpg\"";
            var thumbnail = "\"/" + reffColour + ".jpg\"";
            var gallery = "\"" + BuildGalleryImages(t2TreFs, reff) + "\"";
            const string condition = "\"new\"";
            const string ean = "\"\"";
            var description = "\"" + descriptions.Where(x => x.T2TRef == reff).Select(y => y.Description).FirstOrDefault()?.TrimEnd() + "\"";
            var model = "\"" + dr["SHORT"] + "\"";

            var newLine = $"{store}," +
                          $"{websites},{attribut_set},{type},{sku},{hasOption},{name.TrimEnd()},{pageLayout},{optionsContainer},{price},{weight},{status},{visibility}," +
                          $"{shortDescription},{gty},{productName},{color}," +
                          $"{sizeRange},{taxClass},{configurableAttribute},{simpleSku},{manufactor},{isInStock}," +
                          $"{category},{season},{stockType},{image},{smallImage},{thumbnail},{gallery},{condition},{ean}," +
                          $"{description},{model}";
            return newLine;
        }


        private static string BuildGalleryImages(IEnumerable<string> t2TreFs, string reff)
        {
            var images = t2TreFs.Where(t2TRef => t2TRef.Contains(reff)).ToList();
            return images.Aggregate("", (current, image) => current + ("/" + image + ".jpg;"));
        }

        private static string BuildSimpleSku(IEnumerable<string> t2TreFs, string reff)
        {
            var output = t2TreFs.Where(t2TReff => t2TReff.Contains(reff)).Aggregate("\"", (current, t2TReff) => current + t2TReff + ",");
            return output.Remove(output.Length - 1) + "\"";
        }

        private static string Websites()
        {
            var output = "\"";
            var fields = new string[] { "admin", "base" };
            foreach (var field in fields)
            {
                output += field + ",";
            }
            return output.Remove(output.Length - 1) + "\"";
        }

        private static string Visibility()
        {
            var fields = new [] { "Catalog", "Search" };
            var output = fields.Aggregate("\"", (current, field) => current + (field + ","));
            return output.Remove(output.Length - 1) + "\"";
        }

        public void DoCleanup()
        {
            if (File.Exists(System.Configuration.ConfigurationManager.AppSettings["OutputPath"]))
            {
                File.Delete(System.Configuration.ConfigurationManager.AppSettings["OutputPath"]);
            }
        }

        private static string BuildShortDescription(Descriptions description)
        {
            if (description == null)
                return "<ul></ul>";

            if (string.IsNullOrEmpty(description.Bullet1) && string.IsNullOrEmpty(description.Bullet2) &&
                string.IsNullOrEmpty(description.Bullet3) && string.IsNullOrEmpty(description.Bullet4) &&
                string.IsNullOrEmpty(description.Bullet5) && string.IsNullOrEmpty(description.Bullet6) &&
                string.IsNullOrEmpty(description.Bullet7))
            {
                return "<ul></ul>";
            }
            var bullet1 = string.IsNullOrEmpty(description.Bullet1) ? "" : "<li>" + description.Bullet1 + "</li>";
            var bullet2 = string.IsNullOrEmpty(description.Bullet2) ? "" : "<li>" + description.Bullet2 + "</li>";
            var bullet3 = string.IsNullOrEmpty(description.Bullet3) ? "" : "<li>" + description.Bullet3 + "</li>";
            var bullet4 = string.IsNullOrEmpty(description.Bullet4) ? "" : "<li>" + description.Bullet4 + "</li>";
            var bullet5 = string.IsNullOrEmpty(description.Bullet5) ? "" : "<li>" + description.Bullet5 + "</li>";
            var bullet6 = string.IsNullOrEmpty(description.Bullet6) ? "" : "<li>" + description.Bullet6 + "</li>";
            var bullet7 = string.IsNullOrEmpty(description.Bullet7) ? "" : "<li>" + description.Bullet7 + "</li>";
            return "<ul>" + bullet1 + bullet2 + bullet3 + bullet4 + bullet5 + bullet6 + bullet7 + "</ul>";
        }
    }
}