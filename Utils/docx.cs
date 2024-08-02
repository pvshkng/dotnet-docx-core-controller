using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;

namespace Generation.Utils
{
    public class DocxGenerator
    {
        public void GenerateAndSaveDocxFile(string jsonString)
        {
            var jsonData = JArray.Parse(jsonString);

            string resultFolderPath = Path.Combine(Directory.GetCurrentDirectory(), "result");
            if (!Directory.Exists(resultFolderPath))
            {
                Directory.CreateDirectory(resultFolderPath);
            }

            string timestamp = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
            string fileName = $"file-{timestamp}.docx";
            string filePath = Path.Combine(resultFolderPath, fileName);

            using (MemoryStream stream = new MemoryStream())
            {
                using (
                    WordprocessingDocument wordDocument = WordprocessingDocument.Create(
                        stream,
                        DocumentFormat.OpenXml.WordprocessingDocumentType.Document
                    )
                )
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body docBody = new Body();
                    mainPart.Document.Append(docBody);

                    foreach (var item in jsonData)
                    {
                        try
                        {
                            if (item is JObject data)
                            {
                                Table table = CreateTable(data);
                                docBody.Append(table);
                                docBody.Append(new Paragraph(new Run(new Text("\n"))));
                            }
                            else
                            {
                                Console.WriteLine("Error: Data item is not a JObject.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error generating table: {ex.Message}");
                        }
                    }

                    mainPart.Document.Save();
                }

                File.WriteAllBytes(filePath, stream.ToArray());
            }
        }

        private Table CreateTable(JObject data)
        {
            Table table = new Table();
            TableProperties tblProps = new TableProperties(
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }
            );
            table.Append(tblProps);

            var fiscalYears = data["fiscal_years"]?.ToObject<string[]>() ?? new string[0];
            var metricsKeys = new[]
            {
                "Revenue",
                "% Revenue Growth",
                "Cost of Goods Sold/Service %",
                "Gross Profit Margin %",
                "SG&A%",
                "Depreciation + Amortization",
                "Operating Margin %",
                "EBITDA (excl. Other Income)",
                "EBITDA % Change Y/Y",
                "EBITDA Margin %",
                "Other Income",
                "Interest Expense",
                "Net Profit",
                "Net % Change Y/Y",
                "Net Profit Margin %"
            };

            TableRow headerRow = new TableRow();
            TableCell headerCell = new TableCell();
            Paragraph headerParagraph = new Paragraph(
                new Run(new Text(data["company_name"].ToString()))
            );
            headerCell.Append(headerParagraph);
            headerCell.TableCellProperties = new TableCellProperties
            {
                TableCellWidth = new TableCellWidth
                {
                    Width = "0",
                    Type = TableWidthUnitValues.Auto
                }
            };
            headerRow.Append(headerCell);
            table.Append(headerRow);

            TableRow metricsHeaderRow = new TableRow();
            metricsHeaderRow.Append(
                new TableCell(new Paragraph(new Run(new Text("Financial Metrics"))))
            );
            foreach (var year in fiscalYears)
            {
                metricsHeaderRow.Append(
                    new TableCell(new Paragraph(new Run(new Text(year.ToString()))))
                );
            }
            table.Append(metricsHeaderRow);

            foreach (var metric in metricsKeys)
            {
                TableRow row = new TableRow();
                row.Append(new TableCell(new Paragraph(new Run(new Text(metric)))));

                var values = data[metric]?.ToObject<string[]>() ?? new string[fiscalYears.Length];
                for (int i = 0; i < fiscalYears.Length; i++)
                {
                    string value = i < values.Length ? values[i] : string.Empty;
                    row.Append(new TableCell(new Paragraph(new Run(new Text(value)))));
                }

                table.Append(row);
            }

            var settings = new XmlWriterSettings
            {
                Indent = true,
                OmitXmlDeclaration = true,
                NewLineOnAttributes = true
            };

            using (var stringWriter = new StringWriter())
            using (var xmlWriter = XmlWriter.Create(stringWriter, settings))
            {
                table.WriteTo(xmlWriter);
                xmlWriter.Flush();
                string xmlString = stringWriter.ToString();
                // Console.WriteLine("Table XML:");
                // Console.WriteLine(xmlString);
            }

            return table;
        }

        public List<TableXml> GenerateTableXmls(string jsonString)
        {
            var jsonData = JArray.Parse(jsonString);
            var tableXmls = new List<TableXml>();

            foreach (var item in jsonData)
            {
                if (item is JObject data)
                {
                    Table table = CreateTable(data);
                    string xml = TableToXmlString(table);
                    tableXmls.Add(
                        new TableXml { company_name = data["company_name"].ToString(), xml = xml }
                    );
                }
            }

            return tableXmls;
        }

        private string TableToXmlString(Table table)
        {
            var settings = new XmlWriterSettings
            {
                OmitXmlDeclaration = true,
                NewLineHandling = NewLineHandling.None,
                Indent = false
            };

            using (var stringWriter = new StringWriter())
            using (var xmlWriter = XmlWriter.Create(stringWriter, settings))
            {
                table.WriteTo(xmlWriter);
                xmlWriter.Flush();
                string xmlString = stringWriter.ToString();

                // Remove whitespace
                xmlString = System.Text.RegularExpressions.Regex.Replace(xmlString, @">\s+<", "><");

                // Remove xmlns
                xmlString = System.Text.RegularExpressions.Regex.Replace(
                    xmlString,
                    @"\s+xmlns(?::\w+)?=""[^""]*""",
                    ""
                );

                return xmlString;
            }
        }

        public class TableXml
        {
            public string company_name { get; set; }
            public string xml { get; set; }
        }
    }
}
