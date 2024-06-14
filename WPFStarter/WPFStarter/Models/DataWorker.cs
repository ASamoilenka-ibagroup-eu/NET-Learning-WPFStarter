using System.Globalization;
using System.IO;
using CsvHelper;
using CsvHelper.Configuration;
using WPFStarter.Models.Data;
using OfficeOpenXml;
using System.Xml.Linq;

namespace WPFStarter.Models
{
    public static class DataWorker
    {
        public static void ImportCsv(string filePath, ApplicationDbContext context)
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ";",
                HasHeaderRecord = true,
            };

            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, config))
            {
                var records = csv.GetRecords<DataRecord>().ToList();
                context.Records.AddRange(records);
                context.SaveChanges();
            }
        }

        public static void ExportToExcel(List<DataRecord> records, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Data");

                worksheet.Cells[1, 1].Value = "Date";
                worksheet.Cells[1, 2].Value = "FirstName";
                worksheet.Cells[1, 3].Value = "LastName";
                worksheet.Cells[1, 4].Value = "SurName";
                worksheet.Cells[1, 5].Value = "City";
                worksheet.Cells[1, 6].Value = "Country";

                for (int i = 0; i < records.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = records[i].Date;
                    worksheet.Cells[i + 2, 2].Value = records[i].FirstName;
                    worksheet.Cells[i + 2, 3].Value = records[i].LastName;
                    worksheet.Cells[i + 2, 4].Value = records[i].SurName;
                    worksheet.Cells[i + 2, 5].Value = records[i].City;
                    worksheet.Cells[i + 2, 6].Value = records[i].Country;
                }

                package.SaveAs(new FileInfo(filePath));
            }
        }

        public static void ExportToXml(List<DataRecord> records, string filePath)
        {
            var xDoc = new XDocument(
                new XElement("TestProgram",
                    records.Select(record =>
                        new XElement("Record",
                            new XAttribute("id", record.Id),
                            new XElement("Date", record.Date),
                            new XElement("FirstName", record.FirstName),
                            new XElement("LastName", record.LastName),
                            new XElement("SurName", record.SurName),
                            new XElement("City", record.City),
                            new XElement("Country", record.Country)))));

            xDoc.Save(filePath);
        }

        public static List<DataRecord> GetRecords(ApplicationDbContext context)
        {
            return context.Records.ToList();
        }

        public static List<DataRecord> FilterRecords(ApplicationDbContext context, string city, string lastName)
        {
            return context.Records
                .Where(r => r.City == city && r.LastName == lastName)
                .ToList();
        }
    }
}