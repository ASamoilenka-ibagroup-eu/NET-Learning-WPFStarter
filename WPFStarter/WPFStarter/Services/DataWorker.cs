using WPFStarter.Models;
using System.IO;
using CsvHelper;
using OfficeOpenXml;
using System.Xml.Linq;

namespace WPFStarter.Services
{
    public class DataWorker : IDataWorker
    {
        public void ImportCsv(ApplicationDbContext context, string filePath)
        {
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, System.Globalization.CultureInfo.InvariantCulture))
            {
                var records = csv.GetRecords<DataRecord>().ToList();
                context.Records.AddRange(records);
                context.SaveChanges();
            }
        }

        public void ExportToExcel(ApplicationDbContext context, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Records");
                worksheet.Cells["A1"].LoadFromCollection(context.Records.ToList(), true);
                package.SaveAs(new FileInfo(filePath));
            }
        }

        public void ExportToXml(ApplicationDbContext context, string filePath)
        {
            var records = context.Records.ToList();
            var xDocument = new XDocument(new XElement("TestProgram",
                records.Select(record => new XElement("Record",
                    new XAttribute("id", record.Id),
                    new XElement("Date", record.Date),
                    new XElement("FirstName", record.FirstName),
                    new XElement("LastName", record.LastName),
                    new XElement("SurName", record.SurName),
                    new XElement("City", record.City),
                    new XElement("Country", record.Country)
                ))));

            xDocument.Save(filePath);
        }
    }
}