using System.Globalization;
using System.IO;
using System.Xml.Linq;
using CsvHelper;
using OfficeOpenXml;
using WPFStarter.Models;
using WPFStarter.Models.Data;

namespace WPFStarter.Services
{
    public interface IDataWorker
    {
        void ImportCsv(ApplicationDbContext context, string filePath);
        void ExportToExcel(ApplicationDbContext context, string filePath);
        void ExportToXml(ApplicationDbContext context, string filePath);
    }

    public class DataWorker : IDataWorker
    {
        public void ImportCsv(ApplicationDbContext context, string filePath)
        {
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var records = csv.GetRecords<DataRecord>().ToList();
                context.Records.AddRange(records);
                context.SaveChanges();
            }
        }

        public void ExportToExcel(ApplicationDbContext context, string filePath)
        {
            var records = context.Records.ToList();
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Records");
                worksheet.Cells[1, 1].Value = "Id";
                worksheet.Cells[1, 2].Value = "Date";
                worksheet.Cells[1, 3].Value = "FirstName";
                worksheet.Cells[1, 4].Value = "LastName";
                worksheet.Cells[1, 5].Value = "SurName";
                worksheet.Cells[1, 6].Value = "City";
                worksheet.Cells[1, 7].Value = "Country";

                for (int i = 0; i < records.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = records[i].Id;
                    worksheet.Cells[i + 2, 2].Value = records[i].Date.ToString("yyyy-MM-dd");
                    worksheet.Cells[i + 2, 3].Value = records[i].FirstName;
                    worksheet.Cells[i + 2, 4].Value = records[i].LastName;
                    worksheet.Cells[i + 2, 5].Value = records[i].SurName;
                    worksheet.Cells[i + 2, 6].Value = records[i].City;
                    worksheet.Cells[i + 2, 7].Value = records[i].Country;
                }

                package.SaveAs(new FileInfo(filePath));
            }
        }

        public void ExportToXml(ApplicationDbContext context, string filePath)
        {
            var records = context.Records.ToList();
            var xml = new XElement("TestProgram",
                records.Select(r => new XElement("Record",
                    new XAttribute("id", r.Id),
                    new XElement("Date", r.Date.ToString("yyyy-MM-dd")),
                    new XElement("FirstName", r.FirstName),
                    new XElement("LastName", r.LastName),
                    new XElement("SurName", r.SurName),
                    new XElement("City", r.City),
                    new XElement("Country", r.Country)
                ))
            );

            xml.Save(filePath);
        }
    }
}