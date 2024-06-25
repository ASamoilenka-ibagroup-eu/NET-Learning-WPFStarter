using System.IO;
using OfficeOpenXml;
using System.Xml.Linq;
using WPFStarter.Models;

namespace WPFStarter.Services
{
    public class DataWorker : IDataWorker
    {
        public DataWorker()
        {
            // Set the license context for EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async Task ImportCsvAsync(ApplicationDbContext context, string filePath)
        {
            try
            {
                var lines = await File.ReadAllLinesAsync(filePath);
                var records = new List<DataRecord>();

                bool isFirstLine = true;
                foreach (var line in lines)
                {
                    // Skip the first line with headers
                    if (isFirstLine)
                    {
                        isFirstLine = false;
                        continue;
                    }

                    var values = line.Split(';');
                    var record = new DataRecord
                    {
                        Date = DateTime.Parse(values[0]),
                        FirstName = values[1],
                        LastName = values[2],
                        SurName = values[3],
                        City = values[4],
                        Country = values[5]
                    };
                    records.Add(record);
                }

                await context.Records.AddRangeAsync(records);
                await context.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error importing CSV file at {filePath}: {ex.Message}", ex);
            }
        }

        public async Task ExportToExcelAsync(IEnumerable<DataRecord> records, string filePath)
        {
            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Data");
                    worksheet.Cells["A1"].LoadFromCollection(records, true);
                    await package.SaveAsAsync(new FileInfo(filePath));
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error exporting to Excel file at {filePath}: {ex.Message}", ex);
            }
        }

        public async Task ExportToXmlAsync(IEnumerable<DataRecord> records, string filePath)
        {
            try
            {
                var xDocument = new XDocument(
                    new XElement("TestProgram",
                        records.Select(r => new XElement("Record",
                            new XAttribute("id", r.Id),
                            new XElement("Date", r.Date),
                            new XElement("FirstName", r.FirstName),
                            new XElement("LastName", r.LastName),
                            new XElement("SurName", r.SurName),
                            new XElement("City", r.City),
                            new XElement("Country", r.Country)))));

                await Task.Run(() => xDocument.Save(filePath));
            }
            catch (Exception ex)
            {
                throw new Exception($"Error exporting to XML file at {filePath}: {ex.Message}", ex);
            }
        }
    }
}