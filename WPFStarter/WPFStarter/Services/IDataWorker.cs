using WPFStarter.Models;

namespace WPFStarter.Services
{
    public interface IDataWorker
    {
        Task ImportCsvAsync(ApplicationDbContext context, string filePath);
        Task ExportToExcelAsync(IEnumerable<DataRecord> records, string filePath);
        Task ExportToXmlAsync(IEnumerable<DataRecord> records, string filePath);
    }
}