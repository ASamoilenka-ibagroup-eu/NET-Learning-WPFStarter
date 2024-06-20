using WPFStarter.Models;

namespace WPFStarter.Services
{
    public interface IDataWorker
    {
        void ImportCsv(ApplicationDbContext context, string filePath);
        void ExportToExcel(ApplicationDbContext context, string filePath);
        void ExportToXml(ApplicationDbContext context, string filePath);
    }
}