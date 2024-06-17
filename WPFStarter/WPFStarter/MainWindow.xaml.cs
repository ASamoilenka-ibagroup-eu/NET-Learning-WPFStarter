using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Win32;
using OfficeOpenXml;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Xml.Linq;
using WPFStarter.Models;
using WPFStarter.Models.Data;

namespace WPFStarter
{
    public partial class MainWindow : Window
    {
        private readonly ApplicationDbContext _context;

        public MainWindow(ApplicationDbContext context)
        {
            InitializeComponent();
            _context = context;
            LoadData();
        }

        private void LoadData()
        {
            DataGridRecords.ItemsSource = _context.Records.ToList();
        }

        private void ImportCsv_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV files (*.csv)|*.csv";
            if (openFileDialog.ShowDialog() == true)
            {
                using (var reader = new StreamReader(openFileDialog.FileName))
                using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    Delimiter = ";"
                }))
                {
                    var records = csv.GetRecords<DataRecord>().ToList();
                    _context.Records.AddRange(records);
                    _context.SaveChanges();
                    LoadData();
                }
            }
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            if (saveFileDialog.ShowDialog() == true)
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Records");
                    var records = _context.Records.ToList();
                    worksheet.Cells.LoadFromCollection(records, true);
                    package.SaveAs(new FileInfo(saveFileDialog.FileName));
                }
            }
        }

        private void ExportToXml_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XML files (*.xml)|*.xml";
            if (saveFileDialog.ShowDialog() == true)
            {
                var records = _context.Records.ToList();
                var xDocument = new XDocument(
                    new XElement("TestProgram",
                        records.Select(r =>
                            new XElement("Record",
                                new XAttribute("id", r.Id),
                                new XElement("Date", r.Date),
                                new XElement("FirstName", r.FirstName),
                                new XElement("LastName", r.LastName),
                                new XElement("SurName", r.SurName),
                                new XElement("City", r.City),
                                new XElement("Country", r.Country)
                            )
                        )
                    )
                );
                xDocument.Save(saveFileDialog.FileName);
            }
        }
    }
}