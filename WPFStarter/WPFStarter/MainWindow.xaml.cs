using Microsoft.Win32;
using System.Windows;
using WPFStarter.Models.Data;
using WPFStarter.Services;

namespace WPFStarter
{
    public partial class MainWindow : Window
    {
        private readonly ApplicationDbContext _context;
        private readonly IDataWorker _dataWorker;

        public MainWindow(ApplicationDbContext context, IDataWorker dataWorker)
        {
            InitializeComponent();
            _context = context;
            _dataWorker = dataWorker;
            LoadData();
        }

        private void LoadData()
        {
            DataGridRecords.ItemsSource = _context.Records.ToList();
        }

        private void ImportCsv_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                _dataWorker.ImportCsv(_context, openFileDialog.FileName);
                LoadData();
            }
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog { Filter = "Excel Files|*.xlsx" };
            if (saveFileDialog.ShowDialog() == true)
            {
                _dataWorker.ExportToExcel(_context, saveFileDialog.FileName);
            }
        }

        private void ExportToXml_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog { Filter = "XML Files|*.xml" };
            if (saveFileDialog.ShowDialog() == true)
            {
                _dataWorker.ExportToXml(_context, saveFileDialog.FileName);
            }
        }
    }
}