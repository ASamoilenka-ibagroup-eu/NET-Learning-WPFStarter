using Microsoft.Win32;
using System.Collections.ObjectModel;
using WPFStarter.Models;
using WPFStarter.Services;

namespace WPFStarter.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private readonly ApplicationDbContext _context;
        private readonly IDataWorker _dataWorker;

        public ObservableCollection<DataRecord> Records { get; set; }

        public RelayCommand ImportCsvCommand { get; }
        public RelayCommand ExportToExcelCommand { get; }
        public RelayCommand ExportToXmlCommand { get; }

        public MainViewModel(ApplicationDbContext context, IDataWorker dataWorker)
        {
            _context = context;
            _dataWorker = dataWorker;

            Records = new ObservableCollection<DataRecord>(_context.Records.ToList());

            ImportCsvCommand = new RelayCommand(_ => ImportCsv());
            ExportToExcelCommand = new RelayCommand(_ => ExportToExcel());
            ExportToXmlCommand = new RelayCommand(_ => ExportToXml());
        }

        private void ImportCsv()
        {
            var openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                _dataWorker.ImportCsv(_context, openFileDialog.FileName);
                LoadData();
            }
        }

        private void ExportToExcel()
        {
            var saveFileDialog = new SaveFileDialog { Filter = "Excel Files|*.xlsx" };
            if (saveFileDialog.ShowDialog() == true)
            {
                _dataWorker.ExportToExcel(_context, saveFileDialog.FileName);
            }
        }

        private void ExportToXml()
        {
            var saveFileDialog = new SaveFileDialog { Filter = "XML Files|*.xml" };
            if (saveFileDialog.ShowDialog() == true)
            {
                _dataWorker.ExportToXml(_context, saveFileDialog.FileName);
            }
        }

        private void LoadData()
        {
            Records.Clear();
            var records = _context.Records.ToList();
            foreach (var record in records)
            {
                Records.Add(record);
            }
        }
    }
}