using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.Windows;
using WPFStarter.Models;
using WPFStarter.Services;

namespace WPFStarter.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private readonly ApplicationDbContext _context;
        private readonly IDataWorker _dataWorker;

        public ObservableCollection<DataRecord> Records { get; set; }
        public ObservableCollection<DataRecord> FilteredRecords { get; set; }

        private DateTime? _startDate;
        public DateTime? StartDate
        {
            get => _startDate;
            set
            {
                if (SetProperty(ref _startDate, value))
                {
                    ApplyFilters();
                }
            }
        }

        private DateTime? _endDate;
        public DateTime? EndDate
        {
            get => _endDate;
            set
            {
                if (SetProperty(ref _endDate, value))
                {
                    ApplyFilters();
                }
            }
        }

        private string _city;
        public string City
        {
            get => _city;
            set
            {
                if (SetProperty(ref _city, value))
                {
                    ApplyFilters();
                }
            }
        }

        private string _lastName;
        public string LastName
        {
            get => _lastName;
            set
            {
                if (SetProperty(ref _lastName, value))
                {
                    ApplyFilters();
                }
            }
        }

        private string _country;
        public string Country
        {
            get => _country;
            set
            {
                if (SetProperty(ref _country, value))
                {
                    ApplyFilters();
                }
            }
        }

        public RelayCommand ImportCsvCommand { get; }
        public RelayCommand ExportToExcelCommand { get; }
        public RelayCommand ExportToXmlCommand { get; }

        public MainViewModel(ApplicationDbContext context, IDataWorker dataWorker)
        {
            _context = context;
            _dataWorker = dataWorker;

            Records = new ObservableCollection<DataRecord>(_context.Records.ToList());
            FilteredRecords = new ObservableCollection<DataRecord>(Records);

            ImportCsvCommand = new RelayCommand(async _ => await ImportCsvAsync());
            ExportToExcelCommand = new RelayCommand(async _ => await ExportToExcelAsync());
            ExportToXmlCommand = new RelayCommand(async _ => await ExportToXmlAsync());
        }

        private async Task ImportCsvAsync()
        {
            try
            {
                var openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == true)
                {
                    await _dataWorker.ImportCsvAsync(_context, openFileDialog.FileName);
                    LoadData();
                    MessageBox.Show("CSV file imported successfully");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error importing CSV file: {ex.Message}");
            }
        }

        private async Task ExportToExcelAsync()
        {
            try
            {
                var saveFileDialog = new SaveFileDialog { Filter = "Excel Files|*.xlsx" };
                if (saveFileDialog.ShowDialog() == true)
                {
                    var filteredRecords = GetFilteredRecords();
                    await _dataWorker.ExportToExcelAsync(filteredRecords, saveFileDialog.FileName);
                    MessageBox.Show("Data exported to Excel successfully");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to Excel file: {ex.Message}");
            }
        }

        private async Task ExportToXmlAsync()
        {
            try
            {
                var saveFileDialog = new SaveFileDialog { Filter = "XML Files|*.xml" };
                if (saveFileDialog.ShowDialog() == true)
                {
                    var filteredRecords = GetFilteredRecords();
                    await _dataWorker.ExportToXmlAsync(filteredRecords, saveFileDialog.FileName);
                    MessageBox.Show("Data exported to XML successfully");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to XML file: {ex.Message}");
            }
        }

        private IQueryable<DataRecord> GetFilteredRecords()
        {
            var filtered = Records.AsQueryable();

            if (StartDate.HasValue)
                filtered = filtered.Where(r => r.Date >= StartDate.Value);

            if (EndDate.HasValue)
                filtered = filtered.Where(r => r.Date <= EndDate.Value);

            if (!string.IsNullOrEmpty(City))
                filtered = filtered.Where(r => r.City.Contains(City, StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrEmpty(LastName))
                filtered = filtered.Where(r => r.LastName.Contains(LastName, StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrEmpty(Country))
                filtered = filtered.Where(r => r.Country.Contains(Country, StringComparison.OrdinalIgnoreCase));

            return filtered;
        }

        private void ApplyFilters()
        {
            var filtered = GetFilteredRecords().ToList();

            FilteredRecords.Clear();
            foreach (var record in filtered)
            {
                FilteredRecords.Add(record);
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
            ApplyFilters();
        }
    }
}