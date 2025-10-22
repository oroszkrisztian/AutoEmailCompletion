using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using ClosedXML.Excel;
using EmailCompleteApp.Models;

namespace EmailCompleteApp.Services
{
    public class SearchService
    {
        private static SearchService? _instance;
        private static readonly object _lock = new object();

        private List<Client> _allClients = new();
        private List<Transportator> _allTransportators = new();
        private List<Location> _allLocations = new();
        private List<HistoryTransport> _allHistoryTransports = new();
        private bool _dataLoaded = false;

        public event Action<string>? ProgressChanged;
        public event Action<string>? DetailChanged;

        public static SearchService Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new SearchService();
                        }
                    }
                }
                return _instance;
            }
        }

        private SearchService()
        {
            
        }

        public async Task<List<string>> SearchClientNamesAsync(string searchText)
        {
            await EnsureDataLoadedAsync();

            if (string.IsNullOrWhiteSpace(searchText))
                return _allClients.Take(10).Select(c => c.Name).ToList();

            return _allClients
                .Where(client => client.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .Take(10)
                .Select(c => c.Name)
                .ToList();
        }

        public async Task<List<string>> SearchTransportatorNamesAsync(string searchText)
        {
            await EnsureDataLoadedAsync();

            if (string.IsNullOrWhiteSpace(searchText))
                return _allTransportators.Take(10).Select(t => t.Name).ToList();

            return _allTransportators
                .Where(transportator => transportator.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .Take(10)
                .Select(t => t.Name)
                .ToList();
        }

        public async Task<List<string>> SearchLocationAddressesAsync(string searchText)
        {
            await EnsureDataLoadedAsync();

            if (string.IsNullOrWhiteSpace(searchText))
                return _allLocations.Take(10).Select(l => l.Address).ToList();

            return _allLocations
                .Where(location => location.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) || 
                                  location.Address.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .Take(10)
                .Select(l => l.Address)
                .ToList();
        }

        public async Task<List<string>> GetAllClientNamesAsync()
        {
            await EnsureDataLoadedAsync();
            return _allClients.Select(c => c.Name).ToList();
        }

        public async Task<List<string>> GetAllTransportatorNamesAsync()
        {
            await EnsureDataLoadedAsync();
            return _allTransportators.Select(t => t.Name).ToList();
        }

        public async Task<List<string>> GetAllLocationAddressesAsync()
        {
            await EnsureDataLoadedAsync();
            return _allLocations.Select(l => l.Address).ToList();
        }

        

        private async Task EnsureDataLoadedAsync()
        {
            if (!_dataLoaded)
            {
                await LoadAllDataAsync();
            }
        }

        public async Task LoadAllDataAsync()
        {
            if (_dataLoaded) return;

            await Task.Run(() =>
            {
                try
                {
                    ProgressChanged?.Invoke("Initializing...");
                    DetailChanged?.Invoke("Locating database file");

                    string excelPath = GetDatabaseExcelPath();
                    
                    if (!File.Exists(excelPath))
                    {
                        ProgressChanged?.Invoke("Database not found");
                        DetailChanged?.Invoke($"File not found: {excelPath}");
                        
                        Application.Current.Dispatcher.Invoke(() =>
                            MessageBox.Show($"Database Excel file not found at: {excelPath}", 
                                           "Warning", MessageBoxButton.OK, MessageBoxImage.Warning));
                        return;
                    }

                    ProgressChanged?.Invoke("Opening database...");
                    DetailChanged?.Invoke("Reading Excel workbook");

                    using var workbook = new XLWorkbook(excelPath);
                    
                    // Load Clients
                    ProgressChanged?.Invoke("Loading clients...");
                    DetailChanged?.Invoke("Reading client data from Excel");
                    LoadClientsFromSheet(workbook, "Clients");
                    
                    // Load Transportators
                    ProgressChanged?.Invoke("Loading transportators...");
                    DetailChanged?.Invoke("Reading transportator data from Excel");
                    LoadTransportatorsFromSheet(workbook, "Transportators");
                    
                    // Load Locations
                    ProgressChanged?.Invoke("Loading locations...");
                    DetailChanged?.Invoke("Reading location data from Excel");
                    LoadLocationsFromSheet(workbook, "Locations");

                    //Load HistoryTransports
                    ProgressChanged?.Invoke("Loading history transports...");
                    DetailChanged?.Invoke("Reading history transport data from Excel");
                    LoadHistoryTransportsFromSheet(workbook, "HistoryTransports");

                    ProgressChanged?.Invoke("Finalizing...");
                    DetailChanged?.Invoke("Data loading complete");
                    
                    _dataLoaded = true;
                    
                    System.Diagnostics.Debug.WriteLine($"SearchService: Loaded {_allClients.Count} clients, {_allTransportators.Count} transportators, {_allLocations.Count} locations");
                    
                    ProgressChanged?.Invoke("Ready!");
                    DetailChanged?.Invoke($"Loaded {_allClients.Count} clients, {_allTransportators.Count} transportators, {_allLocations.Count} locations");
                }
                catch (Exception ex)
                {
                    ProgressChanged?.Invoke("Error occurred");
                    DetailChanged?.Invoke($"Failed to load data: {ex.Message}");
                    
                    Application.Current.Dispatcher.Invoke(() =>
                        MessageBox.Show($"An error occurred while loading data from Excel file: {ex.Message}", 
                                       "Error", MessageBoxButton.OK, MessageBoxImage.Error));
                }
            });
        }
        private void LoadHistoryTransportsFromSheet(XLWorkbook workbook, string sheetName)
        {
            try
            {
                var targetSheet = workbook.Worksheets.FirstOrDefault(s => 
                    s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                
                if (targetSheet == null)
                {
                    System.Diagnostics.Debug.WriteLine($"Sheet '{sheetName}' not found in Excel file");
                    DetailChanged?.Invoke($"Warning: Sheet '{sheetName}' not found");
                    return;
                }
                var lastRow = targetSheet.LastRowUsed()?.RowNumber() ?? 1;
                
                for (int row = 2; row <= lastRow; row++)
                {
                    try
                    {
                        var numarComanda = targetSheet.Cell(row, 1).GetValue<int>();
                        var transportName = targetSheet.Cell(row, 2).GetString();
                        var camClient = targetSheet.Cell(row, 3).GetString();
                        var route = targetSheet.Cell(row, 4).GetString();
                        var transportator = targetSheet.Cell(row, 5).GetString();
                        var dataTransport = targetSheet.Cell(row, 6).GetDateTime();

                        var historyTransport = new HistoryTransport(numarComanda, transportName.Trim(), camClient.Trim(), 
                                                                    route.Trim(), transportator.Trim(), dataTransport);

                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error loading history transport row {row}: {ex.Message}");
                        continue;
                    }
                }
                
                // Sort the list by date descending
                _allHistoryTransports.Sort((h1, h2) => h2.DataTransport.CompareTo(h1.DataTransport));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading data from sheet '{sheetName}': {ex.Message}");
                DetailChanged?.Invoke($"Error loading {sheetName}: {ex.Message}");
            }
        }
        private void LoadClientsFromSheet(XLWorkbook workbook, string sheetName)
        {
            try
            {
                var targetSheet = workbook.Worksheets.FirstOrDefault(s => 
                    s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                
                if (targetSheet == null)
                {
                    System.Diagnostics.Debug.WriteLine($"Sheet '{sheetName}' not found in Excel file");
                    DetailChanged?.Invoke($"Warning: Sheet '{sheetName}' not found");
                    return;
                }

                var lastRow = targetSheet.LastRowUsed()?.RowNumber() ?? 1;
                
                for (int row = 2; row <= lastRow; row++)
                {
                    try
                    {
                        var id = targetSheet.Cell(row, 1).GetValue<int>();
                        var name = targetSheet.Cell(row, 2).GetString();
                        var address = targetSheet.Cell(row, 3).GetString();
                        var bank = targetSheet.Cell(row, 4).GetString();
                        var iban = targetSheet.Cell(row, 5).GetString();
                        var vat = targetSheet.Cell(row, 6).GetString();
                        var camera = targetSheet.Cell(row, 7).GetString();
                        var termen = targetSheet.Cell(row, 8).GetString();
                        
                        if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(address))
                        {
                            var client = new Client(name.Trim(), address.Trim(), bank?.Trim() ?? "", iban?.Trim() ?? "",
                                                  vat?.Trim() ?? "", camera?.Trim() ?? "", termen?.Trim() ?? "");
                            client.Id = id;
                            _allClients.Add(client);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error loading client row {row}: {ex.Message}");
                        continue;
                    }
                }
                
                // Sort the list 
                _allClients.Sort((c1, c2) => string.Compare(c1.Name, c2.Name, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading data from sheet '{sheetName}': {ex.Message}");
                DetailChanged?.Invoke($"Error loading {sheetName}: {ex.Message}");
            }
        }

        private void LoadTransportatorsFromSheet(XLWorkbook workbook, string sheetName)
        {
            try
            {
                var targetSheet = workbook.Worksheets.FirstOrDefault(s => 
                    s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                
                if (targetSheet == null)
                {
                    System.Diagnostics.Debug.WriteLine($"Sheet '{sheetName}' not found in Excel file");
                    DetailChanged?.Invoke($"Warning: Sheet '{sheetName}' not found");
                    return;
                }

                var lastRow = targetSheet.LastRowUsed()?.RowNumber() ?? 1;
                
                for (int row = 2; row <= lastRow; row++)
                {
                    try
                    {
                        var id = targetSheet.Cell(row, 1).GetValue<int>();
                        var name = targetSheet.Cell(row, 2).GetString();
                        var address = targetSheet.Cell(row, 3).GetString();
                        var bank = targetSheet.Cell(row, 4).GetString();
                        var iban = targetSheet.Cell(row, 5).GetString();
                        var vat = targetSheet.Cell(row, 6).GetString();
                        var camera = targetSheet.Cell(row, 7).GetString();
                        var termen = targetSheet.Cell(row, 8).GetString();
                        
                        if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(address))
                        {
                            var transportator = new Transportator(name.Trim(), address.Trim(), bank?.Trim() ?? "", iban?.Trim() ?? "",
                                                                vat?.Trim() ?? "", camera?.Trim() ?? "", termen?.Trim() ?? "");
                            transportator.Id = id;
                            _allTransportators.Add(transportator);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error loading transportator row {row}: {ex.Message}");
                        continue;
                    }
                }
                
                // Sort the list for better user experience
                _allTransportators.Sort((t1, t2) => string.Compare(t1.Name, t2.Name, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading data from sheet '{sheetName}': {ex.Message}");
                DetailChanged?.Invoke($"Error loading {sheetName}: {ex.Message}");
            }
        }

        private void LoadLocationsFromSheet(XLWorkbook workbook, string sheetName)
        {
            try
            {
                var targetSheet = workbook.Worksheets.FirstOrDefault(s => 
                    s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                
                if (targetSheet == null)
                {
                    System.Diagnostics.Debug.WriteLine($"Sheet '{sheetName}' not found in Excel file");
                    DetailChanged?.Invoke($"Warning: Sheet '{sheetName}' not found");
                    return;
                }

                var lastRow = targetSheet.LastRowUsed()?.RowNumber() ?? 1;
                
                for (int row = 2; row <= lastRow; row++)
                {
                    try
                    {
                        var id = targetSheet.Cell(row, 1).GetValue<int>();
                        var name = targetSheet.Cell(row, 2).GetString();
                        var address = targetSheet.Cell(row, 3).GetString();
                        
                        if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(address))
                        {
                            var location = new Location(id, name.Trim(), address.Trim());
                            _allLocations.Add(location);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error loading location row {row}: {ex.Message}");
                        continue;
                    }
                }
                
                // Sort the list for better user experience
                _allLocations.Sort((l1, l2) => string.Compare(l1.Name, l2.Name, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading data from sheet '{sheetName}': {ex.Message}");
                DetailChanged?.Invoke($"Error loading {sheetName}: {ex.Message}");
            }
        }

        private static string GetDatabaseExcelPath()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            string? current = baseDir;
            for (int i = 0; i < 6 && current != null; i++)
            {
                string docDir = Path.Combine(current, "doc");
                if (Directory.Exists(docDir))
                {
                    return Path.Combine(docDir, "database.xlsx");
                }
                current = Directory.GetParent(current)?.FullName;
            }

            var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(docs, "AutoEmailCompletion", "database.xlsx");
        }

        /// <summary>
        /// Refresh data from Excel file (useful if the Excel file has been updated)
        /// </summary>
        public async Task RefreshDataAsync()
        {
            _allClients.Clear();
            _allTransportators.Clear();
            _allLocations.Clear();
            _dataLoaded = false;
            
            await LoadAllDataAsync();
        }

        /// <summary>
        /// Get all client objects
        /// </summary>
        public async Task<List<Client>> GetAllClientsAsync()
        {
            await EnsureDataLoadedAsync();
            return _allClients.ToList();
        }

        /// <summary>
        /// Get all transportator objects
        /// </summary>
        public async Task<List<Transportator>> GetAllTransportatorsAsync()
        {
            await EnsureDataLoadedAsync();
            return _allTransportators.ToList();
        }

        /// <summary>
        /// Get all location objects
        /// </summary>
        public async Task<List<Location>> GetAllLocationsAsync()
        {
            await EnsureDataLoadedAsync();
            return _allLocations.ToList();
        }

        /// <summary>
        /// Get all history objects
        /// </summary>
        public async Task<List<HistoryTransport>> GetAllHistoryTransportsAsync()
        {
            await EnsureDataLoadedAsync();
            return _allHistoryTransports.ToList();
        }

        /// <summary>
        /// Get client by name
        /// </summary>
        public async Task<Client?> GetClientByNameAsync(string name)
        {
            await EnsureDataLoadedAsync();
            return _allClients.FirstOrDefault(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// getClientCmareDeComert by name
        /// </summary>
        internal async Task<string> GetClientCameraDeComert(string clientName)
        {
            await EnsureDataLoadedAsync();
            var client = _allClients.FirstOrDefault(c => c.Name.Equals(clientName, StringComparison.OrdinalIgnoreCase));
            return client?.CameraDeComert ?? string.Empty;
        }
        /// <summary>
        /// Get transportator by name
        /// </summary>
        public async Task<Transportator?> GetTransportatorByNameAsync(string name)
        {
            await EnsureDataLoadedAsync();
            return _allTransportators.FirstOrDefault(t => t.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

       

        /// <summary>
        /// Add a client to the in-memory list (call this after successfully inserting to Excel)
        /// </summary>
        public async Task AddClientToMemoryAsync(Client client)
        {
            await EnsureDataLoadedAsync(); // Ensure data is loaded first
            if (_allClients.Any(c => c.Id == client.Id)) return; // Avoid duplicates
            _allClients.Add(client);
            _allClients.Sort((c1, c2) => string.Compare(c1.Name, c2.Name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Add a transportator to the in-memory list (call this after successfully inserting to Excel)
        /// </summary>
        public async Task AddTransportatorToMemoryAsync(Transportator transportator)
        {
            await EnsureDataLoadedAsync(); // Ensure data is loaded first
            if (_allTransportators.Any(t => t.Id == transportator.Id)) return; // Avoid duplicates
            _allTransportators.Add(transportator);
            _allTransportators.Sort((t1, t2) => string.Compare(t1.Name, t2.Name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Add a location to the in-memory list (call this after successfully inserting to Excel)
        /// </summary>
        public async Task AddLocationToMemoryAsync(Location location)
        {
            await EnsureDataLoadedAsync(); // Ensure data is loaded first
            if (_allLocations.Any(l => l.Id == location.Id)) return; // Avoid duplicates
            _allLocations.Add(location);
            _allLocations.Sort((l1, l2) => string.Compare(l1.Name, l2.Name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Add a comanda  to the history (call this after successfully inserting to Excel)
        /// </summary>
        public  async  Task AddHistoryTransportToMemoryAsync(HistoryTransport historyTransport)
        {
            await EnsureDataLoadedAsync(); 
            _allHistoryTransports.Add(historyTransport);
            _allHistoryTransports.Sort((h1, h2) => h2.DataTransport.CompareTo(h1.DataTransport));
        }

        
    }
}