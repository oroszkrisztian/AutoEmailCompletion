using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.EntityFrameworkCore;
using EmailCompleteApp.Models;
using EmailCompleteApp.Services.Repositories;

namespace EmailCompleteApp.Services
{
    /// <summary>
    /// Search service with in-memory cache for fast autocomplete
    /// - Loads all data from Supabase at startup (3 queries via repositories)
    /// - Searches happen in memory (FAST - no database queries)
    /// - Refreshes cache after inserts and email sent (3 queries via repositories)
    /// </summary>
    public class SearchService
    {
        private static SearchService? _instance;
        private static readonly object _lock = new object();

        // In-memory cache for fast searches
        private List<Client> _allClients = new();
        private List<Transportator> _allTransportators = new();
        private List<Location> _allLocations = new();
        private List<Product> _allProducts = new();
        private List<Contact> _allContacts = new();
        private bool _dataLoaded = false;

        // Events for UI progress updates
        public event Action<string>? ProgressChanged;
        public event Action<string>? DetailChanged;

        // Repository instances
        private readonly ClientRepository _clientRepo = ClientRepository.Instance;
        private readonly TransportatorRepository _transportatorRepo = TransportatorRepository.Instance;
        private readonly LocationRepository _locationRepo = LocationRepository.Instance;
        private readonly ProductRrepository _productRepo = ProductRrepository.Instance;
        private readonly ContactRepository _contactRepo = ContactRepository.Instance;

        public static SearchService Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        _instance ??= new SearchService();
                    }
                }
                return _instance;
            }
        }

        private SearchService() { }

        #region Connection & Setup

        /// <summary>
        /// Test connection to Supabase
        /// </summary>
        public async Task<bool> TestConnectionAsync()
        {
            try
            {


                using var context = DatabaseConfig.CreateDbContext();
                var canConnect = await context.Database.CanConnectAsync();

                if (canConnect)
                {
                    System.Diagnostics.Debug.WriteLine("✅ Supabase connection successful");
                }

                return canConnect;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Supabase connection error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Ensure database tables exist
        /// </summary>
        public async Task EnsureDatabaseCreatedAsync()
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                await context.Database.EnsureCreatedAsync();
                System.Diagnostics.Debug.WriteLine("✅ Supabase tables verified/created");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error creating tables: {ex.Message}");
                throw;
            }
        }

        #endregion

        #region Data Loading from Supabase

        /// <summary>
        /// 🔥 Load all data from Supabase at app startup (3 queries via repositories)
        /// </summary>
        public async Task LoadAllDataAsync()
        {
            if (_dataLoaded) return;

            await Task.Run(async () =>
            {
                try
                {
                    ProgressChanged?.Invoke("Connecting to Supabase...");
                    DetailChanged?.Invoke("Testing database connection");



                    if (!await TestConnectionAsync())
                    {
                        ProgressChanged?.Invoke("Connection failed");
                        DetailChanged?.Invoke("Could not connect to Supabase");

                        Application.Current.Dispatcher.Invoke(() =>
                            MessageBox.Show(
                                "Could not connect to Supabase!\n\n" +
                                "Check:\n" +
                                "1. Internet connection\n" +
                                "2. Password in DatabaseConfig.cs\n" +
                                "3. Supabase project is active",
                                "Connection Error",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error));
                        return;
                    }

                    ProgressChanged?.Invoke("Setting up database...");
                    await EnsureDatabaseCreatedAsync();

                    // 🔥 Load all data from Supabase via repositories (3 queries)
                    ProgressChanged?.Invoke("Loading clients...");
                    DetailChanged?.Invoke("Fetching clients from Supabase");
                    _allClients = await _clientRepo.LoadAllAsync();

                    ProgressChanged?.Invoke("Loading transportators...");
                    DetailChanged?.Invoke("Fetching transportators from Supabase");
                    _allTransportators = await _transportatorRepo.LoadAllAsync();

                    ProgressChanged?.Invoke("Loading locations...");
                    DetailChanged?.Invoke("Fetching locations from Supabase");
                    _allLocations = await _locationRepo.LoadAllAsync();

                    ProgressChanged?.Invoke("Loading products...");
                    DetailChanged?.Invoke("Fetching products from Supabase");
                   _allProducts = await _productRepo.LoadAllAsync();

                    ProgressChanged?.Invoke("Loading contacts...");
                    DetailChanged?.Invoke("Fetching contacts from Supabase");
                    _allContacts = await _contactRepo.LoadAllAsync();


                    _dataLoaded = true;

                    System.Diagnostics.Debug.WriteLine(
                        $"✅ Data loaded: {_allClients.Count} clients, " +
                        $"{_allTransportators.Count} transportators, " +
                        $"{_allLocations.Count} locations");

                    ProgressChanged?.Invoke("Ready!");
                    DetailChanged?.Invoke(
                        $"Loaded {_allClients.Count} clients, " +
                        $"{_allTransportators.Count} transportators, " +
                        $"{_allLocations.Count} locations");
                }
                catch (Exception ex)
                {
                    ProgressChanged?.Invoke("Error occurred");
                    DetailChanged?.Invoke($"Failed: {ex.Message}");

                    Application.Current.Dispatcher.Invoke(() =>
                        MessageBox.Show(
                            $"Error loading data from Supabase:\n\n{ex.Message}",
                            "Database Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error));
                }
            });
        }

        
        public async Task RefreshDataAsync()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("🔄 Refreshing data from Supabase...");

                ProgressChanged?.Invoke("Refreshing data...");

                _allClients = await _clientRepo.LoadAllAsync();
                _allTransportators = await _transportatorRepo.LoadAllAsync();
                _allLocations = await _locationRepo.LoadAllAsync();
                _allProducts = await _productRepo.LoadAllAsync();
                _allContacts = await _contactRepo.LoadAllAsync();

                System.Diagnostics.Debug.WriteLine(
                    $"✅ Refreshed: {_allClients.Count} clients, " +
                    $"{_allTransportators.Count} transportators, " +
                    $"{_allLocations.Count} locations");

                ProgressChanged?.Invoke("Refreshed!");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Refresh failed: {ex.Message}");
                throw;
            }
        }

        #endregion


        public async Task<List<Client>> SearchClientsAsync(string searchText)
        {
            await EnsureDataLoadedAsync();

            if (string.IsNullOrWhiteSpace(searchText))
                return _allClients.ToList();

            return _allClients
                .Where(c => c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        public async Task<List<Transportator>> SearchTransportatorsAsync(string searchText)
        {
            await EnsureDataLoadedAsync();

            if (string.IsNullOrWhiteSpace(searchText))
                return _allTransportators.ToList();

            return _allTransportators
                .Where(t => t.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        public async Task<List<Location>> SearchLocationsAsync(string searchText)
        {
            await EnsureDataLoadedAsync();

            if (string.IsNullOrWhiteSpace(searchText))
                return _allLocations.ToList();

            return _allLocations
                .Where(l =>
                    l.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    l.Address.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                    l.City.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        public async Task<List<Product>> SearchProductsAsync(string searchText)
        {
            await EnsureDataLoadedAsync();
            if (string.IsNullOrWhiteSpace(searchText))
                return _allProducts.ToList();
            return _allProducts
                .Where(p => p.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        public async Task<List<Contact>> SearchContactsAsync(string searchText)
        {
            await EnsureDataLoadedAsync();
            if (string.IsNullOrWhiteSpace(searchText))
                return _allContacts.ToList();
            return _allContacts
                .Where(c => c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase))                  
                .ToList();
        }


        private async Task EnsureDataLoadedAsync()
        {
            if (!_dataLoaded)
            {
                await LoadAllDataAsync();
            }
        }
    } 
}