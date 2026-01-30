using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using EmailCompleteApp.Models;
using EmailCompleteApp.Services;
using EmailCompleteApp.Services.Repositories;
using Microsoft.EntityFrameworkCore.Metadata;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace EmailCompleteApp.ViewModels;

public partial class ComandaTransportViewModel : ObservableObject
{
    private readonly SearchService _searchService;
    private readonly DocumentCompletion _documentCompletion;
    private readonly HistoryRepository _historyRepository;
    private readonly ContactRepository _contactRepository ;
    private readonly ProductRrepository _productRepository;

    // Search debounce settings
    private const int SearchDebounceDelayMs = 300;
    private const int InitialSuggestions = 10;

    // Cancellation tokens for search operations
    private CancellationTokenSource? _clientSearchCts;
    private CancellationTokenSource? _transportatorSearchCts;
    private CancellationTokenSource? _incarcareSearchCts;
    private CancellationTokenSource? _descarcareSearchCts;
    private CancellationTokenSource? _productSearchCts;
    private CancellationTokenSource? _contactSearchCts;

    // Flag to prevent search when updating from selection
    private bool _isUpdatingFromSelection = false;

    // Edit mode original history item
    private HistoryTransport? _editingHistoryItem = null;

    #region NumarComanda Helper
    private HistoryTransport _lastOrder = new HistoryTransport();

    #endregion


    #region Observable Properties
    [ObservableProperty] private string _lastOrderString = "";
    [ObservableProperty] private string _numarComanda = string.Empty;
    [ObservableProperty] private string _numarClient = string.Empty;
    [ObservableProperty] private string _client = string.Empty;
    [ObservableProperty] private string _tarif = string.Empty;
    [ObservableProperty] private int _monedaIndex = 0;
    [ObservableProperty] private int _tipIndex = 0;
    [ObservableProperty] private string _transportator = string.Empty;
    [ObservableProperty] private string _transportatorTarif = string.Empty;
    [ObservableProperty] private int _transportatorMonedaIndex = 0;
    [ObservableProperty] private int _transportatorTipIndex = 0;
    [ObservableProperty] private DateTime? _dataIncarcare = DateTime.UtcNow.Date;
    [ObservableProperty] private DateTime? _dataDescarcare = DateTime.UtcNow.Date.AddDays(1);
    [ObservableProperty] private string _produs = string.Empty;
    [ObservableProperty] private string _contact = string.Empty;
    [ObservableProperty] private string _cantitate = string.Empty;
    [ObservableProperty] private int _tipAdrIndex = 0;
    [ObservableProperty] private string _clasa = string.Empty;
    [ObservableProperty] private string _un = string.Empty;
    [ObservableProperty] private string _numarInmatriculare = string.Empty;
    [ObservableProperty] private string _locatieIncarcare = string.Empty;
    [ObservableProperty] private string _locatieDescarcare = string.Empty;
    [ObservableProperty] private string _termenPlata = string.Empty;
    [ObservableProperty] private bool _isClientDropDownOpen = false;
    [ObservableProperty] private bool _isTransportatorDropDownOpen = false;
    [ObservableProperty] private bool _isIncarcareDropDownOpen = false;
    [ObservableProperty] private bool _isDescarcareDropDownOpen = false;
    [ObservableProperty] private bool _isSendingEmail = false;
    [ObservableProperty] private bool _isFormVisible = false;

    // Edit mode properties
    [ObservableProperty] private bool _isEditMode;
    [ObservableProperty] private string _submitButtonText = "Send Email";

    // Computed property for Send Email button visibility
    public bool ShowSendEmailButton => IsFormVisible && !IsEditMode;
    
    partial void OnIsFormVisibleChanged(bool value)
    {
        OnPropertyChanged(nameof(ShowSendEmailButton));
    }
    
    partial void OnIsEditModeChanged(bool value)
    {
        OnPropertyChanged(nameof(ShowSendEmailButton));
    }

    // Additional fields required by DocumentCompletion
    [ObservableProperty] private string _commentUser = string.Empty;

    // Pickup location components
    [ObservableProperty] private string _locatieIncarcareAddress = string.Empty;
    [ObservableProperty] private string _locatieIncarcareName = string.Empty;
    [ObservableProperty] private string _locatieIncarcareCity = string.Empty;
    [ObservableProperty] private string _locatieIncarcareCountryCode = string.Empty;
    [ObservableProperty] private string _locatieIncarcarePostalCode = string.Empty;
    [ObservableProperty] private string _locatieIncarcareCounty = string.Empty;

    // Delivery location components
    [ObservableProperty] private string _locatieDescarcareAddress = string.Empty;
    [ObservableProperty] private string _locatieDescarcareName = string.Empty;
    [ObservableProperty] private string _locatieDescarcareCity = string.Empty;
    [ObservableProperty] private string _locatieDescarcareCountryCode = string.Empty;
    [ObservableProperty] private string _locatieDescarcarePostalCode = string.Empty;
    [ObservableProperty] private string _locatieDescarcareCounty = string.Empty;

    #endregion

    #region Collections

    public ObservableCollection<Client> ClientSuggestions { get; } = new();
    public ObservableCollection<Transportator> TransportatorSuggestions { get; } = new();
    public ObservableCollection<Location> IncarcareSuggestions { get; } = new();
    public ObservableCollection<Location> DescarcareSuggestions { get; } = new();
    public ObservableCollection<Product> ProductSuggestions { get; } = new();
    public ObservableCollection<Contact> ContactSuggestions { get; } = new();

    #endregion

    #region Constructor

    public ComandaTransportViewModel()
    {
        _searchService = SearchService.Instance;
        _documentCompletion = DocumentCompletion.Instance;
        _historyRepository = HistoryRepository.Instance;
        _contactRepository = ContactRepository.Instance;
        _productRepository = ProductRrepository.Instance;
        IsFormVisible = false;
        IsEditMode = false;
        SubmitButtonText = "Send Email";
        _ = InitializeSuggestionsAsync();
    }

    // Constructor for edit mode
    public ComandaTransportViewModel(HistoryTransport historyItem)
    {
        _searchService = SearchService.Instance;
        _documentCompletion = DocumentCompletion.Instance;
        _historyRepository = HistoryRepository.Instance;
        _contactRepository = ContactRepository.Instance;
        _productRepository = ProductRrepository.Instance;

        _editingHistoryItem = historyItem;
        IsEditMode = true;
        IsFormVisible = true; // Show form immediately in edit mode
        SubmitButtonText = "Update";

        _ = InitializeSuggestionsAsync();
        _ = LoadHistoryDataForEdit(historyItem);
    }

    #endregion

    #region Initialization

    private async Task LoadHistoryDataForEdit(HistoryTransport historyItem)
    {
        await Task.Delay(100); 

        // Load all data from history item
        NumarComanda = historyItem.NumarComanda ?? string.Empty;
        NumarClient = historyItem.NumarClient ?? string.Empty;
        Client = historyItem.Client ?? string.Empty;
        Tarif = historyItem.Tarif?.ToString() ?? string.Empty;
        MonedaIndex = historyItem.MonedaIndex ?? 0;
        TipIndex = historyItem.TipIndex ?? 0;

        Transportator = historyItem.Transportator ?? string.Empty;
        TransportatorTarif = historyItem.TransportatorTarif?.ToString() ?? string.Empty;
        TransportatorMonedaIndex = historyItem.TransportatorMonedaIndex ?? 0;
        TransportatorTipIndex = historyItem.TransportatorTipIndex ?? 0;

        DataIncarcare = historyItem.DataIncarcare ?? DateTime.UtcNow.Date;
        DataDescarcare = historyItem.DataDescarcare ?? DateTime.UtcNow.Date.AddDays(1);

        Produs = historyItem.Produs ?? string.Empty;
        Cantitate = historyItem.Cantitate?.ToString() ?? string.Empty;
        TipAdrIndex = historyItem.TipAdrIndex ?? 0;
        Clasa = historyItem.Clasa ?? string.Empty;
        Un = historyItem.Un ?? string.Empty;
        NumarInmatriculare = historyItem.NumarInmatriculare ?? string.Empty;

        // Pickup location
        LocatieIncarcareAddress = historyItem.LocatieIncarcareAddress ?? string.Empty;
        LocatieIncarcareName = historyItem.LocatieIncarcareName ?? string.Empty;
        LocatieIncarcareCity = historyItem.LocatieIncarcareCity ?? string.Empty;
        LocatieIncarcareCountryCode = historyItem.LocatieIncarcareCountryCode ?? string.Empty;
        LocatieIncarcarePostalCode = historyItem.LocatieIncarcarePostalCode ?? string.Empty;
        LocatieIncarcareCounty = historyItem.LocatieIncarcareCounty ?? string.Empty;

        // Build display string for pickup location
        var pickupParts = new[] {
            LocatieIncarcareName,
            LocatieIncarcareAddress,
            LocatieIncarcareCity,
            LocatieIncarcareCounty,
            LocatieIncarcarePostalCode,
            LocatieIncarcareCountryCode
        }.Where(p => !string.IsNullOrWhiteSpace(p));
        LocatieIncarcare = string.Join(", ", pickupParts);

        // Delivery location
        LocatieDescarcareAddress = historyItem.LocatieDescarcareAddress ?? string.Empty;
        LocatieDescarcareName = historyItem.LocatieDescarcareName ?? string.Empty;
        LocatieDescarcareCity = historyItem.LocatieDescarcareCity ?? string.Empty;
        LocatieDescarcareCountryCode = historyItem.LocatieDescarcareCountryCode ?? string.Empty;
        LocatieDescarcarePostalCode = historyItem.LocatieDescarcarePostalCode ?? string.Empty;
        LocatieDescarcareCounty = historyItem.LocatieDescarcareCounty ?? string.Empty;

        // Build display string for delivery location
        var deliveryParts = new[] {
            LocatieDescarcareName,
            LocatieDescarcareAddress,
            LocatieDescarcareCity,
            LocatieDescarcareCounty,
            LocatieDescarcarePostalCode,
            LocatieDescarcareCountryCode
        }.Where(p => !string.IsNullOrWhiteSpace(p));
        LocatieDescarcare = string.Join(", ", deliveryParts);

        TermenPlata = historyItem.TermenPlata?.ToString() ?? string.Empty;
        CommentUser = historyItem.CommentUser ?? string.Empty;

        Debug.WriteLine($"✅ Loaded history data for editing: Order #{NumarComanda}");
    }

    private async Task InitializeSuggestionsAsync()
    {
        try
        {

            await Task.WhenAll(
                InitLastOrder(),
                _searchService.LoadAllDataAsync(),
                LoadInitialClientsAsync(),
                LoadInitialTransportatorsAsync(),
                LoadInitialLocationsAsync(IncarcareSuggestions),
                LoadInitialLocationsAsync(DescarcareSuggestions),
                LoadInitialProductsAsync(),
                LoadInitialContactsAsync()

            );

            Debug.WriteLine("✅ Initial suggestions loaded from Supabase cache");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Error initializing suggestions: {ex.Message}");
        }
    }

    #endregion

    #region Last Order Helper

    private async Task InitLastOrder()
    {
        // Skip in edit mode
        if (IsEditMode) return;

        _lastOrder = await _historyRepository.GetLastOrder();
        LastOrderString = _lastOrder.HistorySummary();

        // Parse NumarComanda as int and increment
        int nextOrderNum = 1;
        if (int.TryParse(_lastOrder.NumarComanda, out var lastNum))
        {
            nextOrderNum = lastNum + 1;
        }

        if (Application.Current?.Dispatcher?.CheckAccess() == true)
            NumarComanda = nextOrderNum.ToString();
        else
            Application.Current?.Dispatcher?.Invoke(() => NumarComanda = nextOrderNum.ToString());
    }

    [RelayCommand(CanExecute = nameof(CanReloadLastOrder))]
    private async Task ReloadLastOrder()
    {
        await InitLastOrder();
    }

    private bool CanReloadLastOrder()
    {
        return !IsSendingEmail && !IsEditMode;
    }

    #endregion

    #region Form Visibility Toggle

    [RelayCommand]
    private async Task NewOrder()
    {
        // Clean up Email folder before starting new order
        CleanupEmailFolder();

        await InitLastOrder();
        ToggleVisibility();
    }

    private void ToggleVisibility()
    {
        IsFormVisible = !IsFormVisible;
    }

    /// <summary>
    /// Cleans up all .doc and .docx files in the Email folder when starting a new order
    /// </summary>
    private void CleanupEmailFolder()
    {
        try
        {
            string projectRoot = AppDomain.CurrentDomain.BaseDirectory;
            string projectDir = FindProjectDirectory(projectRoot);
            string docDir = System.IO.Path.Combine(projectDir, "doc");
            string emailDir = System.IO.Path.Combine(docDir, "Email");

            if (!System.IO.Directory.Exists(emailDir))
            {
                Debug.WriteLine($"📁 Email folder doesn't exist yet: {emailDir}");
                return;
            }

            // Get all .doc and .docx files in the Email folder
            var docFiles = System.IO.Directory.GetFiles(emailDir, "*.doc", System.IO.SearchOption.TopDirectoryOnly);
            var docxFiles = System.IO.Directory.GetFiles(emailDir, "*.docx", System.IO.SearchOption.TopDirectoryOnly);
            var allFiles = docFiles.Concat(docxFiles).ToArray();

            if (allFiles.Length == 0)
            {
                Debug.WriteLine($"✅ Email folder is already clean");
                return;
            }

            Debug.WriteLine($"🧹 Cleaning up {allFiles.Length} old file(s) from Email folder...");

            int deletedCount = 0;
            foreach (var file in allFiles)
            {
                try
                {
                    System.IO.File.Delete(file);
                    deletedCount++;
                    Debug.WriteLine($"  🗑️ Deleted: {System.IO.Path.GetFileName(file)}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"  ⚠️ Could not delete {System.IO.Path.GetFileName(file)}: {ex.Message}");
                }
            }

            Debug.WriteLine($"✅ Cleaned up {deletedCount} file(s) from Email folder");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"⚠️ Error during Email folder cleanup: {ex.Message}");
        }
    }

    #endregion

    #region Property Change Handlers

    partial void OnClientChanged(string value)
    {
        if (_isUpdatingFromSelection) return;

        _clientSearchCts?.Cancel();
        _clientSearchCts = new CancellationTokenSource();

        if (string.IsNullOrWhiteSpace(value))
            _ = LoadInitialClientsAsync();
        else
            _ = SearchWithDebounceAsync(value, SearchClientAsync, _clientSearchCts);
    }

    partial void OnTransportatorChanged(string value)
    {
        if (_isUpdatingFromSelection) return;

        _transportatorSearchCts?.Cancel();
        _transportatorSearchCts = new CancellationTokenSource();

        if (string.IsNullOrWhiteSpace(value))
            _ = LoadInitialTransportatorsAsync();
        else
            _ = SearchWithDebounceAsync(value, SearchTransportatorAsync, _transportatorSearchCts);
    }

    partial void OnLocatieIncarcareChanged(string value)
    {
        if (_isUpdatingFromSelection) return;

        _incarcareSearchCts?.Cancel();
        _incarcareSearchCts = new CancellationTokenSource();

        if (string.IsNullOrWhiteSpace(value))
            _ = LoadInitialLocationsAsync(IncarcareSuggestions);
        else
            _ = SearchLocationWithDebounceAsync(value, IncarcareSuggestions, _incarcareSearchCts);
    }

    partial void OnLocatieDescarcareChanged(string value)
    {
        if (_isUpdatingFromSelection) return;

        _descarcareSearchCts?.Cancel();
        _descarcareSearchCts = new CancellationTokenSource();

        if (string.IsNullOrWhiteSpace(value))
            _ = LoadInitialLocationsAsync(DescarcareSuggestions);
        else
            _ = SearchLocationWithDebounceAsync(value, DescarcareSuggestions, _descarcareSearchCts);
    }

    partial void OnNumarComandaChanged(string value)
    {
        SendEmailCommand.NotifyCanExecuteChanged();
        UpdateOrderCommand.NotifyCanExecuteChanged();
    }



    #endregion

    #region Debounced Search

    private async Task SearchWithDebounceAsync(
        string searchText,
        Func<string, CancellationToken, Task> searchAction,
        CancellationTokenSource cancellationTokenSource)
    {
        try
        {
            await Task.Delay(SearchDebounceDelayMs, cancellationTokenSource.Token);
            await searchAction(searchText, cancellationTokenSource.Token);
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Search error: {ex.Message}");
        }
    }

    private async Task SearchLocationWithDebounceAsync(
        string searchText,
        ObservableCollection<Location> targetCollection,
        CancellationTokenSource cancellationTokenSource)
    {
        try
        {
            await Task.Delay(SearchDebounceDelayMs, cancellationTokenSource.Token);
            await SearchLocationAsync(searchText, targetCollection, cancellationTokenSource.Token);
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Location search error: {ex.Message}");
        }
    }

    

    #endregion

    #region Search Methods (In-Memory)

    private async Task SearchClientAsync(string searchText, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(searchText) || cancellationToken.IsCancellationRequested)
            return;

        try
        {
            var results = await _searchService.SearchClientsAsync(searchText.Trim());

            if (!cancellationToken.IsCancellationRequested)
                UpdateCollection(ClientSuggestions, results);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Client search error: {ex.Message}");
        }
    }

    private async Task SearchTransportatorAsync(string searchText, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(searchText) || cancellationToken.IsCancellationRequested)
            return;

        try
        {
            var results = await _searchService.SearchTransportatorsAsync(searchText.Trim());

            if (!cancellationToken.IsCancellationRequested)
                UpdateCollection(TransportatorSuggestions, results);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Transportator search error: {ex.Message}");
        }
    }

    private async Task SearchLocationAsync(string searchText, ObservableCollection<Location> targetCollection, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(searchText) || cancellationToken.IsCancellationRequested)
            return;

        try
        {
            var results = await _searchService.SearchLocationsAsync(searchText.Trim());

            if (!cancellationToken.IsCancellationRequested)
                UpdateCollection(targetCollection, results);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Location search error: {ex.Message}");
        }
    }

    
    
    #endregion

    #region Load Initial Data

    private async Task LoadInitialClientsAsync()
    {
        try
        {
            var results = await _searchService.SearchClientsAsync("");
            UpdateCollection(ClientSuggestions, results.ToList());
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Load clients error: {ex.Message}");
        }
    }

    private async Task LoadInitialTransportatorsAsync()
    {
        try
        {
            var results = await _searchService.SearchTransportatorsAsync("");
            UpdateCollection(TransportatorSuggestions, results.ToList());
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Load transportators error: {ex.Message}");
        }
    }

    private async Task LoadInitialLocationsAsync(ObservableCollection<Location> targetCollection)
    {
        try
        {
            var results = await _searchService.SearchLocationsAsync("");
            UpdateCollection(targetCollection, results.ToList());
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Load locations error: {ex.Message}");
        }
    }

    private async Task LoadInitialProductsAsync()
    {
        try
        {
            var results = await _searchService.SearchProductsAsync("");
            UpdateCollection(ProductSuggestions, results.ToList());
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Load products error: {ex.Message}");
        }
    }

    private async Task LoadInitialContactsAsync()
    {
        try
        {
            var results = await _searchService.SearchContactsAsync("");
            UpdateCollection(ContactSuggestions, results.ToList());
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Load contacts error: {ex.Message}");
        }
    }

    #endregion

    #region Collection Helper

    private void UpdateCollection<T>(ObservableCollection<T> collection, List<T> items)
    {
        collection.Clear();
        foreach (var item in items)
            collection.Add(item);
    }

    #endregion

    #region Cancel Command (Edit Mode)

    [RelayCommand]
    private void CancelEdit()
    {
        // Navigate back to history page
        var mainWindow = Application.Current.MainWindow as MainWindow;
        mainWindow?.NavigateToHistory();
    }

    #endregion

    #region Update Command (Edit Mode)

    [RelayCommand(CanExecute = nameof(CanUpdateOrder))]
    private async Task UpdateOrder()
    {
        try
        {
            IsSendingEmail = true;

            Debug.WriteLine($"📝 Updating order #{NumarComanda}...");

            // Parse tarif values as decimals
            decimal? tarifValue = decimal.TryParse(Tarif, out var t) ? t : (decimal?)null;
            decimal? transportatorTarifDecimal = decimal.TryParse(TransportatorTarif, out var tt) ? tt : (decimal?)null;
            int? termenPlataValue = int.TryParse(TermenPlata, out var tp) ? tp : (int?)null;

            // Update the existing history record
            if (_editingHistoryItem != null)
            {
                _editingHistoryItem.NumarComanda = NumarComanda;
                _editingHistoryItem.NumarClient = NumarClient;
                _editingHistoryItem.Client = Client;
                _editingHistoryItem.Tarif = tarifValue;
                _editingHistoryItem.MonedaIndex = MonedaIndex;
                _editingHistoryItem.TipIndex = TipIndex;
                _editingHistoryItem.Transportator = Transportator;
                _editingHistoryItem.TransportatorTarif = transportatorTarifDecimal;
                _editingHistoryItem.TransportatorMonedaIndex = TransportatorMonedaIndex;
                _editingHistoryItem.TransportatorTipIndex = TransportatorTipIndex;
                _editingHistoryItem.DataIncarcare = EnsureUtcDate(DataIncarcare);
                _editingHistoryItem.DataDescarcare = EnsureUtcDate(DataDescarcare);
                _editingHistoryItem.Produs = Produs;
                _editingHistoryItem.Cantitate = Cantitate?.Trim() ?? string.Empty;
                _editingHistoryItem.TipAdrIndex = TipAdrIndex;
                _editingHistoryItem.Clasa = Clasa;
                _editingHistoryItem.Un = Un;
                _editingHistoryItem.NumarInmatriculare = NumarInmatriculare;
                _editingHistoryItem.LocatieIncarcareAddress = LocatieIncarcareAddress;
                _editingHistoryItem.LocatieIncarcareName = LocatieIncarcareName;
                _editingHistoryItem.LocatieIncarcareCity = LocatieIncarcareCity;
                _editingHistoryItem.LocatieIncarcareCountryCode = LocatieIncarcareCountryCode;
                _editingHistoryItem.LocatieIncarcarePostalCode = LocatieIncarcarePostalCode;
                _editingHistoryItem.LocatieIncarcareCounty = LocatieIncarcareCounty;
                _editingHistoryItem.LocatieDescarcareAddress = LocatieDescarcareAddress;
                _editingHistoryItem.LocatieDescarcareName = LocatieDescarcareName;
                _editingHistoryItem.LocatieDescarcareCity = LocatieDescarcareCity;
                _editingHistoryItem.LocatieDescarcareCountryCode = LocatieDescarcareCountryCode;
                _editingHistoryItem.LocatieDescarcarePostalCode = LocatieDescarcarePostalCode;
                _editingHistoryItem.LocatieDescarcareCounty = LocatieDescarcareCounty;
                _editingHistoryItem.TermenPlata = termenPlataValue;
                _editingHistoryItem.CommentUser = CommentUser;

                // Update in database
                var updateResponse = await _historyRepository.UpdateHistory(_editingHistoryItem);

                if (updateResponse == null)
                {
                    throw new Exception("Failed to update history record");
                }

                Debug.WriteLine($"✅ Order #{NumarComanda} updated successfully");

                MessageBox.Show(
                    $"✅ Order #{NumarComanda} updated successfully!",
                    "Success",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                // Navigate back to history page
                var mainWindow = Application.Current.MainWindow as MainWindow;
                mainWindow?.NavigateToHistory();
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Update error: {ex.Message}");
            MessageBox.Show(
                $"❌ Error updating order:\n\n{ex.Message}",
                "Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            IsSendingEmail = false;
        }
    }

    private bool CanUpdateOrder()
    {
        return !string.IsNullOrWhiteSpace(NumarComanda) && !IsSendingEmail && IsEditMode;
    }

    #endregion

    #region Send Email Command

    [RelayCommand(CanExecute = nameof(CanSendEmail))]
    private async Task SendEmail()
    {
        try
        {
            IsSendingEmail = true;

            Debug.WriteLine("📧 Preparing to send email...");

            // Option arrays used to resolve selected indices
            var monedaOptions = new[] { "EUR", "EUR/MT", "RON" };
            var tipOptions = new[] { "TVA", "ALL IN" };
            var tipAdrOptions = new[] { "ADR", "NON-ADR" };
            var clientTarifValue = Tarif + " " + monedaOptions.ElementAtOrDefault(MonedaIndex);
            var transportatorTarifValue = TransportatorTarif + " " + monedaOptions.ElementAtOrDefault(TransportatorMonedaIndex);

            //recheck from api last number again
            await ReloadLastOrder();

            if(!string.IsNullOrEmpty(Contact))
            {
                Contact contactToInsert = new Contact
                {
                    Name = Contact
                };
                await _contactRepository.InsertAsync(contactToInsert);

            }
            
            if(!string.IsNullOrEmpty(Produs) )
            {
                
                Product productToInsert = new Product
                {
                    Name = Produs
                };
                await _productRepository.InsertAsync(productToInsert);
            }
            
            

            // Parse tarif values as decimals
            decimal? tarifValue = decimal.TryParse(Tarif, out var t) ? t : (decimal?)null;
            decimal? transportatorTarifDecimal = decimal.TryParse(TransportatorTarif, out var tt) ? tt : (decimal?)null;
            decimal? cantitateValue = decimal.TryParse(Cantitate, out var c) ? c : (decimal?)null;
            int? termenPlataValue = int.TryParse(TermenPlata, out var tp) ? tp : (int?)null;

            var responseHistorySave = await _historyRepository.InsertHistory(new HistoryTransport
            {
                // Comanda / Client
                NumarComanda = NumarComanda,
                NumarClient = NumarClient,
                Client = Client,
                Contact = Contact,



                // Tarif client
                Tarif = tarifValue,
                MonedaIndex = MonedaIndex,
                TipIndex = TipIndex,

                // Transportator
                Transportator = Transportator,
                TransportatorTarif = transportatorTarifDecimal,
                TransportatorMonedaIndex = TransportatorMonedaIndex,
                TransportatorTipIndex = TransportatorTipIndex,

                // Date
                DataIncarcare = EnsureUtcDate(DataIncarcare),
                DataDescarcare = EnsureUtcDate(DataDescarcare),

                // Marfa
                Produs = Produs,
                Cantitate = Cantitate,
                TipAdrIndex = TipAdrIndex,
                Clasa = Clasa,
                Un = Un,
                NumarInmatriculare = NumarInmatriculare,

                // Locatie incarcare
                LocatieIncarcareAddress = LocatieIncarcareAddress,
                LocatieIncarcareName = LocatieIncarcareName,
                LocatieIncarcareCity = LocatieIncarcareCity,
                LocatieIncarcareCountryCode = LocatieIncarcareCountryCode,
                LocatieIncarcarePostalCode = LocatieIncarcarePostalCode,
                LocatieIncarcareCounty = LocatieIncarcareCounty,

                // Locatie descarcare
                LocatieDescarcareAddress = LocatieDescarcareAddress,
                LocatieDescarcareName = LocatieDescarcareName,
                LocatieDescarcareCity = LocatieDescarcareCity,
                LocatieDescarcareCountryCode = LocatieDescarcareCountryCode,
                LocatieDescarcarePostalCode = LocatieDescarcarePostalCode,
                LocatieDescarcareCounty = LocatieDescarcareCounty,

                // Alte informatii
                TermenPlata = termenPlataValue,
                CommentUser = CommentUser
            });

            if (responseHistorySave == null)
            {
                Debug.WriteLine("❌ Failed to save history record");
                throw new Exception("Failed to save history record before sending email.");
            }

           

            // Generate comanda.docx (saved to disk permanently - NOT attached to email)
            var comandaPath = await _documentCompletion.GenerateAndSendDocumentAsync(
                NumarComanda,
                NumarClient,
                Client,
                Contact,
                Tarif,
                MonedaIndex,
                TipIndex,
                Transportator,
                TransportatorTarif,
                TransportatorMonedaIndex,
                TransportatorTipIndex,
                DataIncarcare,
                DataDescarcare,
                Produs,
                Cantitate,
                TipAdrIndex,
                Clasa,
                Un,
                NumarInmatriculare,
                // Pickup location components
                LocatieIncarcareAddress,
                LocatieIncarcareName,
                LocatieIncarcareCity,
                LocatieIncarcareCountryCode,
                LocatieIncarcarePostalCode,
                LocatieIncarcareCounty,
                // Delivery location components
                LocatieDescarcareAddress,
                LocatieDescarcareName,
                LocatieDescarcareCity,
                LocatieDescarcareCountryCode,
                LocatieDescarcarePostalCode,
                LocatieDescarcareCounty,
                TermenPlata,
                CommentUser,
                monedaOptions,
                tipOptions,
                tipAdrOptions
            );

            if (comandaPath == null)
            {
                Debug.WriteLine("❌ Failed to generate comanda.docx");
                throw new Exception("Failed to generate comanda.docx");
            }

            Debug.WriteLine($"✅ Saved comanda.docx to: {comandaPath}");

            int tipClientIndex = TipIndex;
            int tipTransportatorIndex = TransportatorTipIndex;
            string tvaClient = tipClientIndex == 0 ? "+ " + tipOptions.ElementAtOrDefault(tipClientIndex) : tipOptions.ElementAtOrDefault(tipClientIndex) ?? string.Empty;
            string tvaTransportator = tipTransportatorIndex == 0 ? "+ " + tipOptions.ElementAtOrDefault(tipTransportatorIndex) : tipOptions.ElementAtOrDefault(tipTransportatorIndex) ?? string.Empty;
            // Build replacements for page2.doc
            var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "DataAzi", DateTime.Today.ToString("dd.MM.yyyy") },
                { "NumarComanda", NumarComanda?.Trim() ?? string.Empty },
                { "NumarClient", NumarClient?.Trim() ?? string.Empty },
                { "ClientNume", Client?.Trim() ?? string.Empty },
                { "ContactPers", Contact?.Trim() ?? string.Empty },
                { "ClientTarif", Tarif?.Trim() ?? string.Empty },
                { "ClientMoneda", monedaOptions.ElementAtOrDefault(MonedaIndex) ?? string.Empty },
                { "ClientTip", tvaClient },
                { "TransportatorNume", Transportator?.Trim() ?? string.Empty },
                { "TransportatorTarif", TransportatorTarif?.Trim() ?? string.Empty },
                { "TransportatorMoneda", monedaOptions.ElementAtOrDefault(TransportatorMonedaIndex) ?? string.Empty },
                { "TransportatorTip", tvaTransportator },
                { "DataIncarcare", DataIncarcare?.ToString("dd/MM/yyyy") ?? string.Empty },
                { "DataDescarcare", DataDescarcare?.ToString("dd/MM/yyyy") ?? string.Empty },
                { "Produs", Produs?.Trim() ?? string.Empty },
                { "CantitateComanda", Cantitate?.Trim() ?? string.Empty },
                { "TipADR", tipAdrOptions.ElementAtOrDefault(TipAdrIndex) ?? string.Empty },
                { "Clasa", Clasa?.Trim() ?? string.Empty },
                { "UserUnInput", Un?.Trim() ?? string.Empty },
                { "NumarInmatriculare", NumarInmatriculare?.Trim().ToUpper() ?? string.Empty },
                { "LocatieIncarcareAddress", LocatieIncarcareAddress?.Trim() ?? string.Empty },
                { "LocatieIncarcareName", LocatieIncarcareName?.Trim() ?? string.Empty },
                { "LocatieIncarcareCity", LocatieIncarcareCity?.Trim() ?? string.Empty },
                { "LocatieIncarcareCountryCode", LocatieIncarcareCountryCode?.Trim() ?? string.Empty },
                { "LocatieIncarcarePostalCode", LocatieIncarcarePostalCode?.Trim() ?? string.Empty },
                { "LocatieIncarcareCounty", LocatieIncarcareCounty?.Trim() ?? string.Empty },
                { "LocatieDescarcareAddress", LocatieDescarcareAddress?.Trim() ?? string.Empty },
                { "LocatieDescarcareName", LocatieDescarcareName?.Trim() ?? string.Empty },
                { "LocatieDescarcareCity", LocatieDescarcareCity?.Trim() ?? string.Empty },
                { "LocatieDescarcareCountryCode", LocatieDescarcareCountryCode?.Trim() ?? string.Empty },
                { "LocatieDescarcarePostalCode", LocatieDescarcarePostalCode?.Trim() ?? string.Empty },
                { "LocatieDescarcareCounty", LocatieDescarcareCounty?.Trim() ?? string.Empty },
                { "TermenPlata", TermenPlata?.Trim() ?? string.Empty },
                { "Comments", CommentUser?.Trim() ?? string.Empty }
            };

            // Find page2.docx template path
            string projectRoot = AppDomain.CurrentDomain.BaseDirectory;
            string projectDir = FindProjectDirectory(projectRoot);
            string docDir = System.IO.Path.Combine(projectDir, "doc");
            string page2TemplatePath = System.IO.Path.Combine(docDir, "page2.docx");

            if (System.IO.File.Exists(page2TemplatePath))
            {
                // Generate page2.doc as temp file
                var page2Path = await _documentCompletion.GenerateAndSendPage2DocAsync(page2TemplatePath, replacements, NumarComanda);
                if (page2Path != null)
                {
                    // Open Thunderbird with ONLY page2.doc attached, then cleanup
                    _documentCompletion.OpenThunderbirdAndCleanup(new[] { page2Path }, new[] { page2Path });
                }
                else
                {
                    Debug.WriteLine("❌ Failed to generate page2.doc");
                }
            }
            else
            {
                Debug.WriteLine($"⚠️ page2.docx template not found at: {page2TemplatePath}");
            }


            Debug.WriteLine("📧 Email sent successfully");

            Debug.WriteLine("🔄 Refreshing data from Supabase...");
            await _searchService.RefreshDataAsync();

            await ReloadAllSuggestionsAsync();

            ResetinputFields();
            ToggleVisibility();
            await InitLastOrder();


            Debug.WriteLine("✅ Email sent and data refreshed");

            MessageBox.Show(
                "✅ Email sent successfully!\n\n" +
                "Data refreshed from Supabase.\n" +
                "All users will see latest data.",
                "Success",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Send email error: {ex.Message}");
            MessageBox.Show(
                $"❌ Error:\n\n{ex.Message}",
                "Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            IsSendingEmail = false;
        }
    }

    private bool CanSendEmail()
    {
        return !string.IsNullOrWhiteSpace(NumarComanda) && !IsSendingEmail && !IsEditMode;
    }

    private static string FindProjectDirectory(string startPath)
    {
        const int MaxParentDirectoryLevels = 6;
        const string DocumentFolderName = "doc";

        string? current = startPath;
        for (int i = 0; i < MaxParentDirectoryLevels && current != null; i++)
        {
            string candidate = System.IO.Path.Combine(current, DocumentFolderName);
            if (System.IO.Directory.Exists(candidate))
            {
                return current;
            }
            current = System.IO.Directory.GetParent(current)?.FullName;
        }
        return startPath;
    }

    #endregion

    #region Reload Data

    private async Task ReloadAllSuggestionsAsync()
    {
        try
        {
            Debug.WriteLine("🔄 Reloading UI suggestions...");

            await Task.WhenAll(
                LoadInitialClientsAsync(),
                LoadInitialTransportatorsAsync(),
                LoadInitialLocationsAsync(IncarcareSuggestions),
                LoadInitialLocationsAsync(DescarcareSuggestions)
            );

            Debug.WriteLine("✅ UI suggestions reloaded");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Error reloading: {ex.Message}");
        }
    }

    private void ResetinputFields()
    {
        NumarComanda = string.Empty;
        NumarClient = string.Empty;
        Client = string.Empty;
        Contact = string.Empty;
        Tarif = string.Empty;
        MonedaIndex = 0;
        TipIndex = 0;
        Transportator = string.Empty;
        TransportatorTarif = string.Empty;
        TransportatorMonedaIndex = 0;
        TransportatorTipIndex = 0;
        DataIncarcare = DateTime.UtcNow.Date;
        DataDescarcare = DateTime.UtcNow.Date.AddDays(1);
        Produs = string.Empty;
        Cantitate = string.Empty;
        TipAdrIndex = 0;
        Clasa = string.Empty;
        Un = string.Empty;
        NumarInmatriculare = string.Empty;
        LocatieIncarcare = string.Empty;
        LocatieDescarcare = string.Empty;
        TermenPlata = string.Empty;
        CommentUser = string.Empty;
        LocatieIncarcareAddress = string.Empty;
        LocatieIncarcareName = string.Empty;
        LocatieIncarcareCity = string.Empty;
        LocatieDescarcareAddress = string.Empty;
        LocatieDescarcareName = string.Empty;
        LocatieDescarcareCity = string.Empty;
    }

    #endregion

    #region Selection Helper Methods (added for UI callbacks)

    public void SetUpdatingFromSelection(bool value)
    {
        _isUpdatingFromSelection = value;
    }

    public void UpdatePickupLocation(Location location)
    {
        if (location == null) return;
        // Show full address in the input after selecting from dropdown
        LocatieIncarcare = location.ToString();

        LocatieIncarcareName = location.Name ?? string.Empty;
        LocatieIncarcareCity = location.City ?? string.Empty;
        LocatieIncarcareAddress = location.Address ?? string.Empty;
        LocatieIncarcareCountryCode = location.CountryCode ?? string.Empty;
        LocatieIncarcarePostalCode = location.PostalCode ?? string.Empty;
        LocatieIncarcareCounty = location.County ?? string.Empty;
    }

    public void UpdateDeliveryLocation(Location location)
    {
        if (location == null) return;
        LocatieDescarcare = location.ToString();

        LocatieDescarcareName = location.Name ?? string.Empty;
        LocatieDescarcareCity = location.City ?? string.Empty;
        LocatieDescarcareAddress = location.Address ?? string.Empty;
        LocatieDescarcareCountryCode = location.CountryCode ?? string.Empty;
        LocatieDescarcarePostalCode = location.PostalCode ?? string.Empty;
        LocatieDescarcareCounty = location.County ?? string.Empty;
    }

    public void GetTermenPlata(Transportator transportator)
    {
        if (transportator == null) return;
        TermenPlata = transportator.TermenulDePlata ?? string.Empty;
    }

    #endregion

    #region Helpers

    private static DateTime EnsureUtcDate(DateTime? value)
    {
        if (!value.HasValue)
            return DateTime.UtcNow.Date;

        var v = value.Value;
        if (v.Kind == DateTimeKind.Utc)
            return v;

        // Treat UI-picked date as a date-only in UTC.
        // We only care about the date (00:00), so avoid time zone shifts.
        return DateTime.SpecifyKind(v, DateTimeKind.Utc);
    }
}

    #endregion
