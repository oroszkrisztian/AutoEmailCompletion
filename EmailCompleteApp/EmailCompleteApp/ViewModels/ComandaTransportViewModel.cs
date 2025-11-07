using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using EmailCompleteApp.Models;
using EmailCompleteApp.Services;
using EmailCompleteApp.Services.Repositories;

namespace EmailCompleteApp.ViewModels;

public partial class ComandaTransportViewModel : ObservableObject
{
    private readonly SearchService _searchService;
    private readonly DocumentCompletion _documentCompletion;
    private readonly HistoryRepository _historyRepository;

    // Search debounce settings
    private const int SearchDebounceDelayMs = 300;
    private const int InitialSuggestions = 10;

    // Cancellation tokens for search operations
    private CancellationTokenSource? _clientSearchCts;
    private CancellationTokenSource? _transportatorSearchCts;
    private CancellationTokenSource? _incarcareSearchCts;
    private CancellationTokenSource? _descarcareSearchCts;

    // Flag to prevent search when updating from selection
    private bool _isUpdatingFromSelection = false;

    #region Observable Properties

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

    // Additional fields required by DocumentCompletion
    [ObservableProperty] private string _commentUser = string.Empty;

    // Pickup location components
    [ObservableProperty] private string _locatieIncarcareAddress = string.Empty;
    [ObservableProperty] private string _locatieIncarcareName = string.Empty;
    [ObservableProperty] private string _locatieIncarcareCity = string.Empty;
    [ObservableProperty] private string _locatieincarcareCode = string.Empty;

    // Delivery location components
    [ObservableProperty] private string _locatieDescarcareAddress = string.Empty;
    [ObservableProperty] private string _locatieDescarcareName = string.Empty;
    [ObservableProperty] private string _locatieDescarcareCity = string.Empty;
    [ObservableProperty] private string _locatieDescarcareCode = string.Empty;

    #endregion

    #region Collections

    public ObservableCollection<Client> ClientSuggestions { get; } = new();
    public ObservableCollection<Transportator> TransportatorSuggestions { get; } = new();
    public ObservableCollection<Location> IncarcareSuggestions { get; } = new();
    public ObservableCollection<Location> DescarcareSuggestions { get; } = new();

    #endregion

    #region Constructor

    public ComandaTransportViewModel()
    {
        _searchService = SearchService.Instance;
        _documentCompletion = DocumentCompletion.Instance;
        _historyRepository = HistoryRepository.Instance;
        _ = InitializeSuggestionsAsync();
    }

    #endregion

    #region Initialization

    private async Task InitializeSuggestionsAsync()
    {
        try
        {
            await _searchService.LoadAllDataAsync();

            await Task.WhenAll(
                LoadInitialClientsAsync(),
                LoadInitialTransportatorsAsync(),
                LoadInitialLocationsAsync(IncarcareSuggestions),
                LoadInitialLocationsAsync(DescarcareSuggestions)
            );

            Debug.WriteLine("✅ Initial suggestions loaded from Supabase cache");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Error initializing suggestions: {ex.Message}");
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
            UpdateCollection(ClientSuggestions, results.Take(InitialSuggestions).ToList());
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
            UpdateCollection(TransportatorSuggestions, results.Take(InitialSuggestions).ToList());
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
            UpdateCollection(targetCollection, results.Take(InitialSuggestions).ToList());
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"❌ Load locations error: {ex.Message}");
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

            var responseHisotrySave = await _historyRepository.InsertHistory(new HistoryTransport
            {
                ClientName = Client,
                Route = $"{LocatieIncarcareCity} - {LocatieDescarcareCity}",
                DateLoaded = EnsureUtcDate(DataIncarcare),
                DateUnloaded = EnsureUtcDate(DataDescarcare),
                ClientTarif = Tarif,
                TransportatorTarif = TransportatorTarif,
                NumarComanda = int.TryParse(NumarComanda, out var numar) ? numar : 0
            });

            if (responseHisotrySave == null)
            {
                Debug.WriteLine("❌ Failed to save history record");
                throw new Exception("Failed to save history record before sending email.");
            }

            await _documentCompletion.GenerateAndSendDocumentAsync(
                NumarComanda,
                NumarClient,
                Client,
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
                LocatieincarcareCode,
                // Delivery location components
                LocatieDescarcareAddress,
                LocatieDescarcareName,
                LocatieDescarcareCity,
                LocatieDescarcareCode,
                TermenPlata,
                CommentUser,
                monedaOptions,
                tipOptions,
                tipAdrOptions
            );

            Debug.WriteLine("📧 Email sent successfully");

            // 🔥 Refresh data from Supabase (3 queries via repositories)
            Debug.WriteLine("🔄 Refreshing data from Supabase...");
            await _searchService.RefreshDataAsync();

            // 🔄 Reload UI suggestions
            await ReloadAllSuggestionsAsync();

            ResetinputFields();

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
        return !string.IsNullOrWhiteSpace(NumarComanda) && !IsSendingEmail;
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
        LocatieincarcareCode = location.Code ?? string.Empty;
    }

    public void UpdateDeliveryLocation(Location location)
    {
        if (location == null) return;
        LocatieDescarcare = location.ToString();

        LocatieDescarcareName = location.Name ?? string.Empty;
        LocatieDescarcareCity = location.City ?? string.Empty;
        LocatieDescarcareAddress = location.Address ?? string.Empty;
        LocatieDescarcareCode = location.Code ?? string.Empty;
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

    #endregion
}