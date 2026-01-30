using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using EmailCompleteApp.Models;
using EmailCompleteApp.Services;
using EmailCompleteApp.Services.Repositories;

namespace EmailCompleteApp.ViewModels;

public partial class ClientsViewModel : ObservableObject
{
    private readonly ClientRepository _clientRepo = ClientRepository.Instance;
    private readonly TransportatorRepository _transportatorRepo = TransportatorRepository.Instance;
    private readonly SearchService _searchService = SearchService.Instance;

    [ObservableProperty] private int _selectedTabIndex = 0;
    [ObservableProperty] private bool _isAddingNew = false;

    // Track if we're editing existing records
    [ObservableProperty] private int? _editingClientId = null;
    [ObservableProperty] private int? _editingTransportatorId = null;

    // Search filters
    [ObservableProperty] private string _clientSearchText = string.Empty;
    [ObservableProperty] private string _transportatorSearchText = string.Empty;

    #region Client Properties
    private ObservableCollection<Client> _allClients = new();
    [ObservableProperty] private ObservableCollection<Client> _clients = new ();

    [ObservableProperty] private string _clientName = string.Empty;
    [ObservableProperty] private string _clientAddress = string.Empty;
    [ObservableProperty] private string _clientBank = string.Empty;
    [ObservableProperty] private string _clientIban = string.Empty;
    [ObservableProperty] private string _clientVatNumber = string.Empty;
    [ObservableProperty] private string _clientCameraDeComert = string.Empty;
    [ObservableProperty] private string _clientTermenPlata = string.Empty;

    #endregion

    #region Transportator Properties

    private ObservableCollection<Transportator> _allTransportators = new();
    [ObservableProperty] private ObservableCollection<Transportator> _transportators = new();

    [ObservableProperty] private string _transportatorName = string.Empty;
    [ObservableProperty] private string _transportatorAdresa = string.Empty;
    [ObservableProperty] private string _transportatorContBancar = string.Empty;
    [ObservableProperty] private string _transportatorIban = string.Empty;
    [ObservableProperty] private string _transportatorVatNumber = string.Empty;
    [ObservableProperty] private string _transportatorCameraDeComert = string.Empty;
    [ObservableProperty] private string _transportatorTermenPlata = string.Empty;

    #endregion

    public ClientsViewModel()
    {
        _ = Init();
    }

    public async Task Init()
    {
        await LoadAllClients();
        await LoadAllTransportators();
    }

    partial void OnClientSearchTextChanged(string value)
    {
        FilterClients();
    }

    partial void OnTransportatorSearchTextChanged(string value)
    {
        FilterTransportators();
    }

    private void FilterClients()
    {
        if (string.IsNullOrWhiteSpace(ClientSearchText))
        {
            Clients = new ObservableCollection<Client>(_allClients);
        }
        else
        {
            var filtered = _allClients
                .Where(c => c.Name.Contains(ClientSearchText, StringComparison.OrdinalIgnoreCase))
                .ToList();
            Clients = new ObservableCollection<Client>(filtered);
        }
    }

    private void FilterTransportators()
    {
        if (string.IsNullOrWhiteSpace(TransportatorSearchText))
        {
            Transportators = new ObservableCollection<Transportator>(_allTransportators);
        }
        else
        {
            var filtered = _allTransportators
                .Where(t => t.Name.Contains(TransportatorSearchText, StringComparison.OrdinalIgnoreCase))
                .ToList();
            Transportators = new ObservableCollection<Transportator>(filtered);
        }
    }

    [RelayCommand]
    private void AddNew()
    {
        IsAddingNew = true;
        EditingClientId = null;
        EditingTransportatorId = null;
        ClearClientFields();
        ClearTransportatorFields();
    }

    [RelayCommand]
    private void Cancel()
    {
        IsAddingNew = false;
        EditingClientId = null;
        EditingTransportatorId = null;
        ClearClientFields();
        ClearTransportatorFields();
    }

    /// <summary>
    /// Opens the form with pre-filled data for editing a client
    /// </summary>
    [RelayCommand]
    public void EditClient(Client client)
    {
        if (client == null) return;

        EditingClientId = client.Id;
        EditingTransportatorId = null;
        
        ClientName = client.Name;
        ClientAddress = client.Address;
        ClientBank = client.Bank;
        ClientIban = client.IBAN;
        ClientVatNumber = client.VATNumber;
        ClientCameraDeComert = client.CameraDeComert;
        ClientTermenPlata = client.TermenulDePlata;

        SelectedTabIndex = 0; // Switch to Client tab
        IsAddingNew = true; // Show the form
    }

    /// <summary>
    /// Opens the form with pre-filled data for editing a transportator
    /// </summary>
    [RelayCommand]
    public void EditTransportator(Transportator transportator)
    {
        if (transportator == null) return;

        EditingTransportatorId = transportator.Id;
        EditingClientId = null;
        
        TransportatorName = transportator.Name;
        TransportatorAdresa = transportator.Address;
        TransportatorContBancar = transportator.Bank;
        TransportatorIban = transportator.IBAN;
        TransportatorVatNumber = transportator.VATNumber;
        TransportatorCameraDeComert = transportator.CameraDeComert;
        TransportatorTermenPlata = transportator.TermenulDePlata;

        SelectedTabIndex = 1; // Switch to Transportator tab
        IsAddingNew = true; // Show the form
    }

    [RelayCommand]
    private async Task Submit()
    {
        try
        {
            if (SelectedTabIndex == 0)
            {
                // ===== CLIENT - INSERT OR UPDATE =====
                if (string.IsNullOrWhiteSpace(ClientName))
                {
                    ShowWarn("Numele clientului este obligatoriu.");
                    return;
                }
                if (string.IsNullOrWhiteSpace(ClientAddress))
                {
                    ShowWarn("Adresa clientului este obligatorie.");
                    return;
                }

                var client = new Client
                {
                    Name = ClientName.Trim(),
                    Address = ClientAddress.Trim(),
                    Bank = ClientBank?.Trim() ?? string.Empty,
                    IBAN = ClientIban?.Trim() ?? string.Empty,
                    VATNumber = ClientVatNumber?.Trim() ?? string.Empty,
                    CameraDeComert = ClientCameraDeComert?.Trim() ?? string.Empty,
                    TermenulDePlata = ClientTermenPlata?.Trim() ?? string.Empty
                };

                Client savedClient;
                if (EditingClientId.HasValue)
                {
                    // UPDATE existing client
                    client.Id = EditingClientId.Value;
                    savedClient = await _clientRepo.UpdateAsync(client);

                    MessageBox.Show(
                        $"✅ Client actualizat cu succes în Supabase!\n\n" +
                        $"Nume: {savedClient.Name}\n" +
                        $"ID: {savedClient.Id}\n\n" +
                        $"Datele au fost actualizate și sunt disponibile pentru toți utilizatorii!",
                        "Succes",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                else
                {
                    // INSERT new client
                    savedClient = await _clientRepo.InsertAsync(client);

                    MessageBox.Show(
                        $"✅ Client salvat cu succes în Supabase!\n\n" +
                        $"Nume: {savedClient.Name}\n" +
                        $"ID: {savedClient.Id}\n" +
                        $"Creat: {savedClient.CreatedAt:yyyy-MM-dd HH:mm:ss} UTC\n\n" +
                        $"Datele sunt disponibile pentru toți utilizatorii!",
                        "Succes",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }

                // Refresh cached data
                await _searchService.RefreshDataAsync();

                // Reload the list
                await LoadAllClients();

                // Clear form and return to list view
                ClearClientFields();
                EditingClientId = null;
                IsAddingNew = false;
            }
            else
            {
                // ===== TRANSPORTATOR - INSERT OR UPDATE =====
                if (string.IsNullOrWhiteSpace(TransportatorName))
                {
                    ShowWarn("Numele transportatorului este obligatoriu.");
                    return;
                }
                if (string.IsNullOrWhiteSpace(TransportatorAdresa))
                {
                    ShowWarn("Adresa transportatorului este obligatorie.");
                    return;
                }

                var transportator = new Transportator
                {
                    Name = TransportatorName.Trim(),
                    Address = TransportatorAdresa.Trim(),
                    Bank = TransportatorContBancar?.Trim() ?? string.Empty,
                    IBAN = TransportatorIban?.Trim() ?? string.Empty,
                    VATNumber = TransportatorVatNumber?.Trim() ?? string.Empty,
                    CameraDeComert = TransportatorCameraDeComert?.Trim() ?? string.Empty,
                    TermenulDePlata = TransportatorTermenPlata?.Trim() ?? string.Empty
                };

                Transportator savedTransportator;
                if (EditingTransportatorId.HasValue)
                {
                    // UPDATE existing transportator
                    transportator.Id = EditingTransportatorId.Value;
                    savedTransportator = await _transportatorRepo.UpdateAsync(transportator);

                    MessageBox.Show(
                        $"✅ Transportator actualizat cu succes în Supabase!\n\n" +
                        $"Nume: {savedTransportator.Name}\n" +
                        $"ID: {savedTransportator.Id}\n\n" +
                        $"Datele au fost actualizate și sunt disponibile pentru toți utilizatorii!",
                        "Succes",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                else
                {
                    // INSERT new transportator
                    savedTransportator = await _transportatorRepo.InsertAsync(transportator);

                    MessageBox.Show(
                        $"✅ Transportator salvat cu succes în Supabase!\n\n" +
                        $"Nume: {savedTransportator.Name}\n" +
                        $"ID: {savedTransportator.Id}\n" +
                        $"Creat: {savedTransportator.CreatedAt:yyyy-MM-dd HH:mm:ss} UTC\n\n" +
                        $"Datele sunt disponibile pentru toți utilizatorii!",
                        "Succes",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }

                // Refresh cached data
                await _searchService.RefreshDataAsync();

                // Reload the list
                await LoadAllTransportators();

                // Clear form and return to list view
                ClearTransportatorFields();
                EditingTransportatorId = null;
                IsAddingNew = false;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"❌ Eroare la salvarea în Supabase:\n\n" +
                $"{ex.Message}\n\n" +
                $"Excepție internă:\n{ex.InnerException?.Message}",
                "Eroare de bază de date",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private async void Insert()
    {
        await Submit();
    }

    private async Task<bool> LoadAllClients()
    {
        var clients = await _clientRepo.LoadAllAsync();
        if (clients != null)
        {
            _allClients = new ObservableCollection<Client>(clients);
            FilterClients(); // Apply current filter
            return true;
        }
        else
        {
            MessageBox.Show("Încărcarea clienților a eșuat.", "Eroare", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
    }

    private async Task<bool> LoadAllTransportators()
    {
        var transportators = await _transportatorRepo.LoadAllAsync();
        if (transportators != null)
        {
            _allTransportators = new ObservableCollection<Transportator>(transportators);
            FilterTransportators(); // Apply current filter
            return true;
        }
        else
        {
            MessageBox.Show("Încărcarea transportatorilor a eșuat.", "Eroare", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
    }

    [RelayCommand]
    private async Task DeleteClient(Client client) 
    {
        if (client == null) return;
        var result = MessageBox.Show(
            $"Sunteți sigur că doriți să ștergeți clientul '{client.Name}' (ID: {client.Id})?",
            "Confirmare ștergere",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);
        if (result == MessageBoxResult.Yes)
        {
            try
            {
                await _clientRepo.DeleteAsync(client.Id);
                await _searchService.RefreshDataAsync();
                await LoadAllClients();
                MessageBox.Show(
                    $"✅ Clientul '{client.Name}' a fost șters cu succes!",
                    "Succes",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"❌ Eroare la ștergerea clientului:\n\n{ex.Message}",
                    "Eroare de bază de date",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }

    [RelayCommand]
    private async Task DeleteTransportator(Transportator transportator) 
    {
        if (transportator == null) return;
        var result = MessageBox.Show(
            $"Sunteți sigur că doriți să ștergeți transportatorul '{transportator.Name}' (ID: {transportator.Id})?",
            "Confirmare ștergere",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);
        if (result == MessageBoxResult.Yes)
        {
            try
            {
                await _transportatorRepo.DeleteAsync(transportator.Id);
                await _searchService.RefreshDataAsync();
                await LoadAllTransportators();
                MessageBox.Show(
                    $"✅ Transportatorul '{transportator.Name}' a fost șters cu succes!",
                    "Succes",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"❌ Eroare la ștergerea transportatorului:\n\n{ex.Message}",
                    "Eroare de bază de date",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }

    private void ClearClientFields()
    {
        ClientName = string.Empty;
        ClientAddress = string.Empty;
        ClientBank = string.Empty;
        ClientIban = string.Empty;
        ClientVatNumber = string.Empty;
        ClientCameraDeComert = string.Empty;
        ClientTermenPlata = string.Empty;
    }

    private void ClearTransportatorFields()
    {
        TransportatorName = string.Empty;
        TransportatorAdresa = string.Empty;
        TransportatorContBancar = string.Empty;
        TransportatorIban = string.Empty;
        TransportatorVatNumber = string.Empty;
        TransportatorCameraDeComert = string.Empty;
        TransportatorTermenPlata = string.Empty;
    }

    private static void ShowWarn(string msg) =>
        MessageBox.Show(msg, "Eroare de validare", MessageBoxButton.OK, MessageBoxImage.Warning);
}