using System;
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

    #region Client Properties

    [ObservableProperty] private string _clientName = string.Empty;
    [ObservableProperty] private string _clientAddress = string.Empty;
    [ObservableProperty] private string _clientBank = string.Empty;
    [ObservableProperty] private string _clientIban = string.Empty;
    [ObservableProperty] private string _clientVatNumber = string.Empty;
    [ObservableProperty] private string _clientCameraDeComert = string.Empty;
    [ObservableProperty] private string _clientTermenPlata = string.Empty;

    #endregion

    #region Transportator Properties

    [ObservableProperty] private string _transportatorName = string.Empty;
    [ObservableProperty] private string _transportatorAdresa = string.Empty;
    [ObservableProperty] private string _transportatorContBancar = string.Empty;
    [ObservableProperty] private string _transportatorIban = string.Empty;
    [ObservableProperty] private string _transportatorVatNumber = string.Empty;
    [ObservableProperty] private string _transportatorCameraDeComert = string.Empty;
    [ObservableProperty] private string _transportatorTermenPlata = string.Empty;

    #endregion

    [RelayCommand]
    private async void Insert()
    {
        try
        {
            if (SelectedTabIndex == 0)
            {
                // ===== INSERT CLIENT =====
                if (string.IsNullOrWhiteSpace(ClientName))
                {
                    ShowWarn("Client name is required.");
                    return;
                }
                if (string.IsNullOrWhiteSpace(ClientAddress))
                {
                    ShowWarn("Client address is required.");
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

                // 🔥 Insert via ClientRepository (1 query to Supabase)
                var savedClient = await _clientRepo.InsertAsync(client);

                // 🔥 Refresh cached data (3 queries to Supabase)
                await _searchService.RefreshDataAsync();

                MessageBox.Show(
                    $"✅ Client saved successfully to Supabase!\n\n" +
                    $"Name: {savedClient.Name}\n" +
                    $"ID: {savedClient.Id}\n" +
                    $"Created: {savedClient.CreatedAt:yyyy-MM-dd HH:mm:ss} UTC\n\n" +
                    $"Data refreshed and available to all users!",
                    "Success",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                // Clear input fields
                ClientName = string.Empty;
                ClientAddress = string.Empty;
                ClientBank = string.Empty;
                ClientIban = string.Empty;
                ClientVatNumber = string.Empty;
                ClientCameraDeComert = string.Empty;
                ClientTermenPlata = string.Empty;
            }
            else
            {
                // ===== INSERT TRANSPORTATOR =====
                if (string.IsNullOrWhiteSpace(TransportatorName))
                {
                    ShowWarn("Transportator name is required.");
                    return;
                }
                if (string.IsNullOrWhiteSpace(TransportatorAdresa))
                {
                    ShowWarn("Transportator address is required.");
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

                // 🔥 Insert via TransportatorRepository (1 query to Supabase)
                var savedTransportator = await _transportatorRepo.InsertAsync(transportator);

                // 🔥 Refresh cached data (3 queries to Supabase)
                await _searchService.RefreshDataAsync();

                MessageBox.Show(
                    $"✅ Transportator saved successfully to Supabase!\n\n" +
                    $"Name: {savedTransportator.Name}\n" +
                    $"ID: {savedTransportator.Id}\n" +
                    $"Created: {savedTransportator.CreatedAt:yyyy-MM-dd HH:mm:ss} UTC\n\n" +
                    $"Data refreshed and available to all users!",
                    "Success",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                // Clear input fields
                TransportatorName = string.Empty;
                TransportatorAdresa = string.Empty;
                TransportatorContBancar = string.Empty;
                TransportatorIban = string.Empty;
                TransportatorVatNumber = string.Empty;
                TransportatorCameraDeComert = string.Empty;
                TransportatorTermenPlata = string.Empty;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"❌ Error saving to Supabase:\n\n" +
                $"{ex.Message}\n\n" +
                $"Inner Exception:\n{ex.InnerException?.Message}",
                "Database Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private static void ShowWarn(string msg) =>
        MessageBox.Show(msg, "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
}