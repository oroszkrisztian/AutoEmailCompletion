using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using EmailCompleteApp.Models;
using EmailCompleteApp.Services;
using EmailCompleteApp.Services.Repositories;

namespace EmailCompleteApp.ViewModels;

public partial class LocationViewModel : ObservableObject
{
    private readonly LocationRepository _locationRepo = LocationRepository.Instance;
    private readonly SearchService _searchService = SearchService.Instance;

    [ObservableProperty] private bool _isAddingNew = false;

    // Track if we're editing existing location
    [ObservableProperty] private int? _editingLocationId = null;

    // Search filter
    [ObservableProperty] private string _locationSearchText = string.Empty;

    #region Location Properties
    private ObservableCollection<Location> _allLocations = new();
    [ObservableProperty] private ObservableCollection<Location> _locations = new();

    [ObservableProperty] private string _locationName = string.Empty;
    [ObservableProperty] private string _locationAddress = string.Empty;
    [ObservableProperty] private string _locationCity = string.Empty;
    [ObservableProperty] private string _locationCountryCode = string.Empty;
    [ObservableProperty] private string _locationPostalCode = string.Empty;
    [ObservableProperty] private string _locationCounty = string.Empty;
    #endregion

    public LocationViewModel()
    {
        _ = Init();
    }

    public async Task Init()
    {
        await LoadAllLocations();
    }

    partial void OnLocationSearchTextChanged(string value)
    {
        FilterLocations();
    }

    private void FilterLocations()
    {
        if (string.IsNullOrWhiteSpace(LocationSearchText))
        {
            Locations = new ObservableCollection<Location>(_allLocations);
        }
        else
        {
            var filtered = _allLocations
                .Where(l => 
                    (l.Name?.Contains(LocationSearchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (l.City?.Contains(LocationSearchText, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (l.Address?.Contains(LocationSearchText, StringComparison.OrdinalIgnoreCase) ?? false))
                .ToList();
            Locations = new ObservableCollection<Location>(filtered);
        }
    }

    [RelayCommand]
    private void AddNew()
    {
        IsAddingNew = true;
        EditingLocationId = null;
        ClearLocationFields();
    }

    [RelayCommand]
    private void Cancel()
    {
        IsAddingNew = false;
        EditingLocationId = null;
        ClearLocationFields();
    }

    /// <summary>
    /// Opens the form with pre-filled data for editing a location
    /// </summary>
    public void EditLocation(Location location)
    {
        if (location == null) return;

        EditingLocationId = location.Id;
        
        LocationName = location.Name;
        LocationAddress = location.Address;
        LocationCity = location.City;
        LocationCountryCode = location.CountryCode ?? string.Empty;
        LocationPostalCode = location.PostalCode ?? string.Empty;
        LocationCounty = location.County ?? string.Empty;

        IsAddingNew = true; // Show the form
    }

    [RelayCommand]
    private async Task Submit()
    {
        try
        {
            // Validate required fields
            if (string.IsNullOrWhiteSpace(LocationName))
            {
                ShowWarn("Numele locației este obligatoriu.");
                return;
            }
            if (string.IsNullOrWhiteSpace(LocationAddress))
            {
                ShowWarn("Adresa locației este obligatorie.");
                return;
            }
            if (string.IsNullOrWhiteSpace(LocationCity))
            {
                ShowWarn("Orașul este obligatoriu.");
                return;
            }

            var location = new Location
            {
                Name = LocationName.Trim(),
                Address = LocationAddress.Trim(),
                City = LocationCity.Trim(),
                CountryCode = LocationCountryCode?.Trim() ?? string.Empty,
                PostalCode = LocationPostalCode?.Trim() ?? string.Empty,
                County = LocationCounty?.Trim() ?? string.Empty
            };

            Location savedLocation;
            if (EditingLocationId.HasValue)
            {
                // UPDATE existing location
                location.Id = EditingLocationId.Value;
                savedLocation = await _locationRepo.UpdateAsync(location);

                MessageBox.Show(
                    $"✅ Locație actualizată cu succes în Supabase!\n\n" +
                    $"Nume: {savedLocation.Name}\n" +
                    $"ID: {savedLocation.Id}\n\n" +
                    $"Datele au fost actualizate și sunt disponibile pentru toți utilizatorii!",
                    "Succes",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            else
            {
                // INSERT new location
                savedLocation = await _locationRepo.InsertAsync(location);

                MessageBox.Show(
                    $"✅ Locație salvată cu succes în Supabase!\n\n" +
                    $"Nume: {savedLocation.Name}\n" +
                    $"ID: {savedLocation.Id}\n\n" +
                    $"Datele sunt disponibile pentru toți utilizatorii!",
                    "Succes",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }

            // Refresh cached data
            await _searchService.RefreshDataAsync();

            // Reload the list
            await LoadAllLocations();

            // Clear form and return to list view
            ClearLocationFields();
            EditingLocationId = null;
            IsAddingNew = false;
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
    private async void Save()
    {
        await Submit();
    }

    private async Task<bool> LoadAllLocations()
    {
        var locations = await _locationRepo.LoadAllAsync();
        if (locations != null)
        {
            _allLocations = new ObservableCollection<Location>(locations);
            FilterLocations(); // Apply current filter
            return true;
        }
        else
        {
            MessageBox.Show("Încărcarea locațiilor a eșuat.", "Eroare", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
    }

    private void ClearLocationFields()
    {
        LocationName = string.Empty;
        LocationAddress = string.Empty;
        LocationCity = string.Empty;
        LocationCountryCode = string.Empty;
        LocationPostalCode = string.Empty;
        LocationCounty = string.Empty;
    }

    private static void ShowWarn(string msg) =>
        MessageBox.Show(msg, "Eroare de validare", MessageBoxButton.OK, MessageBoxImage.Warning);
}