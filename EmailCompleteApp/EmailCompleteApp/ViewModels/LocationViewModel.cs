using System;
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

    [ObservableProperty] private string _locationName = string.Empty;
    [ObservableProperty] private string _locationAddress = string.Empty;
    [ObservableProperty] private string _locationCity = string.Empty;
    [ObservableProperty] private string _locationCityCode = string.Empty;

    [RelayCommand]
    private async void Save()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(LocationName))
            {
                Warn("Location name is required.");
                return;
            }
            if (string.IsNullOrWhiteSpace(LocationAddress))
            {
                Warn("Location address is required.");
                return;
            }
            if (string.IsNullOrWhiteSpace(LocationCity))
            {
                Warn("Location city is required.");
                return;
            }

            var location = new Location
            {
                Name = LocationName.Trim(),
                Address = LocationAddress.Trim(),
                City = LocationCity.Trim(),
                Code = LocationCityCode.Trim()
            };

            // 🔥 Insert via LocationRepository (1 query to Supabase)
            var savedLocation = await _locationRepo.InsertAsync(location);

            // 🔥 Refresh cached data (3 queries to Supabase)
            await _searchService.RefreshDataAsync();

            MessageBox.Show(
                $"✅ Location saved successfully to Supabase!\n\n" +
                $"Name: {savedLocation.Name}\n" +
                $"City: {savedLocation.City}\n" +
                $"ID: {savedLocation.Id}\n\n" +
                $"Data refreshed and available to all users!",
                "Success",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            // Clear input fields
            LocationName = string.Empty;
            LocationAddress = string.Empty;
            LocationCity = string.Empty;
            LocationCityCode = string.Empty;
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

    private static void Warn(string msg) =>
        MessageBox.Show(msg, "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
}