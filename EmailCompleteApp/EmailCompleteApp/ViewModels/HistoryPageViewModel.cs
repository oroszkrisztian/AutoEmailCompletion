using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using EmailCompleteApp.Models;
using EmailCompleteApp.Services;
using EmailCompleteApp.Services.Repositories;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace EmailCompleteApp.ViewModels
{
    public partial class HistoryPageViewModel: ObservableObject
    {
        private readonly HistoryRepository _historyRepository;
        
        private ObservableCollection<HistoryTransport> _allHistoryData = new();
        public ObservableCollection<HistoryTransport> HistoryData { get; } = new ();
        
        [ObservableProperty]
        private bool isLoading;

        [ObservableProperty]
        private string orderNumberSearchText = string.Empty;

        [ObservableProperty]
        private string clientNameSearchText = string.Empty;

        // Event to notify when edit is requested
        public event Action<HistoryTransport>? EditRequested;

        public HistoryPageViewModel()
        {
            _historyRepository = HistoryRepository.Instance;
            _ = InitializeHistoryData();
        }

        partial void OnOrderNumberSearchTextChanged(string value)
        {
            FilterHistory();
        }

        partial void OnClientNameSearchTextChanged(string value)
        {
            FilterHistory();
        }

        private void FilterHistory()
        {
            var filtered = _allHistoryData.AsEnumerable();

            if (!string.IsNullOrWhiteSpace(OrderNumberSearchText))
            {
                filtered = filtered.Where(h =>
                    h.NumarComanda.Contains(OrderNumberSearchText, StringComparison.OrdinalIgnoreCase));
            }

            if (!string.IsNullOrWhiteSpace(ClientNameSearchText))
            {
                filtered = filtered.Where(h =>
                    (h.Client ?? string.Empty).Contains(ClientNameSearchText, StringComparison.OrdinalIgnoreCase));
            }

            HistoryData.Clear();
            foreach (var item in filtered)
            {
                HistoryData.Add(item);
            }
        }

        private async Task InitializeHistoryData()
        {
            try
            {
                IsLoading = true; 
                Debug.WriteLine("Initializing history data...");
                var response = await _historyRepository.LoadAllByOrderNumDescAsync();
                if (response != null)
                {
                    _allHistoryData.Clear();
                    foreach (HistoryTransport item in response)
                    {
                        _allHistoryData.Add(item);
                        await PrintLoadedData(item);
                    }
                    FilterHistory(); // Apply initial filter
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ Failed to initialize history data: {ex.Message}");
                MessageBox.Show(
                    $"Eroare la încărcarea istoricului:\n\n{ex.Message}",
                    "Eroare",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }

        [RelayCommand]
        public void EditOrder(HistoryTransport historyItem)
        {
            if (historyItem != null)
            {
                Debug.WriteLine($"Edit requested for order: {historyItem.NumarComanda}");
                EditRequested?.Invoke(historyItem);
            }
        }

        [RelayCommand]
        public void OpenDocument(string orderNumber)
        {
            bool fileNotFound = false;
            string errorMessage = string.Empty;
            try
            {
                if (string.IsNullOrWhiteSpace(orderNumber))
                {
                    fileNotFound = true;
                    errorMessage = "Order number is null or empty.";
                    return;
                }

                string projectRoot = AppDomain.CurrentDomain.BaseDirectory;
                string projectDir = FindProjectDirectory(projectRoot);
                string generatedDir = Path.Combine(projectDir, "doc", "Generated");

                // Primary expected filename
                string expectedFileName = $"CAPAC+Comanda transport - {orderNumber}.docx";
                string expectedPath = Path.Combine(generatedDir, expectedFileName);

                string fileToOpen = expectedPath;

                if (!File.Exists(fileToOpen))
                {
                    // If exact name doesn't match, try a best-effort search for files containing the order number.
                    if (Directory.Exists(generatedDir))
                    {
                        var matches = Directory.GetFiles(generatedDir, $"*{orderNumber}*.docx", SearchOption.TopDirectoryOnly);
                        if (matches != null && matches.Length > 0)
                        {
                            fileToOpen = matches[0];
                        }
                        else
                        {
                            fileNotFound = true;
                            errorMessage = $"Order document not found: {expectedFileName} in {generatedDir}";
                            return;
                        }
                    }
                    else
                    {
                        fileNotFound = true;
                        errorMessage = $"Generated documents folder not found: {generatedDir}";
                        return;
                    }
                }

                // Open with default associated application
                Task.Run(() => Process.Start(new ProcessStartInfo(fileToOpen) { UseShellExecute = true }));
            }
            catch (Exception ex)
            {
                fileNotFound = true;
                errorMessage = $"Order document not found: {ex.Message}";
            }
            finally
            {
                if (fileNotFound)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show(
                            string.IsNullOrEmpty(errorMessage) ? "File not found. Please check the order number or verify the document's availability." : errorMessage,
                            "Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    });
                }
            }
        }

        private static string FindProjectDirectory(string startPath)
        {
            const int MaxParentDirectoryLevels = 6;
            const string DocumentFolderName = "doc";

            string? current = startPath;
            for (int i = 0; i < MaxParentDirectoryLevels && current != null; i++)
            {
                string candidate = Path.Combine(current, DocumentFolderName);
                if (Directory.Exists(candidate))
                {
                    return current;
                }
                current = Directory.GetParent(current)?.FullName;
            }
            // If not found, fallback to startPath so callers still have a path to attempt
            return startPath;
        }

        public static Task PrintLoadedData(HistoryTransport item)
        {
            var route = $"{item.LocatieIncarcareCity ?? "?"} - {item.LocatieDescarcareCity ?? "?"}";
            var clientTarif = item.Tarif.HasValue ? $"{item.Tarif.Value}" : "N/A";
            var transportatorTarif = item.TransportatorTarif.HasValue ? $"{item.TransportatorTarif.Value}" : "N/A";
            
            Debug.WriteLine($"History Item - ID: {item.Id}, Client: {item.Client}, Route: {route}, " +
                          $"Date Loaded: {item.DataIncarcare?.ToString("dd/MM/yyyy") ?? "N/A"}, " +
                          $"Date Unloaded: {item.DataDescarcare?.ToString("dd/MM/yyyy") ?? "N/A"}, " +
                          $"Client Tarif: {clientTarif}, Transportator Tarif: {transportatorTarif}, " +
                          $"Created At: {item.CreatedAt}, Order Number: {item.NumarComanda}");
            return Task.CompletedTask;
        }

    }
}

