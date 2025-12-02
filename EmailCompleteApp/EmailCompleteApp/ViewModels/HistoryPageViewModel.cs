using CommunityToolkit.Mvvm.ComponentModel;
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

namespace EmailCompleteApp.ViewModels
{
    public partial class HistoryPageViewModel: ObservableObject
    {
        private readonly HistoryRepository _historyRepository;
        public ObservableCollection<HistoryTransport> HistoryData { get; } = new ();
        
        
        [ObservableProperty]
        private bool isLoading;

        public HistoryPageViewModel()
        {
            _historyRepository = HistoryRepository.Instance;
            _ = InitializeHistoryData();
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
                    HistoryData.Clear();
                    foreach (HistoryTransport item in response)
                    {
                        HistoryData.Add(item);
                        await PrintLoadedData(item);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to initialize history data: ", ex);
            }
            finally
            {
                IsLoading = false;
            }
        }

        public void OpenDocument(string orderNumber)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(orderNumber))
                {
                    throw new FileNotFoundException("Order number is null or empty.");
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
                            throw new FileNotFoundException($"Order document not found: {expectedFileName} in {generatedDir}");
                        }
                    }
                    else
                    {
                        throw new DirectoryNotFoundException($"Generated documents folder not found: {generatedDir}");
                    }
                }

                // Open with default associated application
                Task.Run(() => Process.Start(new ProcessStartInfo(fileToOpen) { UseShellExecute = true }));
            }
            catch(Exception ex)
            {
                // Preserve original behavior of throwing a FileNotFoundException, include inner details.
                throw new FileNotFoundException("Ordr document not found " + ex);
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
            Debug.WriteLine($"History Item - ID: {item.Id}, Client: {item.ClientName}, Route: {item.Route}, Date Loaded: {item.DateLoaded}, Date Unloaded: {item.DateUnloaded}, Client Tarif: {item.ClientTarif}, Transportator Tarif: {item.TransportatorTarif}, Created At: {item.CreatedAt}, Order Number: {item.NumarComanda}");
            return Task.CompletedTask;
        }

    }
}

