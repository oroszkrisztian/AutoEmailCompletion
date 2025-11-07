using CommunityToolkit.Mvvm.ComponentModel;
using EmailCompleteApp.Models;
using EmailCompleteApp.Services;
using EmailCompleteApp.Services.Repositories;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailCompleteApp.ViewModels
{
    public partial class HistoryPageViewModel: ObservableObject
    {
        private readonly HistoryRepository _historyRepository;
        public ObservableCollection<HistoryTransport> HistoryData { get; } = new ();
        public HistoryPageViewModel()
        {
            _historyRepository = HistoryRepository.Instance;
            _ = InitializeHistoryData();
        }

        private async Task InitializeHistoryData()
        {
            try
            {
                //debug console write
                Debug.WriteLine("Initializing history data...");
                var response = await _historyRepository.LoadAllAsync();
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
        }

        public static Task PrintLoadedData(HistoryTransport item)
        {
            Debug.WriteLine($"History Item - ID: {item.Id}, Client: {item.ClientName}, Route: {item.Route}, Date Loaded: {item.DateLoaded}, Date Unloaded: {item.DateUnloaded}, Client Tarif: {item.ClientTarif}, Transportator Tarif: {item.TransportatorTarif}, Created At: {item.CreatedAt}, Order Number: {item.NumarComanda}");
            return Task.CompletedTask;
        }

    }
}

