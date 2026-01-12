using EmailCompleteApp.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailCompleteApp.Services.Repositories
{
    internal class HistoryRepository
    {
        private static HistoryRepository? _instance;
        private static readonly object _lock = new object();

        public static HistoryRepository Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        _instance ??= new HistoryRepository();
                    }
                }
                return _instance;
            }
        }

        private HistoryRepository() { }
        
        public async Task<HistoryTransport> GetLastOrder()
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                var allOrders = await context.HistoryTransports.ToListAsync();
                
                var lastOrder = allOrders
                    .OrderByDescending(h => int.TryParse(h.NumarComanda, out var num) ? num : 0)
                    .FirstOrDefault();
                
                if (lastOrder != null)
                {
                    System.Diagnostics.Debug.WriteLine($"🔢 Last order number retrieved: {lastOrder.NumarComanda}");
                    return lastOrder;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("ℹ️ No previous orders found. Returning default order.");
                    return new HistoryTransport { NumarComanda = "0" };
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error retrieving last order number: {ex.Message}");
                throw;
            }
        }
        
        public async Task<HistoryTransport> InsertHistory(HistoryTransport history) 
        {
            try 
            { 
                history.CreatedAt = DateTime.UtcNow;
                using var context = DatabaseConfig.CreateDbContext();
                var existing = await context.HistoryTransports
                    .FirstOrDefaultAsync(h => h.NumarComanda == history.NumarComanda);
                
                if(existing == null)
                {
                    context.HistoryTransports.Add(history);
                    await context.SaveChangesAsync();
                    System.Diagnostics.Debug.WriteLine($"✅ History record inserted: Order #{history.NumarComanda}");
                    return history;
                }
                else
                {
                    // Update all fields
                    existing.NumarClient = history.NumarClient;
                    existing.Client = history.Client;
                    existing.Tarif = history.Tarif;
                    existing.MonedaIndex = history.MonedaIndex;
                    existing.TipIndex = history.TipIndex;
                    existing.Transportator = history.Transportator;
                    existing.TransportatorTarif = history.TransportatorTarif;
                    existing.TransportatorMonedaIndex = history.TransportatorMonedaIndex;
                    existing.TransportatorTipIndex = history.TransportatorTipIndex;
                    existing.DataIncarcare = history.DataIncarcare;
                    existing.DataDescarcare = history.DataDescarcare;
                    existing.Produs = history.Produs;
                    existing.Cantitate = history.Cantitate;
                    existing.TipAdrIndex = history.TipAdrIndex;
                    existing.Clasa = history.Clasa;
                    existing.Un = history.Un;
                    existing.NumarInmatriculare = history.NumarInmatriculare;
                    existing.LocatieIncarcareAddress = history.LocatieIncarcareAddress;
                    existing.LocatieIncarcareName = history.LocatieIncarcareName;
                    existing.LocatieIncarcareCity = history.LocatieIncarcareCity;
                    existing.LocatieIncarcareCountryCode = history.LocatieIncarcareCountryCode;
                    existing.LocatieIncarcarePostalCode = history.LocatieIncarcarePostalCode;
                    existing.LocatieIncarcareCounty = history.LocatieIncarcareCounty;
                    existing.LocatieDescarcareAddress = history.LocatieDescarcareAddress;
                    existing.LocatieDescarcareName = history.LocatieDescarcareName;
                    existing.LocatieDescarcareCity = history.LocatieDescarcareCity;
                    existing.LocatieDescarcareCountryCode = history.LocatieDescarcareCountryCode;
                    existing.LocatieDescarcarePostalCode = history.LocatieDescarcarePostalCode;
                    existing.LocatieDescarcareCounty = history.LocatieDescarcareCounty;
                    existing.TermenPlata = history.TermenPlata;
                    existing.CommentUser = history.CommentUser;
                    existing.CreatedAt = history.CreatedAt;

                    await context.SaveChangesAsync();
                    System.Diagnostics.Debug.WriteLine($"✅ History record updated: Order #{history.NumarComanda}");
                    return existing;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error inserting/updating history: {ex.Message}");
                throw;
            }
        }

        public async Task<HistoryTransport?> UpdateHistory(HistoryTransport history)
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                var existing = await context.HistoryTransports
                    .FirstOrDefaultAsync(h => h.Id == history.Id);

                if (existing == null)
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ History record not found for update: ID #{history.Id}");
                    return null;
                }

                // Update all fields
                existing.NumarComanda = history.NumarComanda;
                existing.NumarClient = history.NumarClient;
                existing.Client = history.Client;
                existing.Tarif = history.Tarif;
                existing.MonedaIndex = history.MonedaIndex;
                existing.TipIndex = history.TipIndex;
                existing.Transportator = history.Transportator;
                existing.TransportatorTarif = history.TransportatorTarif;
                existing.TransportatorMonedaIndex = history.TransportatorMonedaIndex;
                existing.TransportatorTipIndex = history.TransportatorTipIndex;
                existing.DataIncarcare = history.DataIncarcare;
                existing.DataDescarcare = history.DataDescarcare;
                existing.Produs = history.Produs;
                existing.Cantitate = history.Cantitate;
                existing.TipAdrIndex = history.TipAdrIndex;
                existing.Clasa = history.Clasa;
                existing.Un = history.Un;
                existing.NumarInmatriculare = history.NumarInmatriculare;
                existing.LocatieIncarcareAddress = history.LocatieIncarcareAddress;
                existing.LocatieIncarcareName = history.LocatieIncarcareName;
                existing.LocatieIncarcareCity = history.LocatieIncarcareCity;
                existing.LocatieIncarcareCountryCode = history.LocatieIncarcareCountryCode;
                existing.LocatieIncarcarePostalCode = history.LocatieIncarcarePostalCode;
                existing.LocatieIncarcareCounty = history.LocatieIncarcareCounty;
                existing.LocatieDescarcareAddress = history.LocatieDescarcareAddress;
                existing.LocatieDescarcareName = history.LocatieDescarcareName;
                existing.LocatieDescarcareCity = history.LocatieDescarcareCity;
                existing.LocatieDescarcareCountryCode = history.LocatieDescarcareCountryCode;
                existing.LocatieDescarcarePostalCode = history.LocatieDescarcarePostalCode;
                existing.LocatieDescarcareCounty = history.LocatieDescarcareCounty;
                existing.TermenPlata = history.TermenPlata;
                existing.CommentUser = history.CommentUser;
                // Do not update CreatedAt - keep original

                await context.SaveChangesAsync();
                System.Diagnostics.Debug.WriteLine($"✅ History record updated: Order #{history.NumarComanda}");
                return existing;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error updating history: {ex.Message}");
                throw;
            }
        }

        public async Task<List<HistoryTransport>> LoadAllAsync()
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                var histories = await context.HistoryTransports
                    .OrderByDescending(h => h.CreatedAt)
                    .AsNoTracking()
                    .ToListAsync();
                System.Diagnostics.Debug.WriteLine($"📊 Loaded {histories.Count} history records from Supabase");
                return histories;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error loading history records: {ex.Message}");
                throw;
            }
        }

        public async Task<List<HistoryTransport>> LoadAllByOrderNumDescAsync()
        { 
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                var histories = await context.HistoryTransports
                    .AsNoTracking()
                    .ToListAsync();
                
                // Sort in memory after fetching
                var sorted = histories
                    .OrderByDescending(h => int.TryParse(h.NumarComanda, out var num) ? num : 0)
                    .ToList();
                
                System.Diagnostics.Debug.WriteLine($"📊 Loaded {sorted.Count} history records ordered by order number");
                return sorted;
            }
            catch(Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error loading history records descending by order number: {ex.Message}");
                throw;
            }
        }
    }
}
