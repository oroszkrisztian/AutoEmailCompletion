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

        public async Task<HistoryTransport> InsertHistory(HistoryTransport history) 
        {
            try 
            { 
                history.CreatedAt = DateTime.UtcNow;
                using var context = DatabaseConfig.CreateDbContext();
                context.HistoryTransports.Add(history);
                await context.SaveChangesAsync();
                return history;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error inserting history: {ex.Message}");
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
    }
}
