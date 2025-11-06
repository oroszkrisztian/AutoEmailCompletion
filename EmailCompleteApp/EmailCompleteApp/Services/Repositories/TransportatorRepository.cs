using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using EmailCompleteApp.Models;

namespace EmailCompleteApp.Services.Repositories
{
    /// <summary>
    /// Repository for Transportator database operations
    /// </summary>
    public class TransportatorRepository
    {
        private static TransportatorRepository? _instance;
        private static readonly object _lock = new object();

        public static TransportatorRepository Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        _instance ??= new TransportatorRepository();
                    }
                }
                return _instance;
            }
        }

        private TransportatorRepository() { }

        /// <summary>
        /// Load all transportators from Supabase, ordered by name
        /// </summary>
        public async Task<List<Transportator>> LoadAllAsync()
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                var transportators = await context.Transportators
                    .OrderBy(t => t.Name)
                    .AsNoTracking()
                    .ToListAsync();

                System.Diagnostics.Debug.WriteLine($"📊 Loaded {transportators.Count} transportators from Supabase");
                return transportators;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error loading transportators: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Insert a new transportator to Supabase
        /// </summary>
        public async Task<Transportator> InsertAsync(Transportator transportator)
        {
            try
            {
                transportator.CreatedAt = DateTime.UtcNow;

                using var context = DatabaseConfig.CreateDbContext();
                context.Transportators.Add(transportator);
                await context.SaveChangesAsync();

                System.Diagnostics.Debug.WriteLine($"✅ Transportator '{transportator.Name}' saved to Supabase (ID: {transportator.Id})");
                return transportator;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error inserting transportator: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Update an existing transportator in Supabase
        /// </summary>
        public async Task<Transportator> UpdateAsync(Transportator transportator)
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                context.Transportators.Update(transportator);
                await context.SaveChangesAsync();

                System.Diagnostics.Debug.WriteLine($"✅ Transportator '{transportator.Name}' updated in Supabase");
                return transportator;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error updating transportator: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Delete a transportator from Supabase by ID
        /// </summary>
        public async Task DeleteAsync(int transportatorId)
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                var transportator = await context.Transportators.FindAsync(transportatorId);

                if (transportator != null)
                {
                    context.Transportators.Remove(transportator);
                    await context.SaveChangesAsync();
                    System.Diagnostics.Debug.WriteLine($"✅ Transportator ID {transportatorId} deleted from Supabase");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error deleting transportator: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Get a transportator by exact name match
        /// </summary>
        public async Task<Transportator?> GetByNameAsync(string name)
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                return await context.Transportators
                    .AsNoTracking()
                    .FirstOrDefaultAsync(t => t.Name == name);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error getting transportator by name: {ex.Message}");
                return null;
            }
        }
    }
}