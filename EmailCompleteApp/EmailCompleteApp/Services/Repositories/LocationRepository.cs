using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using EmailCompleteApp.Models;

namespace EmailCompleteApp.Services.Repositories
{
    /// <summary>
    /// Repository for Location database operations
    /// </summary>
    public class LocationRepository
    {
        private static LocationRepository? _instance;
        private static readonly object _lock = new object();

        public static LocationRepository Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        _instance ??= new LocationRepository();
                    }
                }
                return _instance;
            }
        }

        private LocationRepository() { }

        /// <summary>
        /// Load all locations from Supabase, ordered by name
        /// </summary>
        public async Task<List<Location>> LoadAllAsync()
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                var locations = await context.Locations
                    .OrderBy(l => l.Name)
                    .AsNoTracking()
                    .ToListAsync();

                System.Diagnostics.Debug.WriteLine($"📊 Loaded {locations.Count} locations from Supabase");
                return locations;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error loading locations: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Insert a new location to Supabase
        /// </summary>
        public async Task<Location> InsertAsync(Location location)
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                context.Locations.Add(location);
                await context.SaveChangesAsync();

                System.Diagnostics.Debug.WriteLine($"✅ Location '{location.Name}' saved to Supabase (ID: {location.Id})");
                return location;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error inserting location: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Update an existing location in Supabase
        /// </summary>
        public async Task<Location> UpdateAsync(Location location)
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                context.Locations.Update(location);
                await context.SaveChangesAsync();

                System.Diagnostics.Debug.WriteLine($"✅ Location '{location.Name}' updated in Supabase");
                return location;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error updating location: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Delete a location from Supabase by ID
        /// </summary>
        public async Task DeleteAsync(int locationId)
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                var location = await context.Locations.FindAsync(locationId);

                if (location != null)
                {
                    context.Locations.Remove(location);
                    await context.SaveChangesAsync();
                    System.Diagnostics.Debug.WriteLine($"✅ Location ID {locationId} deleted from Supabase");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error deleting location: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Get a location by exact name match
        /// </summary>
        public async Task<Location?> GetByNameAsync(string name)
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                return await context.Locations
                    .AsNoTracking()
                    .FirstOrDefaultAsync(l => l.Name == name);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error getting location by name: {ex.Message}");
                return null;
            }
        }
    }
}