using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using EmailCompleteApp.Models;

namespace EmailCompleteApp.Services.Repositories
{
    
    public class ClientRepository
    {
        private static ClientRepository? _instance;
        private static readonly object _lock = new object();

        public static ClientRepository Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        _instance ??= new ClientRepository();
                    }
                }
                return _instance;
            }
        }

        private ClientRepository() { }

       
        public async Task<List<Client>> LoadAllAsync()
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                var clients = await context.Clients
                    .OrderBy(c => c.Name)
                    .AsNoTracking()
                    .ToListAsync();

                System.Diagnostics.Debug.WriteLine($"📊 Loaded {clients.Count} clients from Supabase");
                return clients;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error loading clients: {ex.Message}");
                throw;
            }
        }

       
        public async Task<Client> InsertAsync(Client client)
        {
            try
            {
                client.CreatedAt = DateTime.UtcNow;

                using var context = DatabaseConfig.CreateDbContext();
                context.Clients.Add(client);
                await context.SaveChangesAsync();

                System.Diagnostics.Debug.WriteLine($"✅ Client '{client.Name}' saved to Supabase (ID: {client.Id})");
                return client;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error inserting client: {ex.Message}");
                throw;
            }
        }

        
        public async Task<Client> UpdateAsync(Client client)
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                context.Clients.Update(client);
                await context.SaveChangesAsync();

                System.Diagnostics.Debug.WriteLine($"✅ Client '{client.Name}' updated in Supabase");
                return client;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error updating client: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Delete a client from Supabase by ID
        /// </summary>
        public async Task DeleteAsync(int clientId)
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                var client = await context.Clients.FindAsync(clientId);

                if (client != null)
                {
                    context.Clients.Remove(client);
                    await context.SaveChangesAsync();
                    System.Diagnostics.Debug.WriteLine($"✅ Client ID {clientId} deleted from Supabase");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error deleting client: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Get a client by exact name match
        /// </summary>
        public async Task<Client?> GetByNameAsync(string name)
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                return await context.Clients
                    .AsNoTracking()
                    .FirstOrDefaultAsync(c => c.Name == name);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error getting client by name: {ex.Message}");
                return null;
            }
        }
    }
}