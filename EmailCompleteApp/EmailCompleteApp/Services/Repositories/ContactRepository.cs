using EmailCompleteApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;

namespace EmailCompleteApp.Services.Repositories
{
    class ContactRepository
    {
        private static ContactRepository? _instance;
        private static readonly object _lock = new object();
        public static ContactRepository Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        _instance ??= new ContactRepository();
                    }
                }
                return _instance;
            }
        }

        private ContactRepository() { }

        //insert 
        public async Task InsertAsync(Contact contact)
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                var esistingContact = await context.Contacts
                    .FirstOrDefaultAsync(c => c.Name == contact.Name);
                if (esistingContact == null)
                    {
                    await context.Contacts.AddAsync(contact);
                    await context.SaveChangesAsync();
                    System.Diagnostics.Debug.WriteLine($"✅ Inserted contact '{contact.Name}' into Supabase");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ Contact '{contact.Name}' already exists in Supabase");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error inserting contact: {ex.Message}");
                throw;
            }
        }

        public async Task<List<Contact>> LoadAllAsync()
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                var contacts = await context.Contacts
                    .OrderBy(c => c.Name)
                    .AsNoTracking()
                    .ToListAsync();
                System.Diagnostics.Debug.WriteLine($"📊 Loaded {contacts.Count} contacts from Supabase");
                return contacts;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error loading contacts: {ex.Message}");
                throw;
            }
        }
    }
}
