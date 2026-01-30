using EmailCompleteApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;


namespace EmailCompleteApp.Services.Repositories
{
    class ProductRrepository
    {
        private static ProductRrepository? _instance;
        private static readonly object _lock = new object();
        public static ProductRrepository Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        _instance ??= new ProductRrepository();
                    }
                }
                return _instance;
            }
        }
        public ProductRrepository() { }

        public async Task InsertAsync(Product product)
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                var existingProduct = await context.Products
                    .FirstOrDefaultAsync(p => p.Name == product.Name);
                if (existingProduct == null)
                {
                    await context.Products.AddAsync(product);
                    await context.SaveChangesAsync();
                    System.Diagnostics.Debug.WriteLine($"✅ Inserted product '{product.Name}' into Supabase");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ Product '{product.Name}' already exists in Supabase");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error inserting product: {ex.Message}");
                throw;
            }
        }


        public async Task<List<Product>> LoadAllAsync()
        {
            try
            {
                using var context = DatabaseConfig.CreateDbContext();
                var products = await context.Products
                    .OrderBy(p => p.Name)
                    .AsNoTracking()
                    .ToListAsync();
                System.Diagnostics.Debug.WriteLine($"📊 Loaded {products.Count} products from Supabase");
                return products;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ Error loading products: {ex.Message}");
                throw;
            }
        }
    }
}
