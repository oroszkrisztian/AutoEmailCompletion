using EmailCompleteApp.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailCompleteApp.Data
{
    public class AppDbContext: DbContext
    {
        public DbSet<Client> Clients { get; set; }
        public DbSet<Transportator> Transportators { get; set; }
        public DbSet<Location> Locations { get; set; }

        public AppDbContext(DbContextOptions<AppDbContext> options) : base(options)
        {
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);
            //Client
            modelBuilder.Entity<Client>(entity =>
            {
                entity.ToTable("clients");
                entity.HasKey(e => e.Id);
                entity.Property(e => e.Id).HasColumnName("id").UseIdentityColumn();
                entity.Property(e => e.Name).HasColumnName("name").IsRequired().HasMaxLength(255);
                entity.Property(e => e.Address).HasColumnName("address").IsRequired().HasMaxLength(500);
                entity.Property(e => e.Bank).HasColumnName("bank").HasMaxLength(255);
                entity.Property(e => e.IBAN).HasColumnName("iban").HasMaxLength(100);
                entity.Property(e => e.VATNumber).HasColumnName("vat_number").HasMaxLength(100);
                entity.Property(e => e.CameraDeComert).HasColumnName("camera_de_comert").HasMaxLength(255);
                entity.Property(e => e.TermenulDePlata).HasColumnName("termen_plata").HasMaxLength(100);
                entity.Property(e => e.CreatedAt).HasColumnName("created_at").HasDefaultValueSql("CURRENT_TIMESTAMP");
                entity.HasIndex(e => e.Name).HasDatabaseName("idx_clients_name");
            });

            // Configure Transportator entity
            modelBuilder.Entity<Transportator>(entity =>
            {
                entity.ToTable("transportators");
                entity.HasKey(e => e.Id);
                entity.Property(e => e.Id).HasColumnName("id").UseIdentityColumn();
                entity.Property(e => e.Name).HasColumnName("name").IsRequired().HasMaxLength(255);
                entity.Property(e => e.Address).HasColumnName("address").IsRequired().HasMaxLength(500);
                entity.Property(e => e.Bank).HasColumnName("bank").HasMaxLength(255);
                entity.Property(e => e.IBAN).HasColumnName("iban").HasMaxLength(100);
                entity.Property(e => e.VATNumber).HasColumnName("vat_number").HasMaxLength(100);
                entity.Property(e => e.CameraDeComert).HasColumnName("camera_de_comert").HasMaxLength(255);
                entity.Property(e => e.TermenulDePlata).HasColumnName("termen_plata").HasMaxLength(100);
                entity.Property(e => e.CreatedAt).HasColumnName("created_at").HasDefaultValueSql("CURRENT_TIMESTAMP");
                entity.HasIndex(e => e.Name).HasDatabaseName("idx_transportators_name");
            });

            // Configure Location entity
            modelBuilder.Entity<Location>(entity =>
            {
                entity.ToTable("locations");
                entity.HasKey(e => e.Id);
                entity.Property(e => e.Id).HasColumnName("id").UseIdentityColumn();
                entity.Property(e => e.Name).HasColumnName("name").IsRequired().HasMaxLength(255);
                entity.Property(e => e.Address).HasColumnName("address").IsRequired().HasMaxLength(500);
                entity.Property(e => e.City).HasColumnName("city").HasMaxLength(100);
                entity.HasIndex(e => e.Name).HasDatabaseName("idx_locations_name");
                entity.HasIndex(e => e.City).HasDatabaseName("idx_locations_city");
            });
        }
    }
}
