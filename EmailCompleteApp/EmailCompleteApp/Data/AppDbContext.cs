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

        public DbSet<HistoryTransport> HistoryTransports { get; set; }

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
                entity.Property(e => e.CountryCode).HasColumnName("country_code").HasMaxLength(10);
                entity.Property(e => e.PostalCode).HasColumnName("postal_code").HasMaxLength(20);
                entity.Property(e => e.County).HasColumnName("county").HasMaxLength(100);
            });

            
            modelBuilder.Entity<HistoryTransport>(entity =>
            {
                entity.ToTable("history");
                entity.HasKey(e => e.Id);
                entity.Property(e => e.Id).HasColumnName("id").UseIdentityColumn();
                
                // Comanda / Client
                entity.Property(e => e.NumarComanda).HasColumnName("numar_comanda").IsRequired();
                entity.Property(e => e.NumarClient).HasColumnName("numar_client");
                entity.Property(e => e.Client).HasColumnName("client");
                
                // Tarif client
                entity.Property(e => e.Tarif).HasColumnName("tarif").HasPrecision(12, 2);
                entity.Property(e => e.MonedaIndex).HasColumnName("moneda_index");
                entity.Property(e => e.TipIndex).HasColumnName("tip_index");
                
                // Transportator
                entity.Property(e => e.Transportator).HasColumnName("transportator");
                entity.Property(e => e.TransportatorTarif).HasColumnName("transportator_tarif").HasPrecision(12, 2);
                entity.Property(e => e.TransportatorMonedaIndex).HasColumnName("transportator_moneda_index");
                entity.Property(e => e.TransportatorTipIndex).HasColumnName("transportator_tip_index");
                
                // Date
                entity.Property(e => e.DataIncarcare).HasColumnName("data_incarcare");
                entity.Property(e => e.DataDescarcare).HasColumnName("data_descarcare");
                
                // Marfa
                entity.Property(e => e.Produs).HasColumnName("produs");
                entity.Property(e => e.Cantitate).HasColumnName("cantitate").HasPrecision(12, 3);
                entity.Property(e => e.TipAdrIndex).HasColumnName("tip_adr_index");
                entity.Property(e => e.Clasa).HasColumnName("clasa");
                entity.Property(e => e.Un).HasColumnName("un");
                entity.Property(e => e.NumarInmatriculare).HasColumnName("numar_inmatriculare");
                
                // Locatie incarcare
                entity.Property(e => e.LocatieIncarcareAddress).HasColumnName("locatie_incarcare_address");
                entity.Property(e => e.LocatieIncarcareName).HasColumnName("locatie_incarcare_name");
                entity.Property(e => e.LocatieIncarcareCity).HasColumnName("locatie_incarcare_city");
                entity.Property(e => e.LocatieIncarcareCountryCode).HasColumnName("locatie_incarcare_country_code");
                entity.Property(e => e.LocatieIncarcarePostalCode).HasColumnName("locatie_incarcare_postal_code");
                entity.Property(e => e.LocatieIncarcareCounty).HasColumnName("locatie_incarcare_county");
                
                // Locatie descarcare
                entity.Property(e => e.LocatieDescarcareAddress).HasColumnName("locatie_descarcare_address");
                entity.Property(e => e.LocatieDescarcareName).HasColumnName("locatie_descarcare_name");
                entity.Property(e => e.LocatieDescarcareCity).HasColumnName("locatie_descarcare_city");
                entity.Property(e => e.LocatieDescarcareCountryCode).HasColumnName("locatie_descarcare_country_code");
                entity.Property(e => e.LocatieDescarcarePostalCode).HasColumnName("locatie_descarcare_postal_code");
                entity.Property(e => e.LocatieDescarcareCounty).HasColumnName("locatie_descarcare_county");
                
                // Alte informatii
                entity.Property(e => e.TermenPlata).HasColumnName("termen_plata");
                entity.Property(e => e.CommentUser).HasColumnName("comment_user");
                
                entity.Property(e => e.CreatedAt).HasColumnName("created_at").IsRequired().HasDefaultValueSql("CURRENT_TIMESTAMP");
            });
        }
    }
}
