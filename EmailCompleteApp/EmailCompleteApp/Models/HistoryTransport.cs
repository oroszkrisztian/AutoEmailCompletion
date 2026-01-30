using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace EmailCompleteApp.Models
{
    [Table("history")]
    public class HistoryTransport
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("id")]
        public int Id { get; set; }

        // Comanda / Client
        [Column("numar_comanda")]
        public string NumarComanda { get; set; }

        [Column("numar_client")]
        public string? NumarClient { get; set; }

        [Column("client")]
        public string? Client { get; set; }

        [Column("contact")]
        public string? Contact { get; set; }

        // Tarif client
        [Column("tarif")]
        public decimal? Tarif { get; set; }

        [Column("moneda_index")]
        public int? MonedaIndex { get; set; }

        [Column("tip_index")]
        public int? TipIndex { get; set; }

        // Transportator
        [Column("transportator")]
        public string? Transportator { get; set; }

        [Column("transportator_tarif")]
        public decimal? TransportatorTarif { get; set; }

        [Column("transportator_moneda_index")]
        public int? TransportatorMonedaIndex { get; set; }

        [Column("transportator_tip_index")]
        public int? TransportatorTipIndex { get; set; }

        // Date
        [Column("data_incarcare")]
        public DateTime? DataIncarcare { get; set; }

        [Column("data_descarcare")]
        public DateTime? DataDescarcare { get; set; }

        // Marfa
        [Column("produs")]
        public string? Produs { get; set; }

        [Column("cantitate")]
        public string? Cantitate { get; set; }

        [Column("tip_adr_index")]
        public int? TipAdrIndex { get; set; }

        [Column("clasa")]
        public string? Clasa { get; set; }

        [Column("un")]
        public string? Un { get; set; }

        [Column("numar_inmatriculare")]
        public string? NumarInmatriculare { get; set; }

        // Locatie incarcare
        [Column("locatie_incarcare_address")]
        public string? LocatieIncarcareAddress { get; set; }

        [Column("locatie_incarcare_name")]
        public string? LocatieIncarcareName { get; set; }

        [Column("locatie_incarcare_city")]
        public string? LocatieIncarcareCity { get; set; }

        [Column("locatie_incarcare_country_code")]
        public string? LocatieIncarcareCountryCode { get; set; }

        [Column("locatie_incarcare_postal_code")]
        public string? LocatieIncarcarePostalCode { get; set; }

        [Column("locatie_incarcare_county")]
        public string? LocatieIncarcareCounty { get; set; }

        // Locatie descarcare
        [Column("locatie_descarcare_address")]
        public string? LocatieDescarcareAddress { get; set; }

        [Column("locatie_descarcare_name")]
        public string? LocatieDescarcareName { get; set; }

        [Column("locatie_descarcare_city")]
        public string? LocatieDescarcareCity { get; set; }

        [Column("locatie_descarcare_country_code")]
        public string? LocatieDescarcareCountryCode { get; set; }

        [Column("locatie_descarcare_postal_code")]
        public string? LocatieDescarcarePostalCode { get; set; }

        [Column("locatie_descarcare_county")]
        public string? LocatieDescarcareCounty { get; set; }

        // Alte informatii
        [Column("termen_plata")]
        public int? TermenPlata { get; set; }

        [Column("comment_user")]
        public string? CommentUser { get; set; }

        [Required]
        [Column("created_at")]
        public DateTime CreatedAt { get; set; }

        // Display properties
        public string DisplayDataIncarcare => DataIncarcare?.Date.ToString("dd/MM/yyyy") ?? string.Empty;
        public string DisplayDataDescarcare => DataDescarcare?.Date.ToString("dd/MM/yyyy") ?? string.Empty;
        public string DisplayDateCreatedAt => CreatedAt.Date.ToString("dd/MM/yyyy");

        // Computed properties for backward compatibility with UI
        public string Route => $"{LocatieIncarcareCity ?? "?"} - {LocatieDescarcareCity ?? "?"}";
        public string DisplayDateLoaded => DisplayDataIncarcare;
        public string DisplayDateUnloaded => DisplayDataDescarcare;
        public string ClientName => Client ?? string.Empty;
        public string ClientTarif
        {
            get
            {
                if (!Tarif.HasValue) return string.Empty;
                var monedaOptions = new[] { "EUR", "EUR/MT", "RON" };
                var tipOptions = new[] { "TVA", "ALL IN" };
                var moneda = MonedaIndex.HasValue && MonedaIndex.Value >= 0 && MonedaIndex.Value < monedaOptions.Length 
                    ? monedaOptions[MonedaIndex.Value] 
                    : string.Empty;
                var tip = TipIndex.HasValue && TipIndex.Value >= 0 && TipIndex.Value < tipOptions.Length 
                    ? tipOptions[TipIndex.Value] 
                    : string.Empty;
                if (tip == "TVA")
                {
                    return $"{Tarif.Value} {moneda} + {tip}".Trim();
                }
                else
                {
                    return $"{Tarif.Value} {moneda} {tip}".Trim();
                }
            }
        }
        public string TransportatorTarifDisplay
        {
            get
            {
                if (!TransportatorTarif.HasValue) return string.Empty;
                var monedaOptions = new[] { "EUR", "EUR/MT", "RON" };
                var tipOptions = new[] { "TVA", "ALL IN" };
                var moneda = TransportatorMonedaIndex.HasValue && TransportatorMonedaIndex.Value >= 0 && TransportatorMonedaIndex.Value < monedaOptions.Length 
                    ? monedaOptions[TransportatorMonedaIndex.Value] 
                    : string.Empty;
                var tip = TransportatorTipIndex.HasValue && TransportatorTipIndex.Value >= 0 && TransportatorTipIndex.Value < tipOptions.Length 
                    ? tipOptions[TransportatorTipIndex.Value] 
                    : string.Empty;
                if (tip == "TVA")
                {
                    return $"{TransportatorTarif.Value} {moneda} + {tip}".Trim();
                }
                else
                {
                    return $"{TransportatorTarif.Value} {moneda} {tip}".Trim();
                }
            }
        }

        public string HistorySummary()
        {
            var route = $"{LocatieIncarcareCity ?? "?"} - {LocatieDescarcareCity ?? "?"}";
            return $" Nr: {NumarComanda} Ruta: {route} | Data: {DisplayDataIncarcare} - {DisplayDataDescarcare} ";
        }

        public HistoryTransport() 
        {
            NumarComanda = string.Empty;
            CreatedAt = DateTime.UtcNow;
        }
    }
}

