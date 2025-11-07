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
    [Table("hitory")]
    public class HistoryTransport
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("id")]
        public int Id { get; set; }


        [Required]
        [MaxLength(255)]
        [Column("client")]
        public string ClientName { get; set; }

        [Required]
        [MaxLength(255)]
        [Column("route")]
        public string Route { get; set; }

        [Required]
        [Column("date_loaded")]
        public DateTime DateLoaded { get; set; }

        [Required]
        [Column("date_unloaded")]
        public DateTime DateUnloaded { get; set; }

        [Required]
        [MaxLength(255)]
        [Column("client_tarif")]
        public string ClientTarif { get; set; }

        [Required]
        [MaxLength(255)]
        [Column("transportator_tarif")]
        public string TransportatorTarif { get; set; }

        [Required]
        [Column("created_at")]
        public DateTime CreatedAt { get; set; }

        [Required]
        [Column("order_number")]
        public int NumarComanda { get; set; }

        public HistoryTransport() 
        {
            ClientName = string.Empty;
            Route = string.Empty;
            DateLoaded = DateTime.UtcNow;
            DateUnloaded = DateTime.UtcNow;
            ClientTarif = string.Empty;
            TransportatorTarif = string.Empty;
            CreatedAt = DateTime.UtcNow;
            NumarComanda = 0;
        }

    }
}

