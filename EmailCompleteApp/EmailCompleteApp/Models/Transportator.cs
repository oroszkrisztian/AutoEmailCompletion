using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace EmailCompleteApp.Models
{
    [Table("transportators")]
    public class Transportator
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("id")]
        public int Id { get; set; }

        [Required]
        [MaxLength(255)]
        [Column("name")]
        public string Name { get; set; }

        [Required]
        [MaxLength(500)]
        [Column("address")]
        public string Address { get; set; }

        [MaxLength(255)]
        [Column("bank")]
        public string Bank { get; set; }

        [MaxLength(100)]
        [Column("iban")]
        public string IBAN { get; set; }

        [MaxLength(100)]
        [Column("vat_number")]
        public string VATNumber { get; set; }

        [MaxLength(255)]
        [Column("camera_de_comert")]
        public string CameraDeComert { get; set; }

        [MaxLength(100)]
        [Column("termen_plata")]
        public string TermenulDePlata { get; set; }

        [Column("created_at")]
        public DateTime CreatedAt { get; set; }

        public override string ToString()
        {
            return Name;
        }

        public Transportator()
        {
            Name = string.Empty;
            Address = string.Empty;
            Bank = string.Empty;
            IBAN = string.Empty;
            VATNumber = string.Empty;
            CameraDeComert = string.Empty;
            TermenulDePlata = string.Empty;
            CreatedAt = DateTime.UtcNow;
        }

        public Transportator(int id, string name, string address, string bank, string iban,
                            string vatNumber, string cameraDeComert, string termenulDePlata)
        {
            Id = id;
            Name = name;
            Address = address;
            Bank = bank;
            IBAN = iban;
            VATNumber = vatNumber;
            CameraDeComert = cameraDeComert;
            TermenulDePlata = termenulDePlata;
            CreatedAt = DateTime.UtcNow;
        }
    }
}