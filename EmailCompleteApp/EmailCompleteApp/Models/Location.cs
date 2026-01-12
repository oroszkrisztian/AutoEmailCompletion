using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace EmailCompleteApp.Models
{
    [Table("locations")]
    public class Location
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

        [MaxLength(100)]
        [Column("city")]
        public string City { get; set; }

        [MaxLength(10)]
        [Column("country_code")]
        public string? CountryCode { get; set; }

        [MaxLength(20)]
        [Column("postal_code")]
        public string? PostalCode { get; set; }

        [MaxLength(100)]
        [Column("county")]
        public string? County { get; set; }
        

       
        //override ToString 
        public override string ToString()
        {
            return $"{Name}, {Address}, {City} {CountryCode} - {PostalCode}";
        }

        public Location()
        {
            Name = string.Empty;
            Address = string.Empty;
            City = string.Empty;
            CountryCode = string.Empty;
            PostalCode = string.Empty;
            County = string.Empty;
        }

       
    }
}