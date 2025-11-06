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

       
        //override ToString 
        public override string ToString()
        {
            return $"{Name}, {Address}, {City}";
        }

        public Location()
        {
            Name = string.Empty;
            Address = string.Empty;
            City = string.Empty;
        }

        public Location(int id, string name, string address, string city)
        {
            Id = id;
            Name = name;
            Address = address;
            City = city;
        }
    }
}