
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace NetMVC.Models
{
    [Table("Person")]
    public class Person
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public string? PersonId { get; set; } = string.Empty;

        public string? FullName { get; set; }
        public string? Address { get; set; }
    }
}
