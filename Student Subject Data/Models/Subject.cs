using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication.Models
{
    public class Subject
    {
        [Key]
        public int Id { get; set; }
        [Required]
        public string SubjectName { get; set; }
        public ICollection<Enrollment> enrollments { get; set; }
    }
}
