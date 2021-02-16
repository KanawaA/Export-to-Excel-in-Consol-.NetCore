using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication.Models
{
    public class Enrollment
    {
        [Key]
        public int Id { get; set; }

        public Student student { get; set; }
        public int StudentId { get; set; }
        public Subject subject { get; set; }
        public int SubjectId { get; set; }


        [Required]
        public string TeacherName { get; set; }
    }
}
