using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebApplication.Models;

namespace WebApplication.Context
{
    public class MyContext : DbContext
    {
        public MyContext(DbContextOptions<MyContext> options) : base(options)
        {

        }

        public DbSet<Student> students { get; set; }
        public DbSet<Subject> subjects { get; set; }
        public DbSet<Enrollment> enrollments { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Enrollment>()
                .HasOne(st => st.student)
                .WithMany(st => st.enrollments)
                .HasForeignKey(st => st.StudentId);
            modelBuilder.Entity<Enrollment>()
                .HasOne(st => st.subject)
                .WithMany(st => st.enrollments)
                .HasForeignKey(st => st.SubjectId);
        }
    }
}
