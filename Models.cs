using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace RequestManagerApp
{
    public class User
    {
        public int Id { get; set; }
        public string Username { get; set; } // Логин
        public string Password { get; set; } // Пароль (в реальных проектах нужно хэшировать!)
        public string Name { get; set; }     // ФИО
        public string Role { get; set; }     // "Admin" или "User"
    }

    public class Facility
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string ResponsiblePerson { get; set; }
    }

    public class Service
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string WorkType { get; set; }
        public int DeadlineDays { get; set; }
    }

    public class Request
    {
        public int Id { get; set; }
        public int UserId { get; set; }
        public User User { get; set; }
        public int FacilityId { get; set; }
        public Facility Facility { get; set; }

        public string ServiceName { get; set; }
        public string WorkType { get; set; }
        public int DeadlineDays { get; set; }

        public int Quantity { get; set; }
        public string Status { get; set; } = "В очереди";
        public DateTime CreatedAt { get; set; } = DateTime.Now;

        public string FormattedDate => CreatedAt.ToString("dd/MM/yy HH:mm");
    }

    public class AppDbContext : DbContext
    {
        public DbSet<User> Users { get; set; }
        public DbSet<Facility> Facilities { get; set; }
        public DbSet<Service> Services { get; set; }
        public DbSet<Request> Requests { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            string connectionString = "Server=127.0.0.1;Port=3306;Database=RequestManager;Uid=root;Pwd=1234;";

            optionsBuilder.UseMySql(connectionString, ServerVersion.AutoDetect(connectionString));
        }
    }
}
