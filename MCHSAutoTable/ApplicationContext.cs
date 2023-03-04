using MCHSAutoTable.Entityes.coworker;
using MCHSAutoTable.Entityes.edds;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MCHSAutoTable
{
    public class ApplicationContext : DbContext
    {
        public DbSet<EDDS> EDDS => Set<EDDS>();
        public DbSet<TableEDDS> TableEDDS => Set<TableEDDS>();
        public DbSet<Staff> Staffs => Set<Staff>();
        public DbSet<Position> Positions => Set<Position>();
        public DbSet<SubDepartment> SubDepartments => Set<SubDepartment>();
        public DbSet<Patients> Patients => Set<Patients>();
        public DbSet<Diagnosis> Diagnoses => Set<Diagnosis>();

        public ApplicationContext() => Database.EnsureCreated();

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlite("Data Source=DATABASE.db");
        }
    }
}
