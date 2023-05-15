using Microsoft.EntityFrameworkCore;
using MCHSAutoTable.Entities.Coworker;
using MCHSAutoTable.Entities.EDDS;

namespace MCHSAutoTable
{
    public sealed class ApplicationContext : DbContext
    {
        public DbSet<Edds> Edds => Set<Edds>();
        public DbSet<TableEdds> TableEdds => Set<TableEdds>();
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
