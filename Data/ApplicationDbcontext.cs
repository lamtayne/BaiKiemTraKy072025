using Demo_Mvc.Models;
using DemoMVC.Models;
using Microsoft.EntityFrameworkCore;
using static Demo_Mvc.Models.HeThongPhanPhoi;
using DaiLy = DemoMVC.Models.DaiLy;
namespace Demo_Mvc.Data;

public class ApplicationDbContext : DbContext
{
    public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
        : base(options)
    {
    }
protected override void OnModelCreating(ModelBuilder modelBuilder)
{
    base.OnModelCreating(modelBuilder);

    modelBuilder.Entity<Person>()
        .HasDiscriminator<string>("Discriminator")
        .HasValue<Person>("Person")
        .HasValue<Employee>("Employee");
}


    public DbSet<Person> Persons { get; set; }
    public DbSet<Employee> Employees { get; set; }
    public DbSet<DaiLy> DaiLys { get; set; }
    public DbSet<HeThongPhanPhoi> HeThongPhanPhois { get; set; }
}
