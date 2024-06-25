using Microsoft.EntityFrameworkCore;

namespace WPFStarter.Models
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }

        public DbSet<DataRecord> Records { get; set; }
    }
}