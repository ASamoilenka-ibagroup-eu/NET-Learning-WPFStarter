using Microsoft.EntityFrameworkCore;

namespace WPFStarter.Models.Data
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