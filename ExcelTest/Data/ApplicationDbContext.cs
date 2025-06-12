using Microsoft.EntityFrameworkCore;

namespace ExcelTest.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options)
        {
        }
        public DbSet<Models.Student> Students { get; set; }
    }
    
}
