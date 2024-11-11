using Microsoft.EntityFrameworkCore;

namespace ExcelSheetPractice.Api.Contexts
{
    public class AppDbContext : DbContext
    {
        public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }

    }
}
