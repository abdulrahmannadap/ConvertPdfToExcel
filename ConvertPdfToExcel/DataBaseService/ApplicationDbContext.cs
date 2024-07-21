using ConvertPdfToExcel.Models;
using Microsoft.EntityFrameworkCore;

namespace ConvertPdfToExcel.DataBaseService
{
    public class ApplicationDbContext:DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options):base(options)
        {
        }

        public DbSet<ConvertedExcelData> ConvertedExcelDatas { get; set; }
        public DbSet<ConvertedFile> ConvertedFiles { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ConvertedExcelData>().ToTable("ConvertedExcelDataTable");
            modelBuilder.Entity<ConvertedFile>().ToTable("ConvertedFileTable");
        }
    }
}
