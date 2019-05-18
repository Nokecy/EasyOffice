using EasyOffice.Solutions.Entities;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Solutions.DbContexts
{
    public class OfficeDbContext : DbContext
    {
        public virtual DbSet<ExcelImportErrorLog> ExcelImportLog { get; set; }

        public OfficeDbContext()
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="contextOptions"></param>
        public OfficeDbContext(DbContextOptions<OfficeDbContext> contextOptions) : base(contextOptions)
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="builder"></param>
        protected override void OnModelCreating(ModelBuilder builder)
        {
            builder.Entity<ExcelImportErrorLog>(ConfigureExcelImportLog);
        }

        private void ConfigureExcelImportLog(EntityTypeBuilder<ExcelImportErrorLog> entity)
        {
            entity.ToTable("excel_import_error_log");

            entity.Property(e => e.Id)
               .HasColumnName("id")
               .HasColumnType("varchar(50)");

            entity.Property(e => e.CreatedDate)
                .HasColumnName("created_date")
                .HasColumnType("datetime");

            entity.Property(e => e.CreatedUserId)
                .HasColumnName("created_user_id")
                .HasColumnType("varchar(50)");

            entity.Property(e => e.IsDeleted)
                .HasColumnName("is_deleted")
                .HasDefaultValueSql("'0'");

            entity.Property(e => e.LatestUpdatedDate)
                .HasColumnName("latest_updated_date")
                .HasColumnType("datetime");

            entity.Property(e => e.LatestUpdatedUserId)
                .HasColumnName("latest_updated_user_id")
                .HasColumnType("varchar(50)");

            entity.Property(e => e.TenantId)
                .HasColumnName("tenant_id")
                .HasColumnType("varchar(50)");

            entity.Property(e => e.TopOrgId)
                .HasColumnName("top_org_id")
                .HasColumnType("varchar(50)");

            entity.Property(e => e.Tag)
               .HasColumnName("tag")
               .HasColumnType("varchar(50)");

            entity.Property(e => e.Message)
              .HasColumnName("message")
              .HasColumnType("text");

            entity.Property(e => e.RowNumber)
            .HasColumnName("row_number")
            .HasColumnType("int(11)");
        }
    }
}
