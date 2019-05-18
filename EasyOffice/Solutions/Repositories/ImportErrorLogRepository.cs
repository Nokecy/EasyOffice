using EasyOffice.Models.Excel;
using EasyOffice.Solutions.DbContexts;
using EasyOffice.Solutions.Entities;
using EasyOffice.Solutions.Repositories;
using Bayantu.Extensions.Persistence.Dapper;
using Bayantu.Extensions.Session;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyOffice.Solutions.Repositories
{
    public class ImportErrorLogRepository : IImportErrorLogRepository
    {
        private readonly OfficeDbContext _officeDbContext;
        private readonly IEvosSession _session;
        private readonly IDapperAsyncExecutor _dapperAsyncExecutor;
        public ImportErrorLogRepository(OfficeDbContext officeDbContext, IEvosSession session, IDapperAsyncExecutor dapperAsyncExecutor)
        {
            _officeDbContext = officeDbContext;
            _session = session;
            _dapperAsyncExecutor = dapperAsyncExecutor;
        }

        public async Task<string> AddImportErrorLogAsync(List<ExcelDataRow> excelDataRows, string tag = "")
        {
            if (string.IsNullOrWhiteSpace(tag))
            {
                tag = Guid.NewGuid().ToString();
            }

            var importLogs = new List<ExcelImportErrorLog>();
            excelDataRows.Where(x => !x.IsValid).ToList().ForEach(x =>
            {
                importLogs.Add(new ExcelImportErrorLog()
                {
                    CreatedDate = DateTime.Now,
                    CreatedUserId = _session.EvosClaims.UserId,
                    LatestUpdatedDate = DateTime.Now,
                    LatestUpdatedUserId = _session.EvosClaims.UserId,
                    Message = x.ErrorMsg.Trim(';'),
                    TenantId = _session.EvosClaims.TenantId,
                    TopOrgId = _session.EvosClaims.TopOrgId,
                    IsDeleted = false,
                    Tag = tag,
                    RowNumber = x.RowIndex + 1
                });
            });
            await _officeDbContext.AddRangeAsync(importLogs);
            await _officeDbContext.SaveChangesAsync();

            return tag;
        }

        public async Task<IEnumerable<ExcelImportErrorLog>> GetErrorLogsByTagAsync(string tag)
        {
            string sql = "select id,row_number,message from excel_import_error_log where tag=@tag and is_deleted=0 order by row_number asc";
            return await _dapperAsyncExecutor.QueryAsync<ExcelImportErrorLog>(sql, new { tag });
        }
    }
}
