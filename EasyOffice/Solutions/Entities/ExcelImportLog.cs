using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Solutions.Entities
{
    public class ExcelImportErrorLog
    {
        public string Id { get; set; }

        public string TenantId { get; set; } = "31";


        public string TopOrgId { get; set; }

        public string CreatedUserId { get; set; }

        public DateTime CreatedDate { get; set; }

        public string LatestUpdatedUserId { get; set; }

        public DateTime LatestUpdatedDate { get; set; }

        public bool IsDeleted { get; set; }

        public string Tag { get; set; }
        public string Message { get; set; }
        public int RowNumber { get; set; }
    }
}
