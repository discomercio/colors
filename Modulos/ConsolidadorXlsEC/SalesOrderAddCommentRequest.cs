using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
    public class SalesOrderAddCommentRequest
    {
        public string orderIncrementId { get; set; }
        public string status { get; set; }
        public string comment { get; set; }
        public string notify { get; set; }
    }
}
