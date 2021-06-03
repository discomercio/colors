using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	public class Magento2ProductSearchResponse
	{
		public List<Magento2Product> items { get; set; }
		public Magento2SearchCriteria search_criteria { get; set; }
		public int total_count { get; set; }
	}
}
