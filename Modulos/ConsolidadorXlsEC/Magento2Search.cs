using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	public class Magento2SearchCriteria
	{
		public List<Magento2SearchCriteriaFilterGroups> filter_groups { get; set; }
	}

	public class Magento2SearchCriteriaFilterGroups
	{
		public List<Magento2SearchCriteriaFilterGroupsFilters> filters { get; set; }
	}

	public class Magento2SearchCriteriaFilterGroupsFilters
	{
		public string field { get; set; }
		public string value { get; set; }
		public string condition_type { get; set; }
	}
}
