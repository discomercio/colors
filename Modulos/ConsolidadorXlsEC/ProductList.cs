using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	class ProductList
	{
		public string product_id { get; set; }
		public string sku { get; set; }
		public string name { get; set; }
		public string set { get; set; }
		public string type { get; set; }
		public List<string> category_ids { get; set; } = new List<string>();
		public List<string> website_ids { get; set; } = new List<string>();
		public List<string> UnknownFields { get; set; } = new List<string>();
	}
}
