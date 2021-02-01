using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ART3WebAPI.Models.Entities
{
	public class MagentoApiLoginParameters
	{
		public string urlWebService { get; set; } = "";
		public string username { get; set; } = "";
		public string password { get; set; } = "";
		public int api_versao { get; set; } = 0;
		public string api_rest_endpoint { get; set; } = "";
		public string api_rest_access_token { get; set; } = "";
		public byte api_rest_force_get_sales_order_by_entity_id { get; set; }
	}
}