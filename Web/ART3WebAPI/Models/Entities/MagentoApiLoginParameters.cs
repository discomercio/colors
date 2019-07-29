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
	}
}