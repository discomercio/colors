using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using ART3WebAPI.Models.Domains;

namespace ART3WebAPI.Controllers
{
    public class VersaoController : ApiController
    {
		#region [ GetVersao ]
		[HttpGet]
		public HttpResponseMessage GetVersao()
		{
			HttpResponseMessage result = Request.CreateResponse<string>(HttpStatusCode.OK, "Versão: " + Global.Cte.Versao.M_ID);
			return result;
		}
		#endregion
	}
}
