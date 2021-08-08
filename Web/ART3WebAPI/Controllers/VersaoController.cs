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
			const string NOME_DESTA_ROTINA = "VersaoController.GetVersao()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;

			msg = NOME_DESTA_ROTINA + ": Requisição recebida";
			Global.gravaLogAtividade(httpRequestId, msg);

			HttpResponseMessage result = Request.CreateResponse<string>(HttpStatusCode.OK, "Versão: " + Global.Cte.Versao.M_ID);

			msg = NOME_DESTA_ROTINA + ": Status=" + result.StatusCode.ToString();
			Global.gravaLogAtividade(httpRequestId, msg);

			return result;
		}
		#endregion
	}
}
