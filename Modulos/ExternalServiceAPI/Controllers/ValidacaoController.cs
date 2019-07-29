#region [ using ]
using System;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web.Http;
using System.Web.Script.Serialization;
using System.ServiceModel.Channels;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Web;
using ExternalServiceAPI.Models.Domains;
#endregion

namespace ExternalServiceAPI.Controllers
{
	#region [ ObjetoIE ]
	public class ObjetoIE
	{
		public string NumIE { get; set; }
		public string Uf { get; set; }
		public int Resultado { get; set; }
	}
	#endregion

	#region [ ValidacaoController ]
	public class ValidacaoController : ApiController
	{
		#region [ GetIp ]
		private string GetIp()
		{
			return GetClientIp();
		}
		#endregion

		#region [ GetClientIp ]
		private string GetClientIp(HttpRequestMessage request = null)
		{
			request = request ?? Request;

			if (request.Properties.ContainsKey("MS_HttpContext"))
			{
				return ((HttpContextWrapper)request.Properties["MS_HttpContext"]).Request.UserHostAddress;
			}
			else if (request.Properties.ContainsKey(RemoteEndpointMessageProperty.Name))
			{
				RemoteEndpointMessageProperty prop = (RemoteEndpointMessageProperty)request.Properties[RemoteEndpointMessageProperty.Name];
				return prop.Address;
			}
			else if (HttpContext.Current != null)
			{
				return HttpContext.Current.Request.UserHostAddress;
			}
			else
			{
				return null;
			}
		}
		#endregion

		#region [ Teste ]
		[HttpGet]
		public HttpResponseMessage Teste()
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ValidacaoController.Teste()";
			string strLogAtividade = NOME_DESTA_ROTINA + " (origem: " + (GetIp() ?? "") + ")";
			HttpResponseMessage result;
			#endregion

			try
			{
				result = Request.CreateResponse<string>(HttpStatusCode.OK, "Versão: " + Global.Cte.Aplicativo.M_ID);
				return result;
			}
			catch (Exception ex)
			{
				strLogAtividade += " - Exception: " + ex.Message;

				result = new HttpResponseMessage();
				result.StatusCode = HttpStatusCode.InternalServerError;
				result.Content = new StringContent(ex.ToString(), Encoding.UTF8, "application/json");

				return result;
			}
			finally
			{
				Global.gravaLogAtividade(strLogAtividade);
			}
		}
		#endregion

		[HttpGet]
		// GET: api/Validacao/IE?ie=xxxxxxxx&uf=xx

		#region [ IE(string ie, string uf) ]
		public HttpResponseMessage IE(string ie, string uf)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ValidacaoController.IE()";
			string strLogAtividade = NOME_DESTA_ROTINA + " (origem: " + (GetIp() ?? "") + ")";
			string strJson;
			HttpResponseMessage resultado;
			#endregion

			try // Try-Catch-Finally
			{
				strLogAtividade += ": IE=" + (ie ?? "").Trim() + ", UF=" + (uf ?? "").Trim();

				ObjetoIE objIE = new ObjetoIE { NumIE = Models.Domains.Global.digitos(ie), Uf = uf };
				ComPlusWrapper_DllInscE32.ComPlusWrapper_DllInscE32 obj;
				obj = new ComPlusWrapper_DllInscE32.ComPlusWrapper_DllInscE32();

				objIE.Resultado = obj.ConsisteInscricaoEstadual(objIE.NumIE, objIE.Uf);

				strLogAtividade += ", Resultado=" + objIE.Resultado.ToString();

				strJson = JsonConvert.SerializeObject(objIE, Formatting.Indented);

				resultado = new HttpResponseMessage();
				resultado.StatusCode = HttpStatusCode.OK;
				resultado.Content = new StringContent(strJson, Encoding.UTF8, "application/json");

				return resultado;
			}
			catch (Exception ex)
			{
				strLogAtividade += " - Exception: " + ex.Message;

				resultado = new HttpResponseMessage();
				resultado.StatusCode = HttpStatusCode.InternalServerError;
				resultado.Content = new StringContent(ex.ToString(), Encoding.UTF8, "application/json");

				return resultado;
			}
			finally
			{
				Global.gravaLogAtividade(strLogAtividade);
			}
		}
		#endregion
	}
	#endregion
}