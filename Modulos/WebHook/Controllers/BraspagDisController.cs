#region [ using ]
using System;
using System.Net;
using System.Net.Http;
using System.ServiceModel.Channels;
using System.Text;
using System.Web;
using System.Web.Http;
using WebHook.Models;
using WebHook.Models.Repository;
#endregion

namespace WebHook.Controllers
{
	#region [ BraspagDisController ]
	public class BraspagDisController : ApiController
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
			const string NOME_DESTA_ROTINA = "BraspagDisController.Teste()";
			string strLogAtividade = NOME_DESTA_ROTINA + " (origem: " + (GetIp() ?? "") + ")";
			HttpResponseMessage result;
			#endregion

			try
			{
				result = Request.CreateResponse<string>(HttpStatusCode.OK, "Versão: " + Models.Domains.Global.Cte.Aplicativo.M_ID);
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
				Models.Domains.Global.gravaLogAtividade(strLogAtividade);
			}
		}
		#endregion

		// POST: api/BraspagDis
		/// <summary>
		/// 
		/// </summary>
		/// <param name="NumPedido">Número do pedido do cliente (obrigatório)</param>
		/// <param name="Status">Status do Pagamento = "0" (obrigatório)</param>
		/// <param name="CODPAGAMENTO">Código do Meio de Pagamento (obrigatório)</param>
		/// <returns></returns>
		/// 
		#region [ Post(BraspagParameters parameters) ]
		public HttpResponseMessage Post([FromBody]BraspagParameters parameters)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDisController.Post()";
			string strLogAtividade = NOME_DESTA_ROTINA + " (origem: " + (GetIp() ?? "") + ")";
			int intLinhasAfetadas = 0;
			HttpResponseMessage resultado;
			#endregion

			try // Try-Catch-Finally
			{
				strLogAtividade += ": " + parameters.FormataDados(showOnePropertyPerLine: false, inlinePropertySeparator: ", ");

				#region [ Consistência ]
				if (parameters == null)
				{
					strLogAtividade += ", Resposta: 'não foram informados os parâmetros'";
					resultado = new HttpResponseMessage(HttpStatusCode.NotAcceptable);
					resultado.Content = new StringContent("não foram informados os parâmetros", Encoding.UTF8, "text/plain");
					return resultado;
				}

				if (string.IsNullOrEmpty(parameters.NumPedido))
				{
					strLogAtividade += ", Resposta: 'não foi informado o parâmetro NumPedido'";
					resultado = new HttpResponseMessage(HttpStatusCode.NotAcceptable);
					resultado.Content = new StringContent("não foi informado o parâmetro NumPedido", Encoding.UTF8, "text/plain");
					return resultado;
				}

				if (string.IsNullOrEmpty(parameters.Status))
				{
					strLogAtividade += ", Resposta: 'não foi informado o parâmetro Status'";
					resultado = new HttpResponseMessage(HttpStatusCode.NotAcceptable);
					resultado.Content = new StringContent("não foi informado o parâmetro Status", Encoding.UTF8, "text/plain");
					return resultado;
				}
				#endregion

				#region [ Grava ]
				try
				{
					intLinhasAfetadas = BD.Grava(parameters, empresa: "DIS");

					if (intLinhasAfetadas < 1)
					{
						strLogAtividade += ", Resposta: (vazio) [nenhum registro inserido no BD]";
						resultado = new HttpResponseMessage(HttpStatusCode.OK);
						resultado.Content = new StringContent("", Encoding.UTF8, "text/plain");
						return resultado;
					}
					else
					{
						strLogAtividade += ", Resposta: <status>OK</status>";
						resultado = new HttpResponseMessage(HttpStatusCode.OK);
						resultado.Content = new StringContent("<status>OK</status>", Encoding.UTF8, "text/plain");
						return resultado;
					}
				}
				catch (Exception ex)
				{
					strLogAtividade += " - Exception: " + ex.Message;
					resultado = new HttpResponseMessage(HttpStatusCode.InternalServerError);
					resultado.Content = new StringContent("", Encoding.UTF8, "text/plain");
					return resultado;
				}
				#endregion
			}
			catch (Exception ex)
			{
				strLogAtividade += " - Exception: " + ex.Message;
				resultado = new HttpResponseMessage(HttpStatusCode.InternalServerError);
				resultado.Content = new StringContent("", Encoding.UTF8, "text/plain");
				return resultado;
			}
			finally
			{
				Models.Domains.Global.gravaLogAtividade(strLogAtividade);
			}
		}
		#endregion
	}
	#endregion
}
