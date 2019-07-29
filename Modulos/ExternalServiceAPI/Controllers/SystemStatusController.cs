#region [ using ]
using System;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web.Http;
using ExternalServiceAPI.Models.Repository;
using System.Web;
using System.ServiceModel.Channels;
using ExternalServiceAPI.Models.Domains;
#endregion

namespace ExternalServiceAPI.Controllers
{
	#region [ SystemStatusController ]
	[RoutePrefix("api/SystemStatus")]
	public class SystemStatusController : ApiController
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

		// GET: api/SystemStatus/{id}

		#region [ Get(string id) ]
		public HttpResponseMessage Get(string id = "TestConnection")
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "SystemStatusController.Get()";
			string strLogAtividade = NOME_DESTA_ROTINA + " (origem: " + (GetIp() ?? "") + ")";
			string strRetorno;
			string strId;
			HttpResponseMessage resultado;
			#endregion

			try
			{
				strId = id.ToLower();
				strLogAtividade += ": id=" + strId;

				switch (strId)
				{
					case "testconnection":
						try
						{
							if (!string.IsNullOrEmpty(strRetorno = BD.DataHora()))
							{
								strRetorno = Models.Domains.Global.formataDataDdMmYyyyHhMmSsComSeparador(Convert.ToDateTime(strRetorno));
							}
							else
							{
								strRetorno = "";
							}

							resultado = new HttpResponseMessage(HttpStatusCode.OK);
							resultado.Content = new StringContent(strRetorno, Encoding.UTF8, "text/plain");
						}
						catch (Exception ex)
						{
							strRetorno = ex.ToString();
							resultado = new HttpResponseMessage(HttpStatusCode.InternalServerError);
							resultado.Content = new StringContent(strRetorno, Encoding.UTF8, "text/plain");
						}

						return resultado;

					case "versao":
						return new HttpResponseMessage
						{
							StatusCode = HttpStatusCode.OK,
							Content = new StringContent(
											Models.Domains.Global.Cte.Aplicativo.M_ID,
											Encoding.UTF8,
											"text/plain"
											)
						};
					default:
						return new HttpResponseMessage
						{
							StatusCode = HttpStatusCode.OK,
							Content = new StringContent(
											"Comando inválido",
											Encoding.UTF8,
											"text/plain"
											)
						};
				}
			}
			catch (Exception ex)
			{
				strLogAtividade += " - Exception: " + ex.Message;

				return new HttpResponseMessage
				{
					StatusCode = HttpStatusCode.InternalServerError,
					Content = new StringContent(ex.ToString(), Encoding.UTF8, "text/plain")
				};
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