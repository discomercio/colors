using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.ServiceModel.Channels;
using System.Text;
using System.Web;
using System.Web.Http;
using WebHookV2.Models;
using WebHookV2.Models.Domains;
using WebHookV2.Models.Entities;
using WebHookV2.Models.Repository;

namespace WebHookV2.Controllers
{
	#region [ BraspagController ]
	#endregion
	public class BraspagController : ApiController
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
			const string NOME_DESTA_ROTINA = "BraspagController.Teste()";
			string strLogAtividade = NOME_DESTA_ROTINA + " (IP origem: " + (GetIp() ?? "") + ")";
			HttpResponseMessage result;
			#endregion

			try
			{
				result = Request.CreateResponse<string>(HttpStatusCode.OK, "Versão: " + Global.Cte.Versao.M_ID);
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

		#region [ Post api/Braspag ]
		public HttpResponseMessage Post([FromBody] JToken postData, HttpRequestMessage request)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagController.Post()";
			int intLinhasAfetadas = 0;
			DateTime dtHrInicio = DateTime.Now;
			string strLogAtividade = NOME_DESTA_ROTINA + " (IP origem: " + (GetIp() ?? "") + ")";
			string strMessage;
			HttpResponseMessage result = null;
			BraspagPostNotificacao braspagPost;
			BD bd;
			#endregion

			try //try-catch-finally
			{
				#region [ Log inicial ]
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Post recebido\n" + postData.ToString());
				#endregion

				#region [ Há parâmetros? ]
				if (postData == null)
				{
					strMessage = "Parâmetros inválidos";
					strLogAtividade += ", Resposta: '" + strMessage + "'";
					result = new HttpResponseMessage(HttpStatusCode.BadRequest);
					result.Content = new StringContent(strMessage, Encoding.UTF8, "text/plain");
					return result;
				}
				#endregion

				#region [ Desserializa parâmetros ]
				braspagPost = JsonConvert.DeserializeObject<BraspagPostNotificacao>(postData.ToString());
				if (braspagPost == null)
				{
					strMessage = "Parâmetros incorretos";
					strLogAtividade += ", Resposta: '" + strMessage + "'";
					result = new HttpResponseMessage(HttpStatusCode.BadRequest);
					result.Content = new StringContent(strMessage, Encoding.UTF8, "text/plain");
					return result;
				}
				#endregion

				#region [ Consistência dos parâmetros ]

				#region [ Consiste parâmetro obrigatório 'PaymentId' ]
				if (string.IsNullOrEmpty(braspagPost.PaymentId))
				{
					strMessage = "Identificador obrigatório não informado";
					strLogAtividade += ", Resposta: '" + strMessage + "'";
					result = new HttpResponseMessage(HttpStatusCode.BadRequest);
					result.Content = new StringContent(strMessage, Encoding.UTF8, "text/plain");
					return result;
				}

				if (braspagPost.PaymentId.Trim().Length != 36)
				{
					strMessage = "Identificador obrigatório em formato inválido";
					strLogAtividade += ", Resposta: '" + strMessage + "'";
					result = new HttpResponseMessage(HttpStatusCode.BadRequest);
					result.Content = new StringContent(strMessage, Encoding.UTF8, "text/plain");
					return result;
				}
				#endregion

				#region [ Consiste parâmetro opcional 'RecurrentPaymentId' ]
				if (!string.IsNullOrEmpty(braspagPost.RecurrentPaymentId))
				{
					if (braspagPost.RecurrentPaymentId.Trim().Length != 36)
					{
						strMessage = "Identificador opcional em formato inválido";
						strLogAtividade += ", Resposta: '" + strMessage + "'";
						result = new HttpResponseMessage(HttpStatusCode.BadRequest);
						result.Content = new StringContent(strMessage, Encoding.UTF8, "text/plain");
						return result;
					}
				}
				#endregion

				#region [ Consiste parâmetro obrigatório 'ChangeType' ]
				if (braspagPost.ChangeType == 0)
				{
					strMessage = "Código de notificação inválido";
					strLogAtividade += ", Resposta: '" + strMessage + "'";
					result = new HttpResponseMessage(HttpStatusCode.BadRequest);
					result.Content = new StringContent(strMessage, Encoding.UTF8, "text/plain");
					return result;
				}
				#endregion

				#endregion

				strLogAtividade += ": " + braspagPost.FormataDados(showOnePropertyPerLine: false, inlinePropertySeparator: ", ");

				#region [ Grava no BD ]
				try
				{
					bd = new BD();
					intLinhasAfetadas = bd.insereBraspagWebHookV2(Global.Parametros.Braspag.BraspagWebHookV2Empresa, braspagPost);

					if (intLinhasAfetadas < 1)
					{
						#region [ Retorno para falha ]
						strLogAtividade += ", Falha na gravação dos dados: nenhum registro inserido no BD";
						result = new HttpResponseMessage(HttpStatusCode.InternalServerError);
						result.Content = new StringContent("", Encoding.UTF8, "text/plain");
						return result;
						#endregion
					}
					else
					{
						#region [ Retorno para sucesso ]
						result = Request.CreateResponse<string>(HttpStatusCode.OK, "OK");
						return result;
						#endregion
					}
				}
				catch (Exception ex)
				{
					strLogAtividade += " - Exception: " + ex.Message;
					result = new HttpResponseMessage(HttpStatusCode.InternalServerError);
					result.Content = new StringContent("", Encoding.UTF8, "text/plain");
					return result;
				}
				#endregion
			}
			catch (Exception ex)
			{
				strLogAtividade += " - Exception: " + ex.Message;
				result = new HttpResponseMessage(HttpStatusCode.InternalServerError);
				result.Content = new StringContent("", Encoding.UTF8, "text/plain");
				return result;
			}
			finally
			{
				if (result != null) strLogAtividade += ", Duração: " + Global.formataDuracaoHMSMs(DateTime.Now - dtHrInicio) + ", Result: " + ((int)result.StatusCode).ToString() + " - " + result.StatusCode.ToString();
				Global.gravaLogAtividade(strLogAtividade);
			}
		}
		#endregion
	}
}