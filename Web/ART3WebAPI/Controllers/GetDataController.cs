using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Text;
using System.Web.Script.Serialization;
using ART3WebAPI.Models.Domains;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Repository;
using System.Threading;

namespace ART3WebAPI.Controllers
{
	public class GetDataController : ApiController
	{
		#region [ Produto ]
		[HttpGet]
		public HttpResponseMessage Produto(string codFabricante, string codProduto, string usuario, string sessionToken)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GetDataController.Produto()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;
			string msg_erro;
			Produto produto;
			Usuario usuarioBD;
			HttpResponseMessage result;
			#endregion

			try
			{
				msg = NOME_DESTA_ROTINA + ": Requisição recebida (usuario=" + (usuario ?? "") + ", sessionToken=" + (sessionToken ?? "") + ", codFabricante=" + (codFabricante ?? "") + ", codProduto=" + (codProduto ?? "") + ")";
				Global.gravaLogAtividade(httpRequestId, msg);

				if ((usuario ?? "").Trim().Length == 0)
				{
					throw new Exception("Não foi informada a identificação do usuário!");
				}

				if ((sessionToken ?? "").Trim().Length == 0)
				{
					throw new Exception("Não foi informado o token da sessão do usuário!");
				}

				usuarioBD = GeralDAO.getUsuario(usuario, out msg_erro);
				if (usuarioBD == null)
				{
					throw new Exception("Falha ao tentar validar usuário!");
				}

				if ((!usuarioBD.SessionTokenModuloCentral.Equals(sessionToken)) && (!usuarioBD.SessionTokenModuloLoja.Equals(sessionToken)))
				{
					throw new Exception("Token de sessão inválido!");
				}

				if ((codFabricante ?? "").Length == 0)
				{
					throw new Exception("Código do fabricante é inválido!");
				}

				if ((codProduto ?? "").Length == 0)
				{
					throw new Exception("Código do produto é inválido!");
				}

				produto = ProdutoDAO.getProduto(Global.normalizaCodigoFabricante(codFabricante), Global.normalizaCodigoProduto(codProduto), out msg_erro);

				if (produto == null)
				{
					throw new Exception("Não foi encontrado o registro do produto '" + codProduto.Trim() + "' (fabricante: " + codFabricante + ")");
				}

				#region [ Converte objeto em dados JSON ]
				var serializer = new JavaScriptSerializer();
				var serializedResult = serializer.Serialize(produto);
				#endregion

				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(serializedResult, Encoding.UTF8, "text/html");

				msg = NOME_DESTA_ROTINA + ": Status=" + result.StatusCode.ToString();
				Global.gravaLogAtividade(httpRequestId, msg);

				return result;
			}
			catch (Exception ex)
			{
				msg = NOME_DESTA_ROTINA + ": Exception = " + ex.ToString();
				Global.gravaLogAtividade(httpRequestId, msg);

				return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex);
			}
		}
		#endregion

		#region [ ProdutoBySku ]
		[HttpGet]
		public HttpResponseMessage ProdutoBySku(string codProduto, string usuario, string sessionToken)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GetDataController.ProdutoBySku()";
			Guid httpRequestId = Request.GetCorrelationId();
			string msg;
			string msg_erro;
			Produto produto;
			Usuario usuarioBD;
			HttpResponseMessage result;
			#endregion

			try
			{
				msg = NOME_DESTA_ROTINA + ": Requisição recebida (usuario=" + (usuario ?? "") + ", sessionToken=" + (sessionToken ?? "") + ", codProduto=" + (codProduto ?? "") + ")";
				Global.gravaLogAtividade(httpRequestId, msg);

				if ((usuario ?? "").Trim().Length == 0)
				{
					throw new Exception("Não foi informada a identificação do usuário!");
				}

				if ((sessionToken ?? "").Trim().Length == 0)
				{
					throw new Exception("Não foi informado o token da sessão do usuário!");
				}

				usuarioBD = GeralDAO.getUsuario(usuario, out msg_erro);
				if (usuarioBD == null)
				{
					throw new Exception("Falha ao tentar validar usuário!");
				}

				if ((!usuarioBD.SessionTokenModuloCentral.Equals(sessionToken)) && (!usuarioBD.SessionTokenModuloLoja.Equals(sessionToken)))
				{
					throw new Exception("Token de sessão inválido!");
				}

				if ((codProduto ?? "").Length == 0)
				{
					throw new Exception("Código do produto é inválido!");
				}

				produto = ProdutoDAO.getProdutoBySku(Global.normalizaCodigoProduto(codProduto), out msg_erro);

				if (produto == null)
				{
					throw new Exception("Não foi encontrado o registro do produto '" + codProduto.Trim() + "'");
				}

				#region [ Converte objeto em dados JSON ]
				var serializer = new JavaScriptSerializer();
				var serializedResult = serializer.Serialize(produto);
				#endregion

				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(serializedResult, Encoding.UTF8, "text/html");

				msg = NOME_DESTA_ROTINA + ": Status=" + result.StatusCode.ToString();
				Global.gravaLogAtividade(httpRequestId, msg);

				return result;
			}
			catch (Exception ex)
			{
				msg = NOME_DESTA_ROTINA + ": Exception = " + ex.ToString();
				Global.gravaLogAtividade(httpRequestId, msg);

				return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex);
			}
		}
		#endregion

		#region [ PageContentViaHttpGet ]
		[HttpGet]
		public HttpResponseMessage PageContentViaHttpGet(string usuario, string sessionToken)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GetDataController.PageContentViaHttpGet()";
			const int MAX_LOG_ATIVIDADE_PAGE_CONTENT_SIZE_DEFAULT = 1024;
			Guid httpRequestId = Request.GetCorrelationId();
			System.Net.Http.Headers.HttpRequestHeaders headers;
			RegistroTabelaParametro paramMaxLogAtividadePageContentSize;
			int maxLogAtividadePageContentSize;
			string sHeaderId;
			string urlGet = "";
			string sPageContent;
			string msg;
			string msg_erro;
			Usuario usuarioBD;
			HttpResponseMessage result;
			#endregion

			try
			{
				msg = NOME_DESTA_ROTINA + ": Requisição recebida (usuario=" + (usuario ?? "") + ", sessionToken=" + (sessionToken ?? "") + ")";
				Global.gravaLogAtividade(httpRequestId, msg);

				if ((usuario ?? "").Trim().Length == 0)
				{
					throw new Exception("Não foi informada a identificação do usuário!");
				}

				if ((sessionToken ?? "").Trim().Length == 0)
				{
					throw new Exception("Não foi informado o token da sessão do usuário!");
				}

				usuarioBD = GeralDAO.getUsuario(usuario, out msg_erro);
				if (usuarioBD == null)
				{
					throw new Exception("Falha ao tentar validar usuário!");
				}

				if ((!usuarioBD.SessionTokenModuloCentral.Equals(sessionToken)) && (!usuarioBD.SessionTokenModuloLoja.Equals(sessionToken)))
				{
					throw new Exception("Token de sessão inválido!");
				}

				headers = Request.Headers;
				sHeaderId = "X-Query-Url-Get";
				if (headers.Contains(sHeaderId))
				{
					urlGet = headers.GetValues(sHeaderId).First();
				}

				if ((urlGet ?? "").Length == 0)
				{
					throw new Exception("Não foi informada a URL da página a ser consultada!");
				}

				#region [ Obtém o parâmetro que define o tamanho máximo p/ o log de atividades registrar o conteúdo da página ]
				paramMaxLogAtividadePageContentSize = GeralDAO.getRegistroTabelaParametro(Global.Cte.Parametros.ID_T_PARAMETRO.SSW_RASTREAMENTO_VIA_WEBAPI_MAX_LOG_ATIVIDADE_PAGE_CONTENT_SIZE);
				if (paramMaxLogAtividadePageContentSize == null)
				{
					maxLogAtividadePageContentSize = MAX_LOG_ATIVIDADE_PAGE_CONTENT_SIZE_DEFAULT;
				}
				else
				{
					maxLogAtividadePageContentSize = paramMaxLogAtividadePageContentSize.campo_inteiro;
				}
				#endregion

				if (!enviaRequisicaoGetComRetry(httpRequestId, urlGet, maxLogAtividadePageContentSize, out sPageContent, out msg_erro))
				{
					sPageContent = "Falha ao tentar realizar a consulta!";
				}

				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(sPageContent, Encoding.UTF8, "text/html");

				msg = NOME_DESTA_ROTINA + ": Status=" + result.StatusCode.ToString();
				Global.gravaLogAtividade(httpRequestId, msg);

				return result;
			}
			catch (Exception ex)
			{
				msg = NOME_DESTA_ROTINA + ": Exception = " + ex.ToString();
				Global.gravaLogAtividade(httpRequestId, msg);

				return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex);
			}
		}
		#endregion

		#region [ enviaRequisicaoGetComRetry ]
		private bool enviaRequisicaoGetComRetry(Guid? httpRequestId, string urlGet, int maxLogAtividadeSize, out string pageContent, out string msg_erro)
		{
			#region [ Declarações ]
			const int MAX_TENTATIVAS = 3;
			int qtdeTentativasRealizadas = 0;
			bool blnResposta;
			#endregion

			do
			{
				qtdeTentativasRealizadas++;

				blnResposta = enviaRequisicaoGet(httpRequestId, urlGet, maxLogAtividadeSize, out pageContent, out msg_erro);
				if (blnResposta) break;

				Thread.Sleep(1000);
			} while (qtdeTentativasRealizadas < MAX_TENTATIVAS);

			return blnResposta;
		}
		#endregion

		#region [ enviaRequisicaoGet ]
		private bool enviaRequisicaoGet(Guid? httpRequestId, string urlGet, int maxLogAtividadeSize, out string pageContent, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "GetDataController.enviaRequisicaoGet()";
			string strMsg;
			bool blnSucesso;
			#endregion

			pageContent = "";
			msg_erro = "";

			try
			{
				strMsg = NOME_DESTA_ROTINA + " - TX\n" + urlGet;
				Global.gravaLogAtividade(httpRequestId, strMsg);

				using (HttpClient client = new HttpClient())
				{
					client.BaseAddress = new Uri(urlGet);
					HttpResponseMessage response = client.GetAsync("").Result;  // Blocking call! Program will wait here until a response is received or a timeout occurs.
					if (response.IsSuccessStatusCode)
					{
						var resp = response.Content.ReadAsStringAsync();
						pageContent = resp.Result;
						blnSucesso = true;
					}
					else
					{
						blnSucesso = false;
					}
				}

				strMsg = NOME_DESTA_ROTINA + " - RX\nSucesso: " + blnSucesso.ToString();
				if (blnSucesso)
				{
					if ((maxLogAtividadeSize <= 0) || (pageContent.Length <= maxLogAtividadeSize))
					{
						strMsg += "\n" + pageContent;
					}
					else
					{
						strMsg += "\n" + pageContent.Substring(0, maxLogAtividadeSize) + " ... (truncated)";
					}
				}
				Global.gravaLogAtividade(httpRequestId, strMsg);

				return blnSucesso;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				Global.gravaLogAtividade(httpRequestId, NOME_DESTA_ROTINA + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion
	}
}
