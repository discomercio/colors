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

namespace ART3WebAPI.Controllers
{
	public class GetDataController : ApiController
	{
		#region [ Produto ]
		[HttpGet]
		public HttpResponseMessage Produto(string codFabricante, string codProduto, string usuario, string sessionToken)
		{
			#region [ Declarações ]
			string msg_erro;
			Produto produto;
			Usuario usuarioBD;
			HttpResponseMessage result;
			#endregion

			try
			{
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
				return result;
			}
			catch (Exception ex)
			{
				return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex);
			}
		}
		#endregion

		#region [ ProdutoBySku ]
		[HttpGet]
		public HttpResponseMessage ProdutoBySku(string codProduto, string usuario, string sessionToken)
		{
			#region [ Declarações ]
			string msg_erro;
			Produto produto;
			Usuario usuarioBD;
			HttpResponseMessage result;
			#endregion

			try
			{
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
				return result;
			}
			catch (Exception ex)
			{
				return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex);
			}
		}
		#endregion
	}
}
