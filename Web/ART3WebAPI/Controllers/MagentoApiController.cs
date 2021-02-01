using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Repository;
using ART3WebAPI.Models.Domains;
using System.Threading.Tasks;
using System.Text;

namespace ART3WebAPI.Controllers
{
	public class MagentoApiController : ApiController
	{
		#region [ Teste ]
		[HttpGet]
		public HttpResponseMessage Teste()
		{
			HttpResponseMessage result = Request.CreateResponse<string>(HttpStatusCode.OK, "Versão: " + Global.Cte.Versao.M_ID);
			return result;
		}
		#endregion

		#region [ GetPedido ]
		/// <summary>
		/// Obtém os dados do pedido Magento em formato JSON
		/// </summary>
		/// <param name="numeroPedidoMagento">Número do pedido Magento</param>
		/// <param name="operationControlTicket">
		///		Identificador da operação no front-end (formato GUID).
		///		O objetivo deste identificador é evitar a repetição de requisições via API do Magento dentro de uma mesma operação no front-end.
		///		Os dados consultados no Magento são armazenados no BD para agilizar consultas posteriores.
		///		Caso o pedido armazenado no BD esteja com outro valor de 'operationControlTicket', a consulta via API é realizada para assegurar que os dados estão atualizados.
		/// </param>
		/// <param name="loja">
		///		Número da loja do usuário
		///		O número da loja define a URL do web service da API do Magento
		///	</param>
		/// <param name="usuario">Identificação do usuário</param>
		/// <param name="sessionToken">Token da sessão do usuário: é usado para assegurar que a consulta está sendo realizada por um usuário autenticado</param>
		/// <returns>Retorna os dados do pedido Magento especificado em formato JSON</returns>
		[HttpGet]
		public HttpResponseMessage GetPedido(string numeroPedidoMagento, string operationControlTicket, string loja, string usuario, string sessionToken)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "MagentoApiController.GetPedido()";
			string msg;
			string msg_erro;
			string serializedResult;
			Usuario usuarioBD;
			MagentoApiLoginParameters loginParameters;
			MagentoErpSalesOrder salesOrder = new MagentoErpSalesOrder();
			HttpResponseMessage result;
			#endregion

			#region [ Inicialização ]
			salesOrder.numeroPedidoMagento = numeroPedidoMagento;
			#endregion

			#region [ Log atividade ]
			msg = "MagentoApi.GetPedido() - numeroPedidoMagento = " + numeroPedidoMagento + ", operationControlTicket = " + operationControlTicket + ", loja = " + loja + ", usuario = " + usuario + ", sessionToken = " + sessionToken;
			Global.gravaLogAtividade(msg);
			#endregion

			#region [ Validação de segurança: session token confere? ]
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
			#endregion

			#region [ Obtém parâmetros da API no cadastro da loja ]
			if ((loja ?? "").Trim().Length == 0)
			{
				throw new Exception("O número da loja não foi informado!");
			}

			loginParameters = MagentoApiDAO.getLoginParameters(loja, out msg_erro);
			if (loginParameters == null)
			{
				msg = "Falha ao tentar recuperar os parâmetros de login da API do Magento para a loja " + loja + "!";
				if (msg_erro.Length > 0) msg += "\n" + msg_erro;
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + msg);
				throw new Exception(msg);
			}
			#endregion

			if (loginParameters.api_versao == Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V1_SOAP_XML)
			{
				#region [ Tratamento para API SOAP (XML) do Magento 1.8 ]
				serializedResult = MagentoSoapApi.processaGetPedido(numeroPedidoMagento, operationControlTicket, loja, usuario, sessionToken, loginParameters);
				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(serializedResult, Encoding.UTF8, "text/html");
				#endregion
			}
			else
			{
				#region [ Tratamento para API REST (JSON) do Magento 2 ]
				serializedResult = Magento2RestApi.processaGetPedido(numeroPedidoMagento, operationControlTicket, loja, usuario, sessionToken, loginParameters);
				result = Request.CreateResponse(HttpStatusCode.OK);
				result.Content = new StringContent(serializedResult, Encoding.UTF8, "text/html");
				#endregion
			}

			return result;
		}
		#endregion
	}
}
