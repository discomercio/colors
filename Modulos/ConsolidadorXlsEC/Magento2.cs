using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ConsolidadorXlsEC
{
	#region [ Magento2 ]
	class Magento2
	{
		const int MAX_SIZE_MENSAGEM_LOG_ATIVIDADE = 1 * 1024 * 1024;

		#region[ ReaderWriterLock ]
		// Para garantir que os acessos à API do Magento sejam thread-safe
		public static ReaderWriterLock rwlMagento2RestApi = new ReaderWriterLock();
		#endregion

		#region [ enviaRequisicaoGetComRetry ]
		/// <summary>
		/// Método que executa o enviaRequisicao() dentro de um laço de tentativas até que a execução seja bem sucedida ou a quantidade máxima de tentativas seja atingida.
		/// </summary>
		/// <param name="urlParameters"></param>
		/// <param name="accessToken"></param>
		/// <param name="urlBaseAddress"></param>
		/// <param name="respRest"></param>
		/// <param name="msg_erro"></param>
		/// <returns></returns>
		public static bool enviaRequisicaoGetComRetry(string urlParameters, string accessToken, string urlBaseAddress, out string respRest, out string msg_erro)
		{
			#region [ Declarações ]
			const int MAX_TENTATIVAS = 5;
			int qtdeTentativasRealizadas = 0;
			bool blnResposta;
			#endregion

			do
			{
				qtdeTentativasRealizadas++;

				blnResposta = enviaRequisicaoGet(urlParameters, accessToken, urlBaseAddress, out respRest, out msg_erro);
				if (blnResposta) break;

				Thread.Sleep(1000);
			} while (qtdeTentativasRealizadas < MAX_TENTATIVAS);

			return blnResposta;
		}
		#endregion

		#region [ enviaRequisicaoGet ]
		public static bool enviaRequisicaoGet(string urlParameters, string accessToken, string urlBaseAddress, out string respRest, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Magento2.enviaRequisicaoGet()";
			string strMsg;
			bool blnSucesso;
			Magento2HttpErrorResult httpErrorResult;
			#endregion

			respRest = "";
			msg_erro = "";

			try
			{
				strMsg = NOME_DESTA_ROTINA + " - TX\n" + urlBaseAddress;
				if ((urlParameters ?? "").Length > 0) strMsg += urlParameters;
				Global.gravaLogAtividade(strMsg);

				using (HttpClient client = new HttpClient())
				{
					client.BaseAddress = new Uri(urlBaseAddress);

					// Add an Accept header for JSON format.
					client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
					client.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);

					HttpResponseMessage response = client.GetAsync(urlParameters).Result;  // Blocking call! Program will wait here until a response is received or a timeout occurs.
					if (response.IsSuccessStatusCode)
					{
						var resp = response.Content.ReadAsStringAsync();
						respRest = resp.Result;
						blnSucesso = true;
					}
					else
					{
						var resp = response.Content.ReadAsStringAsync();
						respRest = resp.Result;
						httpErrorResult = JsonConvert.DeserializeObject<Magento2HttpErrorResult>(respRest);
						msg_erro = httpErrorResult.message;
						blnSucesso = false;
					}
				}

				strMsg = NOME_DESTA_ROTINA + " - RX\nSucesso: " + blnSucesso.ToString();
				strMsg += "\n" + respRest;
				Global.gravaLogAtividade(strMsg, MAX_SIZE_MENSAGEM_LOG_ATIVIDADE);

				return blnSucesso;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ enviaRequisicaoPost ]
		public static bool enviaRequisicaoPost(string urlParameters, object parameters, string accessToken, string urlBaseAddress, out string respRest, out HttpResponseMessage response, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "enviaRequisicaoPost()";
			string strMsg;
			bool blnSucesso;
			Magento2HttpErrorResult httpErrorResult;
			#endregion

			respRest = "";
			msg_erro = "";
			response = null;

			try
			{
				strMsg = NOME_DESTA_ROTINA + " - TX\n" + urlBaseAddress;
				if ((urlParameters ?? "").Length > 0) strMsg += urlParameters;
				Global.gravaLogAtividade(strMsg);

				using (HttpClient client = new HttpClient())
				{
					client.BaseAddress = new Uri(urlBaseAddress);

					// Add an Accept header for JSON format.
					client.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
					string p = JsonConvert.SerializeObject(parameters);
					var stringContent = new StringContent(p, Encoding.UTF8, "application/json");
					response = client.PostAsync(urlParameters, stringContent).Result;  // Blocking call! Program will wait here until a response is received or a timeout occurs.
					if (response.IsSuccessStatusCode)
					{
						var resp = response.Content.ReadAsStringAsync();
						respRest = resp.Result;
						blnSucesso = true;
					}
					else
					{
						var resp = response.Content.ReadAsStringAsync();
						respRest = resp.Result;
						httpErrorResult = JsonConvert.DeserializeObject<Magento2HttpErrorResult>(respRest);
						msg_erro = httpErrorResult.message;
						blnSucesso = false;
					}
				}

				strMsg = NOME_DESTA_ROTINA + " - RX\nSucesso: " + blnSucesso.ToString();
				strMsg += "\n" + respRest;
				Global.gravaLogAtividade(strMsg);

				return blnSucesso;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ montaRequisicaoGetSalesOrderInfoByIncrementId ]
		public static string montaRequisicaoGetSalesOrderInfoByIncrementId(string orderIncrementId, string urlEndPoint, out string urlBaseAddress)
		{
			string urlParamReqRest;

			urlBaseAddress = (urlEndPoint ?? "").Trim();
			if (!urlBaseAddress.EndsWith("/")) urlBaseAddress += "/";
			urlBaseAddress += "orders/";
			urlParamReqRest = "?searchCriteria[filterGroups][0][filters][0][field]=increment_id&searchCriteria[filterGroups][0][filters][0][value]=" + orderIncrementId + "&searchCriteria[filterGroups][0][filters][0][conditionType]=eq";

			return urlParamReqRest;
		}
		#endregion

		#region [ montaRequisicaoGetSalesOrderInfoByEntityId ]
		public static string montaRequisicaoGetSalesOrderInfoByEntityId(string orderEntityId, string urlEndPoint, out string urlBaseAddress)
		{
			string urlParamReqRest;

			urlBaseAddress = (urlEndPoint ?? "").Trim();
			if (!urlBaseAddress.EndsWith("/")) urlBaseAddress += "/";
			urlBaseAddress += "orders/";
			urlParamReqRest = orderEntityId;

			return urlParamReqRest;
		}
		#endregion

		#region [ montaRequisicaoGetProducts ]
		public static string montaRequisicaoGetProducts(List<Magento2SearchCriteriaFilterGroups> filtros, string urlEndPoint, out string urlBaseAddress)
		{
			int idxFilterGroups = 0;
			int idxFilters;
			string urlParamReqRest;
			string sFiltro;
			bool bHaFiltro = false;
			StringBuilder sbFiltros = new StringBuilder("");

			urlBaseAddress = (urlEndPoint ?? "").Trim();
			if (!urlBaseAddress.EndsWith("/")) urlBaseAddress += "/";
			urlBaseAddress += "products/";

			if (filtros != null)
			{
				if (filtros.Count > 0)
				{
					bHaFiltro = true;
				}
			}

			if (!bHaFiltro)
			{
				urlParamReqRest = "?searchCriteria[pageSize]=0";
			}
			else
			{
				// Lógica de operação dos filtros:
				// https://devdocs.magento.com/guides/v2.4/rest/performing-searches.html
				// The filter_groups array defines one or more filters. Each filter defines a search term, and the field, value, and condition_type of a search term must be assigned the same index number, starting with 0. Increment additional terms as needed.
				// When constructing a search, keep the following in mind:
				//		To perform a logical OR, specify multiple filters within a filter_groups.
				//		To perform a logical AND, specify multiple filter_groups.
				//		You cannot perform a logical OR across different filter_groups, such as (A AND B) OR (X AND Y). ORs can be performed only within the context of a single filter_groups.
				//		You can only search top-level attributes.
				foreach (Magento2SearchCriteriaFilterGroups filterGroup in filtros)
				{
					idxFilters = 0;
					foreach (Magento2SearchCriteriaFilterGroupsFilters filter in filterGroup.filters)
					{
						sFiltro = "searchCriteria[filterGroups][" + idxFilterGroups.ToString() + "][filters][" + idxFilters.ToString() + "][field]=" + filter.field +
									"&searchCriteria[filterGroups][" + idxFilterGroups.ToString() + "][filters][" + idxFilters.ToString() + "][value]=" + filter.value +
									"&searchCriteria[filterGroups][" + idxFilterGroups.ToString() + "][filters][" + idxFilters.ToString() + "][conditionType]=" + filter.condition_type;

						if (sbFiltros.Length > 0)
						{
							sbFiltros.Append("&" + sFiltro);
						}
						else
						{
							sbFiltros.Append("?" + sFiltro);
						}

						idxFilters++;
					}

					idxFilterGroups++;
				}

				urlParamReqRest = sbFiltros.ToString();
			}

			return urlParamReqRest;
		}
		#endregion

		#region [ montaRequisicaoPostSalesOrderAddComment ]
		public static string montaRequisicaoPostSalesOrderAddComment(string orderId, string urlEndPoint, out string urlBaseAddress)
		{
			string urlParamReqRest;

			urlBaseAddress = (urlEndPoint ?? "").Trim();
			if (!urlBaseAddress.EndsWith("/")) urlBaseAddress += "/";
			urlBaseAddress += "orders/";
			urlParamReqRest = Global.retiraZerosAEsquerda(orderId) + "/comments";

			return urlParamReqRest;
		}
		#endregion

		#region [ getSalesOrderInfo ]
		public static Magento2SalesOrderInfo getSalesOrderInfo(string numeroPedidoMagento, Loja lojaLoginParameters, out string jsonResponse, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Magento2.getSalesOrderInfo()";
			bool blnEnviouOk;
			string msg;
			string urlParamReqRest;
			string urlBaseAddress = "";
			string respJson;
			string msg_erro_aux;
			Magento2SalesOrderSearchResponse salesOrderInfoSearchResponse;
			Magento2SalesOrderInfo salesOrderInfo = null;
			#endregion

			jsonResponse = "";
			msg_erro = "";
			try
			{
				#region [ Consistências ]
				if ((numeroPedidoMagento ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o número do pedido Magento!";
					return null;
				}

				if ((lojaLoginParameters.magento_api_rest_endpoint ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o endpoint da API REST do Magento!";
					return null;
				}

				if ((lojaLoginParameters.magento_api_rest_access_token ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o access token para acessar a API REST do Magento!";
					return null;
				}
				#endregion

				try
				{
					rwlMagento2RestApi.AcquireWriterLock(Global.Cte.Magento2RestApi.TIMEOUT_READER_WRITER_LOCK_EM_MS);
					try // FINALLY: rwlMagento2RestApi.ReleaseWriterLock();
					{
						// IMPORTANTE: a requisição [GET] orders/{id} recebe como parâmetro o nº pedido referente ao campo 'entity_id', que é o ID usado internamente pelo Magento e não é exibido no painel Admin.
						// =========== Portanto, para se realizar a consulta através do nº pedido que é exibido no painel Admin, ou seja, o campo 'increment_id', é usada a requisição [GET] orders combinada com
						// critérios de pesquisa de forma a consultar pelo campo 'increment_id'.
						// Entretanto, essa requisição retorna um array já que dependendo dos critérios definidos pode haver mais de um pedido na resposta. Mas o principal é que sua documentação
						// alerta para o fato de que as informações detalhadas podem não estar inclusas na resposta:
						//		Lists orders that match specified search criteria. This call returns an array of objects, but detailed information about each object’s attributes might not be included.
						//		See https://devdocs.magento.com/codelinks/attributes.html#OrderRepositoryInterface to determine which call to use to get detailed information about all attributes for an object.
						//
						// Para assegurar que o conjunto completo de dados é obtido, adota-se a seguinte lógica:
						//		1) Consulta através de [GET] orders para localizar o pedido através do campo 'increment_id'
						//		2) A partir da resposta obtida na primeira consulta, é obtido o valor do campo 'entity_id' e, em seguida, é feita a consulta [GET] orders/{id}
						//
						// OBSERVAÇÃO: nos testes realizados, a resposta da requisição [GET] orders usando search criteria retornou todas as informações detalhadas do pedido na resposta, tanto quando
						// havia somente um pedido na resposta quanto quando havia vários. Essa verificação foi feita através da comparação dos campos retornados pela consulta [GET] orders/{id}

						urlParamReqRest = Magento2.montaRequisicaoGetSalesOrderInfoByIncrementId(numeroPedidoMagento, lojaLoginParameters.magento_api_rest_endpoint, out urlBaseAddress);
						blnEnviouOk = Magento2.enviaRequisicaoGetComRetry(urlParamReqRest, lojaLoginParameters.magento_api_rest_access_token, urlBaseAddress, out respJson, out msg_erro_aux);
						if (!blnEnviouOk)
						{
							msg_erro = "Falha ao tentar consultar o pedido Magento " + numeroPedidoMagento + " através da API REST!";
							if (msg_erro_aux.Length > 0) msg_erro += "\n" + msg_erro_aux;
							return null;
						}

						jsonResponse = respJson;

						salesOrderInfoSearchResponse = Magento2.decodificaJsonSalesOrderSearchResponse(respJson, out msg_erro_aux);
						if (salesOrderInfoSearchResponse == null)
						{
							msg_erro = "Falha ao tentar decodificar os dados da resposta do pedido Magento " + numeroPedidoMagento + " obtidos através da API REST!";
							if (msg_erro_aux.Length > 0) msg_erro += "\n" + msg_erro_aux;
							return null;
						}

						if (salesOrderInfoSearchResponse.total_count == 0)
						{
							msg_erro = "Pedido Magento " + numeroPedidoMagento + " não foi encontrado!";
							return null;
						}

						salesOrderInfo = salesOrderInfoSearchResponse.items[0];
						if (salesOrderInfo == null)
						{
							msg_erro = "Falha ao tentar decodificar os dados do pedido Magento " + numeroPedidoMagento + " (API REST)";
							return null;
						}

						if (lojaLoginParameters.magento_api_rest_force_get_sales_order_by_entity_id == 0)
						{
							return salesOrderInfo;
						}
						else
						{
							urlParamReqRest = Magento2.montaRequisicaoGetSalesOrderInfoByEntityId(salesOrderInfo.entity_id, lojaLoginParameters.magento_api_rest_endpoint, out urlBaseAddress);
							blnEnviouOk = Magento2.enviaRequisicaoGetComRetry(urlParamReqRest, lojaLoginParameters.magento_api_rest_access_token, urlBaseAddress, out respJson, out msg_erro_aux);

							if (!blnEnviouOk)
							{
								msg_erro = "Falha ao tentar consultar o pedido Magento " + numeroPedidoMagento + " via entity_id através da API REST!";
								if (msg_erro_aux.Length > 0) msg_erro += "\n" + msg_erro_aux;
								return null;
							}

							jsonResponse = respJson;

							salesOrderInfo = Magento2.decodificaJsonSalesOrderInfo(respJson, out msg_erro_aux);
							if (salesOrderInfo == null)
							{
								msg_erro = "Falha ao tentar decodificar os dados da resposta do pedido Magento " + numeroPedidoMagento + " obtidos através da consulta via entity_id da API REST!";
								if (msg_erro_aux.Length > 0) msg_erro += "\n" + msg_erro_aux;
								return null;
							}

							return salesOrderInfo;
						}
					}
					finally
					{
						rwlMagento2RestApi.ReleaseWriterLock();
					}
				}
				catch (Exception ex)
				{
					// Tratamento para exception gerada no timeout do AcquireWriterLock
					msg = NOME_DESTA_ROTINA + " - Exception: " + ex.ToString();
					Global.gravaLogAtividade(msg);
					return null;
				}
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + ex.ToString());
				return null;
			}
		}
		#endregion

		#region [ decodificaJsonSalesOrderSearchResponse ]
		public static Magento2SalesOrderSearchResponse decodificaJsonSalesOrderSearchResponse(string respJson, out string msg_erro)
		{
			Magento2SalesOrderSearchResponse ret;

			msg_erro = "";

			try
			{
				if ((respJson ?? "").Trim().Length == 0) return null;

				ret = JsonConvert.DeserializeObject<Magento2SalesOrderSearchResponse>(respJson);

				return ret;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				return null;
			}
		}
		#endregion

		#region [ decodificaJsonSalesOrderInfo ]
		public static Magento2SalesOrderInfo decodificaJsonSalesOrderInfo(string respJson, out string msg_erro)
		{
			Magento2SalesOrderInfo ret;

			msg_erro = "";

			try
			{
				if ((respJson ?? "").Trim().Length == 0) return null;

				ret = JsonConvert.DeserializeObject<Magento2SalesOrderInfo>(respJson);

				return ret;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				return null;
			}
		}
		#endregion

		#region [ decodificaSalesOrderInfoMage2ParaMage1 ]
		/// <summary>
		/// Decodifica os dados obtidos através da API REST/JSON e transfere para um objeto da classe MagentoSoapApiSalesOrderInfo que continuará sendo usada
		/// para minimizar os ajustes necessários no restante do sistema.
		/// </summary>
		/// <param name="mage2SalesOrderInfo">Conteúdo da resposta recebida via API do Magento 2</param>
		/// <param name="msg_erro">Retorna a mensagem do erro ocorrido no processamento, se houver</param>
		/// <returns>Retorna objeto com os dados do pedido na estrutura do Magento 1</returns>
		public static SalesOrderInfo decodificaSalesOrderInfoMage2ParaMage1(Magento2SalesOrderInfo mage2SalesOrderInfo, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Magento2.decodificaSalesOrderInfoMage2ParaMage1()";
			StringBuilder sbEndereco;
			SalesOrderInfo mage1SalesOrderInfo = new SalesOrderInfo();
			Magento2ExtensionAttributesShippingAssignmentsShipping mage2Shipping;
			Magento2ExtensionAttributesShippingAssignmentsShippingAddress mage2ShippingAddress;
			Magento2ExtensionAttributesSkyhubInfo mage2SkyHubInfo;
			SalesOrderItem mage1SalesOrderItem;
			StatusHistory mage1StatusHistory;
			#endregion

			msg_erro = "";

			try
			{
				if (mage2SalesOrderInfo == null)
				{
					msg_erro = NOME_DESTA_ROTINA + ": objeto com os dados do pedido é inválido!";
					return null;
				}

				mage2Shipping = retornaShipping(mage2SalesOrderInfo);
				mage2ShippingAddress = retornaShippingAddress(mage2SalesOrderInfo);
				mage2SkyHubInfo = retornaSkyHubInfo(mage2SalesOrderInfo);

				#region [ Campos na raiz do corpo do pedido ]
				mage1SalesOrderInfo.increment_id = mage2SalesOrderInfo.increment_id;
				// *** parent_id
				mage1SalesOrderInfo.store_id = mage2SalesOrderInfo.store_id;
				mage1SalesOrderInfo.created_at = mage2SalesOrderInfo.created_at;
				mage1SalesOrderInfo.updated_at = mage2SalesOrderInfo.updated_at;
				// *** is_active
				mage1SalesOrderInfo.customer_id = mage2SalesOrderInfo.customer_id;
				mage1SalesOrderInfo.tax_amount = mage2SalesOrderInfo.tax_amount;
				// *** tax_canceled
				// *** tax_invoiced
				// *** tax_refunded
				mage1SalesOrderInfo.shipping_amount = mage2SalesOrderInfo.shipping_amount;
				// *** shipping_canceled
				// *** shipping_invoiced
				// *** shipping_refunded
				mage1SalesOrderInfo.shipping_tax_amount = mage2SalesOrderInfo.shipping_tax_amount;
				// *** shipping_tax_refunded
				mage1SalesOrderInfo.shipping_discount_amount = mage2SalesOrderInfo.shipping_discount_amount;
				mage1SalesOrderInfo.discount_amount = mage2SalesOrderInfo.discount_amount;
				// *** discount_canceled
				// *** discount_invoiced
				// *** discount_refunded
				mage1SalesOrderInfo.subtotal = mage2SalesOrderInfo.subtotal;
				// *** subtotal_canceled
				// *** subtotal_invoiced
				// *** subtotal_refunded
				mage1SalesOrderInfo.subtotal_incl_tax = mage2SalesOrderInfo.subtotal_incl_tax;
				mage1SalesOrderInfo.grand_total = mage2SalesOrderInfo.grand_total;
				// *** total_paid
				// *** total_refunded
				mage1SalesOrderInfo.total_qty_ordered = mage2SalesOrderInfo.total_qty_ordered;
				// *** total_canceled
				// *** total_invoiced
				mage1SalesOrderInfo.total_due = mage2SalesOrderInfo.total_due;
				// *** total_online_refunded
				// *** total_offline_refunded
				mage1SalesOrderInfo.base_tax_amount = mage2SalesOrderInfo.base_tax_amount;
				// *** base_tax_canceled
				// *** base_tax_invoiced
				// *** base_tax_refunded
				mage1SalesOrderInfo.base_shipping_amount = mage2SalesOrderInfo.base_shipping_amount;
				// *** base_shipping_canceled
				// *** base_shipping_invoiced
				// *** base_shipping_refunded
				mage1SalesOrderInfo.base_shipping_tax_amount = mage2SalesOrderInfo.base_shipping_tax_amount;
				// *** base_shipping_tax_refunded
				mage1SalesOrderInfo.base_discount_amount = mage2SalesOrderInfo.base_discount_amount;
				// *** base_discount_canceled
				// *** base_discount_invoiced
				// *** base_discount_refunded
				mage1SalesOrderInfo.base_subtotal = mage2SalesOrderInfo.base_subtotal;
				// *** base_subtotal_canceled
				// *** base_subtotal_invoiced
				// *** base_subtotal_refunded
				mage1SalesOrderInfo.base_grand_total = mage2SalesOrderInfo.base_grand_total;
				// *** base_total_paid
				// *** base_total_refunded
				// *** base_total_qty_ordered
				// *** base_total_canceled
				// *** base_total_invoiced
				// *** base_total_invoiced_cost
				// *** base_total_online_refunded
				// *** base_total_offline_refunded
				mage1SalesOrderInfo.billing_address_id = mage2SalesOrderInfo.billing_address_id;
				mage1SalesOrderInfo.billing_firstname = mage2SalesOrderInfo.billing_address.firstname;
				mage1SalesOrderInfo.billing_lastname = mage2SalesOrderInfo.billing_address.lastname;

				if (mage2ShippingAddress != null)
				{
					mage1SalesOrderInfo.shipping_address_id = mage2ShippingAddress.entity_id;
					mage1SalesOrderInfo.shipping_firstname = mage2ShippingAddress.firstname;
					mage1SalesOrderInfo.shipping_lastname = mage2ShippingAddress.lastname;
				}

				// *** billing_name
				// *** shipping_name
				mage1SalesOrderInfo.store_to_base_rate = mage2SalesOrderInfo.store_to_base_rate;
				mage1SalesOrderInfo.store_to_order_rate = mage2SalesOrderInfo.store_to_order_rate;
				mage1SalesOrderInfo.base_to_global_rate = mage2SalesOrderInfo.base_to_global_rate;
				mage1SalesOrderInfo.base_to_order_rate = mage2SalesOrderInfo.base_to_order_rate;
				mage1SalesOrderInfo.weight = mage2SalesOrderInfo.weight;
				mage1SalesOrderInfo.store_name = mage2SalesOrderInfo.store_name;
				mage1SalesOrderInfo.remote_ip = mage2SalesOrderInfo.remote_ip;
				mage1SalesOrderInfo.status = mage2SalesOrderInfo.status;
				mage1SalesOrderInfo.state = mage2SalesOrderInfo.state;
				// *** applied_rule_ids
				mage1SalesOrderInfo.global_currency_code = mage2SalesOrderInfo.global_currency_code;
				mage1SalesOrderInfo.base_currency_code = mage2SalesOrderInfo.base_currency_code;
				mage1SalesOrderInfo.store_currency_code = mage2SalesOrderInfo.store_currency_code;
				mage1SalesOrderInfo.order_currency_code = mage2SalesOrderInfo.order_currency_code;
				if (mage2Shipping != null) mage1SalesOrderInfo.shipping_method = mage2Shipping.method;
				mage1SalesOrderInfo.shipping_description = mage2SalesOrderInfo.shipping_description;
				mage1SalesOrderInfo.customer_email = mage2SalesOrderInfo.customer_email;
				mage1SalesOrderInfo.customer_firstname = mage2SalesOrderInfo.customer_firstname;
				mage1SalesOrderInfo.customer_lastname = mage2SalesOrderInfo.customer_lastname;
				// *** customer_middlename
				// *** customer_prefix
				// *** customer_suffix
				mage1SalesOrderInfo.customer_taxvat = mage2SalesOrderInfo.customer_taxvat;
				mage1SalesOrderInfo.quote_id = mage2SalesOrderInfo.quote_id;
				mage1SalesOrderInfo.is_virtual = mage2SalesOrderInfo.is_virtual;
				mage1SalesOrderInfo.customer_group_id = mage2SalesOrderInfo.customer_group_id;
				// *** customer_note
				mage1SalesOrderInfo.customer_note_notify = mage2SalesOrderInfo.customer_note_notify;
				mage1SalesOrderInfo.customer_is_guest = mage2SalesOrderInfo.customer_is_guest;
				mage1SalesOrderInfo.email_sent = mage2SalesOrderInfo.email_sent;
				mage1SalesOrderInfo.order_id = mage2SalesOrderInfo.entity_id;
				// *** gift_message_id
				// *** gift_message
				// *** coupon_code
				mage1SalesOrderInfo.protect_code = mage2SalesOrderInfo.protect_code;
				// *** can_ship_partially
				// *** can_ship_partially_item
				// *** edit_increment
				// *** forced_shipment_with_invoice
				// *** forced_do_shipment_with_invoice
				// *** payment_auth_expiration
				// *** quote_address_id
				// *** adjustment_negative
				// *** adjustment_positive
				// *** base_adjustment_negative
				// *** base_adjustment_positive
				mage1SalesOrderInfo.base_shipping_discount_amount = mage2SalesOrderInfo.base_shipping_discount_amount;
				mage1SalesOrderInfo.base_subtotal_incl_tax = mage2SalesOrderInfo.base_subtotal_incl_tax;
				mage1SalesOrderInfo.base_total_due = mage2SalesOrderInfo.base_total_due;
				// *** payment_authorization_amount
				// *** customer_dob
				// *** discount_description
				// *** ext_customer_id
				// *** ext_order_id
				// *** hold_before_state
				// *** hold_before_status
				// *** original_increment_id
				// *** relation_child_id
				// *** relation_child_real_id
				// *** relation_parent_id
				// *** relation_parent_real_id
				mage1SalesOrderInfo.x_forwarded_for = mage2SalesOrderInfo.x_forwarded_for;
				mage1SalesOrderInfo.total_item_count = mage2SalesOrderInfo.total_item_count;
				// *** customer_gender
				// *** hidden_tax_amount
				// *** base_hidden_tax_amount
				// *** shipping_hidden_tax_amount
				// *** base_shipping_hidden_tax_amnt
				// *** hidden_tax_invoiced
				// *** base_hidden_tax_invoiced
				// *** hidden_tax_refunded
				// *** base_hidden_tax_refunded
				mage1SalesOrderInfo.shipping_incl_tax = mage2SalesOrderInfo.shipping_incl_tax;
				mage1SalesOrderInfo.base_shipping_incl_tax = mage2SalesOrderInfo.base_shipping_incl_tax;
				// *** coupon_rule_name
				// *** paypal_ipn_customer_notified
				// *** firecheckout_delivery_date
				// *** firecheckout_delivery_timerange
				// *** firecheckout_customer_comment
				// *** tm_field1
				// *** tm_field2
				// *** tm_field3
				// *** tm_field4
				// *** tm_field5
				// *** from_lengow
				// *** order_id_lengow
				// *** fees_lengow
				// *** xml_node_lengow
				// *** feed_id_lengow
				// *** message_lengow
				// *** marketplace_lengow
				// *** total_paid_lengow
				// *** carrier_lengow
				// *** carrier_method_lengow
				// *** clearsale_status_code
				// *** session_id
				if (mage2SkyHubInfo != null) mage1SalesOrderInfo.skyhub_code = mage2SkyHubInfo.code;
				// *** commission_value
				// *** installer_document
				// *** installer_id
				// *** commission_discount
				// *** commission_final_discount
				// *** commission_discount_type
				// *** commission_final_value
				// *** base_bseller_payment_total_tax_rate
				// *** bseller_payment_total_tax_rate
				// *** payment_authorization_expiration
				// *** base_shipping_hidden_tax_amount
				// *** clearSale_status
				// *** clearSale_score
				// *** clearSale_packageID
				// *** clearSale_fingerPrintSessionId
				// *** integracommerce_id
				// *** bseller_skyhub
				// *** bseller_skyhub_code
				// *** bseller_skyhub_channel
				// *** bseller_skyhub_invoice_key
				// *** bseller_skyhub_interest
				// *** bseller_skyhub_json
				#endregion

				#region [ Shipping Address ]
				if (mage2ShippingAddress != null)
				{
					mage1SalesOrderInfo.shipping_address.parent_id = mage2ShippingAddress.parent_id;
					mage1SalesOrderInfo.shipping_address.customer_address_id = mage2ShippingAddress.customer_address_id;
					// *** quote_address_id
					mage1SalesOrderInfo.shipping_address.region_id = mage2ShippingAddress.region_id;
					// *** customer_id
					// *** fax
					mage1SalesOrderInfo.shipping_address.region = mage2ShippingAddress.region;
					mage1SalesOrderInfo.shipping_address.postcode = mage2ShippingAddress.postcode;
					mage1SalesOrderInfo.shipping_address.firstname = mage2ShippingAddress.firstname;
					// *** middlename
					mage1SalesOrderInfo.shipping_address.lastname = mage2ShippingAddress.lastname;

					if (mage2ShippingAddress.street != null)
					{
						sbEndereco = new StringBuilder("");
						for (int i = 0; i < mage2ShippingAddress.street.Count; i++)
						{
							// A última linha do endereço não é seguida por quebra de linha
							// Linhas em branco também devem ser incluídas, pois cada linha possui um significado: logradouro, número, complemento, bairro
							if (i == (mage2ShippingAddress.street.Count - 1))
							{
								sbEndereco.Append(mage2ShippingAddress.street[i]);
							}
							else
							{
								sbEndereco.AppendLine(mage2ShippingAddress.street[i]);
							}
						}
						mage1SalesOrderInfo.shipping_address.street = sbEndereco.ToString();
					}

					mage1SalesOrderInfo.shipping_address.city = mage2ShippingAddress.city;
					mage1SalesOrderInfo.shipping_address.email = mage2ShippingAddress.email;
					mage1SalesOrderInfo.shipping_address.telephone = mage2ShippingAddress.telephone;
					mage1SalesOrderInfo.shipping_address.country_id = mage2ShippingAddress.country_id;
					mage1SalesOrderInfo.shipping_address.address_type = mage2ShippingAddress.address_type;
					// *** prefix
					// *** suffix
					// *** company
					// *** vat_id
					// *** vat_is_valid
					// *** vat_request_id
					// *** vat_request_date
					// *** vat_request_success
					// *** tipopessoa
					// *** rg
					// *** ie
					// *** cpfcnpj
					// *** celular
					// *** empresa
					// *** nomefantasia
					// *** cpf
					mage1SalesOrderInfo.shipping_address.address_id = mage2ShippingAddress.entity_id;
					// *** street_detail
				}
				#endregion

				#region [ Billing Address ]
				if (mage2SalesOrderInfo.billing_address != null)
				{
					mage1SalesOrderInfo.billing_address.parent_id = mage2SalesOrderInfo.billing_address.parent_id;
					// *** customer_address_id
					// *** quote_address_id
					mage1SalesOrderInfo.billing_address.region_id = mage2SalesOrderInfo.billing_address.region_id;
					// *** customer_id
					// *** fax
					mage1SalesOrderInfo.billing_address.region = mage2SalesOrderInfo.billing_address.region;
					mage1SalesOrderInfo.billing_address.postcode = mage2SalesOrderInfo.billing_address.postcode;
					mage1SalesOrderInfo.billing_address.firstname = mage2SalesOrderInfo.billing_address.firstname;
					// *** middlename
					mage1SalesOrderInfo.billing_address.lastname = mage2SalesOrderInfo.billing_address.lastname;

					if (mage2SalesOrderInfo.billing_address.street != null)
					{
						sbEndereco = new StringBuilder("");
						for (int i = 0; i < mage2SalesOrderInfo.billing_address.street.Count; i++)
						{
							// A última linha do endereço não é seguida por quebra de linha
							// Linhas em branco também devem ser incluídas, pois cada linha possui um significado: logradouro, número, complemento, bairro
							if (i == (mage2SalesOrderInfo.billing_address.street.Count - 1))
							{
								sbEndereco.Append(mage2SalesOrderInfo.billing_address.street[i]);
							}
							else
							{
								sbEndereco.AppendLine(mage2SalesOrderInfo.billing_address.street[i]);
							}
						}
						mage1SalesOrderInfo.billing_address.street = sbEndereco.ToString();
					}

					mage1SalesOrderInfo.billing_address.city = mage2SalesOrderInfo.billing_address.city;
					mage1SalesOrderInfo.billing_address.email = mage2SalesOrderInfo.billing_address.email;
					mage1SalesOrderInfo.billing_address.telephone = mage2SalesOrderInfo.billing_address.telephone;
					mage1SalesOrderInfo.billing_address.country_id = mage2SalesOrderInfo.billing_address.country_id;
					mage1SalesOrderInfo.billing_address.address_type = mage2SalesOrderInfo.billing_address.address_type;
					// *** prefix
					// *** suffix
					// *** company
					// *** vat_id
					// *** vat_is_valid
					// *** vat_request_id
					// *** vat_request_date
					// *** vat_request_success
					// *** tipopessoa
					// *** rg
					// *** ie
					// *** cpfcnpj
					// *** celular
					// *** empresa
					// *** nomefantasia
					// *** cpf
					mage1SalesOrderInfo.billing_address.address_id = mage2SalesOrderInfo.billing_address.entity_id;
					// *** street_detail
				}
				#endregion

				#region [ Sales Order Item ]
				if (mage2SalesOrderInfo.items != null)
				{
					foreach (Magento2SalesOrderItem mage2SalesOrderItem in mage2SalesOrderInfo.items)
					{
						// Ignora os produtos do tipo 'configurable' (processa somente os do tipo 'simple' e 'virtual')
						if (mage2SalesOrderItem.product_type.Equals("configurable")) continue;

						mage1SalesOrderItem = new SalesOrderItem();
						mage1SalesOrderItem.item_id = mage2SalesOrderItem.item_id;
						mage1SalesOrderItem.order_id = mage2SalesOrderItem.order_id;
						mage1SalesOrderItem.parent_item_id = mage2SalesOrderItem.parent_item_id;
						mage1SalesOrderItem.quote_item_id = mage2SalesOrderItem.quote_item_id;
						mage1SalesOrderItem.store_id = mage2SalesOrderItem.store_id;
						mage1SalesOrderItem.created_at = mage2SalesOrderItem.created_at;
						mage1SalesOrderItem.updated_at = mage2SalesOrderItem.updated_at;
						mage1SalesOrderItem.product_id = mage2SalesOrderItem.product_id;
						mage1SalesOrderItem.product_type = mage2SalesOrderItem.product_type;
						// *** product_options
						mage1SalesOrderItem.weight = mage2SalesOrderItem.weight;
						mage1SalesOrderItem.is_virtual = mage2SalesOrderItem.is_virtual;
						mage1SalesOrderItem.sku = mage2SalesOrderItem.sku;
						mage1SalesOrderItem.name = mage2SalesOrderItem.name;
						// *** description
						// *** applied_rule_ids
						// *** additional_data
						mage1SalesOrderItem.free_shipping = mage2SalesOrderItem.free_shipping;
						mage1SalesOrderItem.is_qty_decimal = mage2SalesOrderItem.is_qty_decimal;
						mage1SalesOrderItem.no_discount = mage2SalesOrderItem.no_discount;
						// *** qty_backordered
						mage1SalesOrderItem.qty_canceled = mage2SalesOrderItem.qty_canceled;
						mage1SalesOrderItem.qty_invoiced = mage2SalesOrderItem.qty_invoiced;
						mage1SalesOrderItem.qty_ordered = mage2SalesOrderItem.qty_ordered;
						mage1SalesOrderItem.qty_refunded = mage2SalesOrderItem.qty_refunded;
						mage1SalesOrderItem.qty_shipped = mage2SalesOrderItem.qty_shipped;
						// *** base_cost
						mage1SalesOrderItem.price = mage2SalesOrderItem.price;
						mage1SalesOrderItem.base_price = mage2SalesOrderItem.base_price;
						mage1SalesOrderItem.original_price = mage2SalesOrderItem.original_price;
						mage1SalesOrderItem.base_original_price = mage2SalesOrderItem.base_original_price;
						mage1SalesOrderItem.tax_percent = mage2SalesOrderItem.tax_percent;
						mage1SalesOrderItem.tax_amount = mage2SalesOrderItem.tax_amount;
						mage1SalesOrderItem.base_tax_amount = mage2SalesOrderItem.base_tax_amount;
						mage1SalesOrderItem.tax_invoiced = mage2SalesOrderItem.tax_invoiced;
						mage1SalesOrderItem.base_tax_invoiced = mage2SalesOrderItem.base_tax_invoiced;
						mage1SalesOrderItem.discount_percent = mage2SalesOrderItem.discount_percent;
						mage1SalesOrderItem.discount_amount = mage2SalesOrderItem.discount_amount;
						mage1SalesOrderItem.base_discount_amount = mage2SalesOrderItem.base_discount_amount;
						mage1SalesOrderItem.discount_invoiced = mage2SalesOrderItem.discount_invoiced;
						mage1SalesOrderItem.base_discount_invoiced = mage2SalesOrderItem.base_discount_invoiced;
						mage1SalesOrderItem.amount_refunded = mage2SalesOrderItem.amount_refunded;
						mage1SalesOrderItem.base_amount_refunded = mage2SalesOrderItem.base_amount_refunded;
						mage1SalesOrderItem.row_total = mage2SalesOrderItem.row_total;
						mage1SalesOrderItem.base_row_total = mage2SalesOrderItem.base_row_total;
						mage1SalesOrderItem.row_invoiced = mage2SalesOrderItem.row_invoiced;
						mage1SalesOrderItem.base_row_invoiced = mage2SalesOrderItem.base_row_invoiced;
						mage1SalesOrderItem.row_weight = mage2SalesOrderItem.row_weight;
						// *** base_tax_before_discount
						// *** tax_before_discount
						// *** ext_order_item_id
						// *** locked_do_invoice
						// *** locked_do_ship
						mage1SalesOrderItem.price_incl_tax = mage2SalesOrderItem.price_incl_tax;
						mage1SalesOrderItem.base_price_incl_tax = mage2SalesOrderItem.base_price_incl_tax;
						mage1SalesOrderItem.row_total_incl_tax = mage2SalesOrderItem.row_total_incl_tax;
						mage1SalesOrderItem.base_row_total_incl_tax = mage2SalesOrderItem.base_row_total_incl_tax;
						// *** hidden_tax_amount
						// *** base_hidden_tax_amount
						// *** hidden_tax_invoiced
						// *** base_hidden_tax_invoiced
						// *** hidden_tax_refunded
						// *** base_hidden_tax_refunded
						// *** is_nominal
						// *** tax_canceled
						// *** hidden_tax_canceled
						// *** tax_refunded
						// *** base_tax_refunded
						// *** discount_refunded
						// *** base_discount_refunded
						// *** gift_message_id
						// *** gift_message_available
						// *** base_weee_tax_applied_amount
						// *** base_weee_tax_applied_row_amnt
						// *** base_weee_tax_applied_row_amount
						// *** weee_tax_applied_amount
						// *** weee_tax_applied_row_amount
						// *** weee_tax_applied
						// *** weee_tax_disposition
						// *** weee_tax_row_disposition
						// *** base_weee_tax_disposition
						// *** base_weee_tax_row_disposition
						// *** installer_document
						// *** commission_type
						// *** commission_value
						// *** has_children

						mage1SalesOrderInfo.items.Add(mage1SalesOrderItem);
					}
				}
				#endregion

				#region [ Sales Order Payment ]
				if (mage2SalesOrderInfo.payment != null)
				{
					mage1SalesOrderInfo.payment.parent_id = mage2SalesOrderInfo.payment.parent_id;
					// *** base_shipping_captured
					// *** shipping_captured
					// *** amount_refunded
					// *** base_amount_paid
					// *** amount_canceled
					mage1SalesOrderInfo.payment.base_amount_authorized = mage2SalesOrderInfo.payment.base_amount_authorized;
					// *** base_amount_paid_online
					// *** base_amount_refunded_online
					mage1SalesOrderInfo.payment.base_shipping_amount = mage2SalesOrderInfo.payment.base_shipping_amount;
					mage1SalesOrderInfo.payment.shipping_amount = mage2SalesOrderInfo.payment.shipping_amount;
					// *** amount_paid
					mage1SalesOrderInfo.payment.amount_authorized = mage2SalesOrderInfo.payment.amount_authorized;
					mage1SalesOrderInfo.payment.base_amount_ordered = mage2SalesOrderInfo.payment.base_amount_ordered;
					// *** base_shipping_refunded
					// *** shipping_refunded
					// *** base_amount_refunded
					mage1SalesOrderInfo.payment.amount_ordered = mage2SalesOrderInfo.payment.amount_ordered;
					// *** base_amount_canceled
					// *** quote_payment_id
					// *** additional_data
					// *** cc_exp_month
					mage1SalesOrderInfo.payment.cc_ss_start_year = mage2SalesOrderInfo.payment.cc_ss_start_year;
					// *** echeck_bank_name
					mage1SalesOrderInfo.payment.method = mage2SalesOrderInfo.payment.method;
					// *** cc_debug_request_body
					// *** cc_secure_verify
					// *** protection_eligibility
					// *** cc_approval
					mage1SalesOrderInfo.payment.cc_last4 = mage2SalesOrderInfo.payment.cc_last4;
					// *** cc_status_description
					// *** echeck_type
					// *** cc_debug_response_serialized
					mage1SalesOrderInfo.payment.cc_ss_start_month = mage2SalesOrderInfo.payment.cc_ss_start_month;
					// *** echeck_account_type
					mage1SalesOrderInfo.payment.last_trans_id = mage2SalesOrderInfo.payment.last_trans_id;
					// *** cc_cid_status
					// *** cc_owner
					// *** cc_type
					// *** po_number
					// *** cc_exp_year
					// *** cc_status
					// *** echeck_routing_number
					mage1SalesOrderInfo.payment.account_status = mage2SalesOrderInfo.payment.account_status;
					// *** anet_trans_method
					// *** cc_debug_response_body
					// *** cc_ss_issue
					// *** echeck_account_name
					// *** cc_avs_status
					// *** cc_number_enc
					// *** cc_trans_id
					// *** paybox_request_number
					// *** address_status
					// *** cc_parcelamento
					// *** cc_type2
					// *** cc_owner2
					// *** cc_last42
					// *** cc_number_enc2
					// *** cc_exp_month2
					// *** cc_exp_year2
					// *** cc_ss_issue2
					// *** cc_cid2
					// *** cc_parcelamento2
					// *** bseller_payment_in_cash
					// *** bseller_payment_installment
					mage1SalesOrderInfo.payment.payment_id = mage2SalesOrderInfo.payment.entity_id;
					// *** integracommerce_name
					// *** integracommerce_installments
					// *** additional_information (OBS: apesar de haver campos com o mesmo nome em ambas as classes, aparentemente são incompatíveis, pois antes a estrutura possuía campos com nomes definidos e agora é um array de strings)
					// *** additional_information2 (OBS: mesmo caso do campo 'additional_information')
				}
				#endregion

				#region [ Status History ]
				if (mage2SalesOrderInfo.status_histories != null)
				{
					foreach (Magento2StatusHistory mage2StatusHistory in mage2SalesOrderInfo.status_histories)
					{
						mage1StatusHistory = new StatusHistory();
						mage1StatusHistory.parent_id = mage2StatusHistory.parent_id;
						mage1StatusHistory.is_customer_notified = mage2StatusHistory.is_customer_notified;
						mage1StatusHistory.is_visible_on_front = mage2StatusHistory.is_visible_on_front;
						mage1StatusHistory.comment = mage2StatusHistory.comment;
						mage1StatusHistory.status = mage2StatusHistory.status;
						mage1StatusHistory.created_at = mage2StatusHistory.created_at;
						mage1StatusHistory.entity_name = mage2StatusHistory.entity_name;
						// *** store_id
						mage1SalesOrderInfo.status_history.Add(mage1StatusHistory);
					}
				}
				#endregion

				return mage1SalesOrderInfo;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				return null;
			}
		}
		#endregion

		#region [ decodificaSalesOrderInfoJsonEntityIdResponseMage2ParaObjMage1 ]
		public static SalesOrderInfo decodificaSalesOrderInfoJsonEntityIdResponseMage2ParaObjMage1(string jsonPedido, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Magento2RestApi.decodificaSalesOrderInfoJsonEntityIdResponseMage2ParaObjMage1()";
			string msg_erro_aux;
			Magento2SalesOrderInfo mage2SalesOrderInfo;
			#endregion

			msg_erro = "";

			try
			{
				mage2SalesOrderInfo = decodificaJsonSalesOrderInfo(jsonPedido, out msg_erro_aux);
				if (mage2SalesOrderInfo == null)
				{
					msg_erro = NOME_DESTA_ROTINA + ": Erro ao tentar decodificar os dados do pedido em formato JSON!";
					if ((msg_erro_aux ?? "").Length > 0) msg_erro += "\n" + msg_erro_aux;
					return null;
				}

				return decodificaSalesOrderInfoMage2ParaMage1(mage2SalesOrderInfo, out msg_erro);
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				return null;
			}
		}
		#endregion

		#region [ decodificaSalesOrderInfoJsonSearchResponseMage2ParaObjMage1 ]
		public static SalesOrderInfo decodificaSalesOrderInfoJsonSearchResponseMage2ParaObjMage1(string jsonPedido, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Magento2RestApi.decodificaSalesOrderInfoJsonSearchResponseMage2ParaObjMage1()";
			string msg_erro_aux;
			Magento2SalesOrderInfo mage2SalesOrderInfo;
			Magento2SalesOrderSearchResponse mage2SearchResponse;
			#endregion

			msg_erro = "";

			try
			{
				mage2SearchResponse = decodificaJsonSalesOrderSearchResponse(jsonPedido, out msg_erro_aux);
				if (mage2SearchResponse == null)
				{
					msg_erro = NOME_DESTA_ROTINA + ": Erro ao tentar decodificar os dados do pedido em formato JSON!";
					if ((msg_erro_aux ?? "").Length > 0) msg_erro += "\n" + msg_erro_aux;
					return null;
				}

				if (mage2SearchResponse.total_count == 0)
				{
					msg_erro = NOME_DESTA_ROTINA + ": Nenhum pedido encontrado no resultado da pesquisa!";
					return null;
				}

				mage2SalesOrderInfo = mage2SearchResponse.items[0];
				if (mage2SalesOrderInfo == null)
				{
					msg_erro = NOME_DESTA_ROTINA + ": Erro ao tentar recuperar os dados do pedido em formato JSON!";
					if ((msg_erro_aux ?? "").Length > 0) msg_erro += "\n" + msg_erro_aux;
					return null;
				}

				return decodificaSalesOrderInfoMage2ParaMage1(mage2SalesOrderInfo, out msg_erro);
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				return null;
			}
		}
		#endregion

		#region [ decodificaJsonProductSearchResponse ]
		public static Magento2ProductSearchResponse decodificaJsonProductSearchResponse(string respJson, out string msg_erro)
		{
			Magento2ProductSearchResponse ret;

			msg_erro = "";

			try
			{
				if ((respJson ?? "").Trim().Length == 0) return null;

				ret = JsonConvert.DeserializeObject<Magento2ProductSearchResponse>(respJson);

				return ret;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				return null;
			}
		}
		#endregion

		#region [ retornaShipping ]
		public static Magento2ExtensionAttributesShippingAssignmentsShipping retornaShipping(Magento2SalesOrderInfo salesOrderInfo)
		{
			if (salesOrderInfo == null) return null;
			if (salesOrderInfo.extension_attributes == null) return null;
			if (salesOrderInfo.extension_attributes.shipping_assignments == null) return null;
			if (salesOrderInfo.extension_attributes.shipping_assignments.Count == 0) return null;
			if (salesOrderInfo.extension_attributes.shipping_assignments[0].shipping == null) return null;

			return salesOrderInfo.extension_attributes.shipping_assignments[0].shipping;
		}
		#endregion

		#region [ retornaShippingAddress ]
		public static Magento2ExtensionAttributesShippingAssignmentsShippingAddress retornaShippingAddress(Magento2SalesOrderInfo salesOrderInfo)
		{
			if (salesOrderInfo == null) return null;
			if (salesOrderInfo.extension_attributes == null) return null;
			if (salesOrderInfo.extension_attributes.shipping_assignments == null) return null;
			if (salesOrderInfo.extension_attributes.shipping_assignments.Count == 0) return null;
			if (salesOrderInfo.extension_attributes.shipping_assignments[0].shipping == null) return null;
			if (salesOrderInfo.extension_attributes.shipping_assignments[0].shipping.address == null) return null;

			return salesOrderInfo.extension_attributes.shipping_assignments[0].shipping.address;
		}
		#endregion

		#region [ retornaSkyHubInfo ]
		public static Magento2ExtensionAttributesSkyhubInfo retornaSkyHubInfo(Magento2SalesOrderInfo salesOrderInfo)
		{
			if (salesOrderInfo == null) return null;
			if (salesOrderInfo.extension_attributes == null) return null;
			if (salesOrderInfo.extension_attributes.skyhub_info == null) return null;

			return salesOrderInfo.extension_attributes.skyhub_info;
		}
		#endregion

		#region [ formataDadosCampoComQuebraLinha ]
		public static string formataDadosCampoComQuebraLinha(string margem, string nomeCampo, string valorCampo)
		{
			#region [ Declarações ]
			string sTexto;
			string sMargemQuebraLinha;
			#endregion

			sTexto = (valorCampo ?? "").ToString();
			if (sTexto.Contains("\n") && (!sTexto.Contains("\r")))
			{
				sTexto = sTexto.Replace("\n", "\r\n");
			}

			if (sTexto.Contains("\r") && (!sTexto.Contains("\n")))
			{
				sTexto = sTexto.Replace("\r", "\r\n");
			}

			sMargemQuebraLinha = new string(' ', (nomeCampo + " = ").Length);
			sTexto = sTexto.Replace("\r\n", "\r\n" + margem + sMargemQuebraLinha);

			return sTexto;
		}
		#endregion
	}
	#endregion
}
