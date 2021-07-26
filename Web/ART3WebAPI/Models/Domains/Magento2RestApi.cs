using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Web;
using System.Web.Script.Serialization;
using ART3WebAPI.Models.Entities;
using ART3WebAPI.Models.Repository;
using Newtonsoft.Json;

namespace ART3WebAPI.Models.Domains
{
	public static class Magento2RestApi
	{
		public const char SEPARADOR_DECIMAL_NUM_REAL = '.';
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
			const string NOME_DESTA_ROTINA = "Magento2RestApi.enviaRequisicaoGet()";
			string strMsg;
			bool blnSucesso;
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
						blnSucesso = false;
					}
				}

				strMsg = NOME_DESTA_ROTINA + " - RX\nSucesso: " + blnSucesso.ToString();
				if (blnSucesso) strMsg += "\n" + respRest;
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

		#region [ getSalesOrderInfo ]
		public static Magento2SalesOrderInfo getSalesOrderInfo(string numeroPedidoMagento, MagentoApiLoginParameters loginParameters, out string jsonResponse, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Magento2RestApi.getSalesOrderInfo()";
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

				if ((loginParameters.api_rest_endpoint ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o endpoint da API REST do Magento!";
					return null;
				}

				if ((loginParameters.api_rest_access_token ?? "").Trim().Length == 0)
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

						urlParamReqRest = Magento2RestApi.montaRequisicaoGetSalesOrderInfoByIncrementId(numeroPedidoMagento, loginParameters.api_rest_endpoint, out urlBaseAddress);
						blnEnviouOk = Magento2RestApi.enviaRequisicaoGetComRetry(urlParamReqRest, loginParameters.api_rest_access_token, urlBaseAddress, out respJson, out msg_erro_aux);
						if (!blnEnviouOk)
						{
							msg_erro = "Falha ao tentar consultar o pedido Magento " + numeroPedidoMagento + " através da API REST!";
							if (msg_erro_aux.Length > 0) msg_erro += "\n" + msg_erro_aux;
							return null;
						}

						jsonResponse = respJson;

						salesOrderInfoSearchResponse = Magento2RestApi.decodificaJsonSalesOrderSearchResponse(respJson, out msg_erro_aux);
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

						if (loginParameters.api_rest_force_get_sales_order_by_entity_id == 0)
						{
							return salesOrderInfo;
						}
						else
						{
							urlParamReqRest = Magento2RestApi.montaRequisicaoGetSalesOrderInfoByEntityId(salesOrderInfo.entity_id, loginParameters.api_rest_endpoint, out urlBaseAddress);
							blnEnviouOk = Magento2RestApi.enviaRequisicaoGetComRetry(urlParamReqRest, loginParameters.api_rest_access_token, urlBaseAddress, out respJson, out msg_erro_aux);

							if (!blnEnviouOk)
							{
								msg_erro = "Falha ao tentar consultar o pedido Magento " + numeroPedidoMagento + " via entity_id através da API REST!";
								if (msg_erro_aux.Length > 0) msg_erro += "\n" + msg_erro_aux;
								return null;
							}

							jsonResponse = respJson;

							salesOrderInfo = Magento2RestApi.decodificaJsonSalesOrderInfo(respJson, out msg_erro_aux);
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
		public static MagentoSoapApiSalesOrderInfo decodificaSalesOrderInfoMage2ParaMage1(Magento2SalesOrderInfo mage2SalesOrderInfo, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Magento2RestApi.decodificaSalesOrderInfoMage2ParaMage1()";
			bool blnUsarParentItem;
			StringBuilder sbEndereco;
			MagentoSoapApiSalesOrderInfo mage1SalesOrderInfo = new MagentoSoapApiSalesOrderInfo();
			Magento2ExtensionAttributesShippingAssignmentsShipping mage2Shipping;
			Magento2ExtensionAttributesShippingAssignmentsShippingAddress mage2ShippingAddress;
			Magento2ExtensionAttributesSkyhubInfo mage2SkyHubInfo;
			MagentoSoapApiSalesOrderItem mage1SalesOrderItem;
			MagentoSoapApiStatusHistory mage1StatusHistory;
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
				if (mage2SkyHubInfo != null)
				{
					mage1SalesOrderInfo.skyhub_code = mage2SkyHubInfo.code;
					mage1SalesOrderInfo.bseller_skyhub_code = mage2SkyHubInfo.code;
					mage1SalesOrderInfo.bseller_skyhub_channel = mage2SkyHubInfo.channel;
					mage1SalesOrderInfo.bseller_skyhub_json = mage2SkyHubInfo.data_source;
				}
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
				// *** bseller_skyhub_invoice_key
				// *** bseller_skyhub_interest
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
				if (mage2SalesOrderInfo.items!=null)
				{
					foreach (Magento2SalesOrderItem mage2SalesOrderItem in mage2SalesOrderInfo.items)
					{
						// Ignora os produtos do tipo 'configurable' (processa somente os do tipo 'simple' e 'virtual')
						// IMPORTANTE: O Magento 2 possui o conceito de produto 'configurable', ou seja, um produto que representa
						// ==========  um determinado tipo de produto que possui algum tipo de variação (potência, voltagem, etc),
						// mas com as demais características permanecendo as mesmas.
						// Um produto 'configurable' possui no seu conjunto de opções vários produtos 'simple'.
						// Mas um produto 'simple' pode não estar vinculado a um produto 'configurable'.
						// Dito isso, o Magento 2 pode retornar os dados de um item do pedido em uma das seguintes formas:
						//    1) Produto 'simple' que não está vinculado a nenhum produto 'configurable': retorna os dados do item
						//       em um bloco de dados em que o campo 'product_type' será igual a 'simple'.
						//    2) Produto que está vinculado a um produto 'configurable': neste caso, dependerá de como o produto foi
						//       adicionado ao carrinho.
						//       2-A) Se o produto estiver sendo exibido em um grid em que está sendo apresentado com os
						//            dados/características de uma determinada versão 'simple' e for adicionado ao carrinho diretamente
						//            desse local, então no resultado da consulta da API os dados serão retornados da mesma forma
						//            como acontece no produto 'simple' que não está vinculado a nenhum produto 'configurable', ou seja,
						//            no mesmo formato do item (1).
						//       2-B) Se o cliente estiver na página do produto que exibe as opções para selecionar potência/voltagem/etc
						//            e adiciona ao carrinho, o resultado da API irá retornar a seguinte estrutura para esse item:
						//            i) Bloco de dados com 'product_type' igual a 'configurable' cujo campo 'name' irá exibir o nome
						//               do produto 'configurable' (nome genérico).
						//            ii) Bloco de dados com 'product_type' igual a 'simple' cujo campo 'name' irá exibir o nome do produto
						//                'simple' selecionado.
						//            iii) Dentro do bloco de dados do item anterior (ii), haverá um campo chamado 'parent_item' que
						//                 irá repetir os mesmos campos já informados no item (i).
						//
						// ATENÇÃO: é fundamental ter conhecimento de que na ocorrência da situação (2-B ii), a grande maioria dos
						// =======  campos de valores irá ser informada com zero, sendo necessário buscar essas informações dentro do
						// bloco de dados informado em 'parent_item'. Até onde se notou nos testes realizados, aparentemente apenas
						// os campos 'price', 'qty_ordered' e 'weight' retornam com valor, os demais sendo informados com zero.
						// Para os demais campos, não foi possível determinar se ocorre o mesmo porque os valores estavam zerados
						// nos dois blocos de dados.
						//
						// Por fim, a forma como o item é adicionado ao carrinho (2A ou 2B) também impacta na exibição de algumas
						// informações do produto na página do pedido no painel Admin.

						if (mage2SalesOrderItem.product_type.Equals("configurable")) continue;

						blnUsarParentItem = false;
						if (mage2SalesOrderItem.product_type.Equals("simple"))
						{
							if (mage2SalesOrderItem.parent_item != null)
							{
								if ((mage2SalesOrderItem.parent_item.product_type.Equals("configurable")) && ((mage2SalesOrderItem.parent_item.sku ?? "").Trim().Length > 0)) blnUsarParentItem = true;
							}
						}

						mage1SalesOrderItem = new MagentoSoapApiSalesOrderItem();
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
						mage1SalesOrderItem.free_shipping = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.free_shipping : mage2SalesOrderItem.free_shipping);
						mage1SalesOrderItem.is_qty_decimal = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.is_qty_decimal : mage2SalesOrderItem.is_qty_decimal);
						mage1SalesOrderItem.no_discount = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.no_discount : mage2SalesOrderItem.no_discount);
						// *** qty_backordered
						mage1SalesOrderItem.qty_canceled = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.qty_canceled : mage2SalesOrderItem.qty_canceled);
						mage1SalesOrderItem.qty_invoiced = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.qty_invoiced : mage2SalesOrderItem.qty_invoiced);
						mage1SalesOrderItem.qty_ordered = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.qty_ordered : mage2SalesOrderItem.qty_ordered);
						mage1SalesOrderItem.qty_refunded = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.qty_refunded : mage2SalesOrderItem.qty_refunded);
						mage1SalesOrderItem.qty_shipped = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.qty_shipped : mage2SalesOrderItem.qty_shipped);
						// *** base_cost
						mage1SalesOrderItem.price = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.price : mage2SalesOrderItem.price);
						mage1SalesOrderItem.base_price = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.base_price : mage2SalesOrderItem.base_price);
						mage1SalesOrderItem.original_price = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.original_price : mage2SalesOrderItem.original_price);
						mage1SalesOrderItem.base_original_price = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.base_original_price : mage2SalesOrderItem.base_original_price);
						mage1SalesOrderItem.tax_percent = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.tax_percent : mage2SalesOrderItem.tax_percent);
						mage1SalesOrderItem.tax_amount = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.tax_amount : mage2SalesOrderItem.tax_amount);
						mage1SalesOrderItem.base_tax_amount = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.base_tax_amount : mage2SalesOrderItem.base_tax_amount);
						mage1SalesOrderItem.tax_invoiced = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.tax_invoiced : mage2SalesOrderItem.tax_invoiced);
						mage1SalesOrderItem.base_tax_invoiced = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.base_tax_invoiced : mage2SalesOrderItem.base_tax_invoiced);
						mage1SalesOrderItem.discount_percent = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.discount_percent : mage2SalesOrderItem.discount_percent);
						mage1SalesOrderItem.discount_amount = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.discount_amount : mage2SalesOrderItem.discount_amount);
						mage1SalesOrderItem.base_discount_amount = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.base_discount_amount : mage2SalesOrderItem.base_discount_amount);
						mage1SalesOrderItem.discount_invoiced = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.discount_invoiced : mage2SalesOrderItem.discount_invoiced);
						mage1SalesOrderItem.base_discount_invoiced = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.base_discount_invoiced : mage2SalesOrderItem.base_discount_invoiced);
						mage1SalesOrderItem.amount_refunded = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.amount_refunded : mage2SalesOrderItem.amount_refunded);
						mage1SalesOrderItem.base_amount_refunded = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.base_amount_refunded : mage2SalesOrderItem.base_amount_refunded);
						mage1SalesOrderItem.row_total = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.row_total : mage2SalesOrderItem.row_total);
						mage1SalesOrderItem.base_row_total = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.base_row_total : mage2SalesOrderItem.base_row_total);
						mage1SalesOrderItem.row_invoiced = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.row_invoiced : mage2SalesOrderItem.row_invoiced);
						mage1SalesOrderItem.base_row_invoiced = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.base_row_invoiced : mage2SalesOrderItem.base_row_invoiced);
						mage1SalesOrderItem.row_weight = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.row_weight : mage2SalesOrderItem.row_weight);
						// *** base_tax_before_discount
						// *** tax_before_discount
						// *** ext_order_item_id
						// *** locked_do_invoice
						// *** locked_do_ship
						mage1SalesOrderItem.price_incl_tax = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.price_incl_tax : mage2SalesOrderItem.price_incl_tax);
						mage1SalesOrderItem.base_price_incl_tax = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.base_price_incl_tax : mage2SalesOrderItem.base_price_incl_tax);
						mage1SalesOrderItem.row_total_incl_tax = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.row_total_incl_tax : mage2SalesOrderItem.row_total_incl_tax);
						mage1SalesOrderItem.base_row_total_incl_tax = (blnUsarParentItem ? mage2SalesOrderItem.parent_item.base_row_total_incl_tax : mage2SalesOrderItem.base_row_total_incl_tax);
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
						mage1StatusHistory = new MagentoSoapApiStatusHistory();
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
		public static MagentoSoapApiSalesOrderInfo decodificaSalesOrderInfoJsonEntityIdResponseMage2ParaObjMage1(string jsonPedido, out string msg_erro)
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
		public static MagentoSoapApiSalesOrderInfo decodificaSalesOrderInfoJsonSearchResponseMage2ParaObjMage1(string jsonPedido, out string msg_erro)
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

		#region [ processaGetPedido ]
		public static string processaGetPedido(string numeroPedidoMagento, string operationControlTicket, string loja, string usuario, string sessionToken, MagentoApiLoginParameters loginParameters)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Magento2RestApi.processaGetPedido()";
			bool blnInserted = false;
			int intParametroFlagCadSemiAutoPedMagentoCadastrarAutomaticamenteClienteNovo;
			string msg;
			string msg_erro = "";
			string sJson = null;
			string s;
			string sParametro;
			string sValue;
			string[] vValue;
			string sMktpOrderDescriptor;
			string[] vMktpOrderDescriptor;
			string sComment;
			string[] vComment;
			string cpfCnpjIdentificado;
			string sNumPedidoMktpIdentificado;
			string sNumPedidoMktpCompletoIdentificado;
			string sOrigemMktpIdentificado;
			string sDDD;
			string sTelefone;
			Cliente cliente;
			MagentoErpPedidoXml readPedidoApiBd = null;
			MagentoErpPedidoXml insertPedidoXml = null;
			MagentoErpSalesOrder salesOrder = new MagentoErpSalesOrder();
			MagentoSoapApiStatusHistory statusHistory;
			List<string> listaPedidosERP;
			Pedido pedidoERP;
			MagentoErpPedidoXmlDecodeEndereco decodeEndereco;
			MagentoErpPedidoXmlDecodeItem decodeItem;
			MagentoErpPedidoXmlDecodeStatusHistory decodeStatusHistory;
			List<CodigoDescricao> listaCodigoDescricao;
			string[] v;
			Magento2SalesOrderInfo mage2SalesOrderInfo;
			#endregion

			#region [ Verifica se o pedido já foi consultado e a resposta se encontra gravada no BD, desde que se trate da mesma operação ]
			if ((operationControlTicket ?? "").Trim().Length > 0)
			{
				if ((numeroPedidoMagento ?? "").Trim().Length == 0)
				{
					msg = "O número do pedido Magento não foi informado!";
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + msg);
					throw new Exception(msg);
				}

				// Apesar da nomenclatura 'xml', os dados obtidos através da API REST/JSON estão armazenados na mesma tabela anterior para minimizar os ajustes necessários,
				// principalmente nas páginas ASP que utilizam os dados armazenados durante o processo de cadastramento semi-automático de pedidos do Magento.
				readPedidoApiBd = MagentoApiDAO.getMagentoPedidoXmlByTicket(numeroPedidoMagento, operationControlTicket, Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V2_REST_JSON, out msg_erro);
				if (readPedidoApiBd != null)
				{
					msg = "Pedido Magento nº " + numeroPedidoMagento + " localizado no BD";
					Global.gravaLogAtividade(msg);

					salesOrder.cpfCnpjIdentificado = readPedidoApiBd.cpfCnpjIdentificado;

					if ((readPedidoApiBd.pedido_json ?? "").Trim().Length > 0)
					{
						sJson = readPedidoApiBd.pedido_json;

						#region [ Converte JSON da resposta do Magento em objeto que representa o pedido no formato do Magento 1 ]
						if (loginParameters.api_rest_force_get_sales_order_by_entity_id == 1)
						{
							salesOrder.magentoSalesOrderInfo = Magento2RestApi.decodificaSalesOrderInfoJsonEntityIdResponseMage2ParaObjMage1(sJson, out msg_erro);
						}
						else
						{
							salesOrder.magentoSalesOrderInfo = Magento2RestApi.decodificaSalesOrderInfoJsonSearchResponseMage2ParaObjMage1(sJson, out msg_erro);
						}

						if (salesOrder.magentoSalesOrderInfo == null)
						{
							msg = "Falha ao tentar decodificar o JSON de resposta da API do Magento!";
							if (msg_erro.Length > 0) msg += "\n" + msg_erro;
							Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + msg);
							throw new Exception(msg);
						}
						#endregion
					}
				}
			}
			#endregion

			#region [ Não encontrou os dados do pedido armazenados no BD, executa consulta via API ]
			if ((readPedidoApiBd == null) || ((sJson ?? "").Trim().Length == 0))
			{
				#region [ Há parâmetros de login cadastrados para a loja? ]
				if ((loginParameters.api_rest_endpoint ?? "").Trim().Length == 0)
				{
					msg = "Falha ao tentar recuperar os parâmetros de acesso à API do Magento: o endpoint da API não está cadastrado para a loja " + loja + "!";
					if (msg_erro.Length > 0) msg += "\n" + msg_erro;
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + msg);
					throw new Exception(msg);
				}

				if ((loginParameters.api_rest_access_token ?? "").Trim().Length == 0)
				{
					msg = "Falha ao tentar recuperar os parâmetros de acesso à API do Magento: o access token não está cadastrado para a loja " + loja + "!";
					if (msg_erro.Length > 0) msg += "\n" + msg_erro;
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + msg);
					throw new Exception(msg);
				}
				#endregion

				#region [ Executa a consulta via API ]
				msg = "Consulta do pedido Magento nº " + numeroPedidoMagento + " via API REST";
				Global.gravaLogAtividade(msg);
				mage2SalesOrderInfo = Magento2RestApi.getSalesOrderInfo(numeroPedidoMagento, loginParameters, out sJson, out msg_erro);
				#endregion

				#region [ Falha ao obter os dados do pedido Magento ]
				if (mage2SalesOrderInfo == null)
				{
					msg = "Falha ao tentar consultar os dados do pedido Magento " + numeroPedidoMagento + " via API REST!";
					if (msg_erro.Length > 0) msg += "\n" + msg_erro;
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + msg);
					throw new Exception(msg);
				}
				#endregion

				#region [ Converte o pedido no formato do Magento 2 para Magento 1 ]
				salesOrder.magentoSalesOrderInfo = decodificaSalesOrderInfoMage2ParaMage1(mage2SalesOrderInfo, out msg_erro);
				if (salesOrder.magentoSalesOrderInfo == null)
				{
					msg = "Falha ao tentar processar a resposta com os dados do pedido Magento " + numeroPedidoMagento + " via API REST!";
					if (msg_erro.Length > 0) msg += "\n" + msg_erro;
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + msg);
					throw new Exception(msg);
				}
				#endregion

				#region [ Obtém o CPF/CNPJ do cliente nos dados do pedido ]
				cpfCnpjIdentificado = "";
				if ((cpfCnpjIdentificado.Length == 0) && ((salesOrder.magentoSalesOrderInfo.billing_address.cpfcnpj ?? "").Trim().Length > 0))
				{
					s = Global.digitos(salesOrder.magentoSalesOrderInfo.billing_address.cpfcnpj);
					if (Global.isCnpjCpfOk(s)) cpfCnpjIdentificado = s;
				}

				if ((cpfCnpjIdentificado.Length == 0) && ((salesOrder.magentoSalesOrderInfo.shipping_address.cpfcnpj ?? "").Trim().Length > 0))
				{
					s = Global.digitos(salesOrder.magentoSalesOrderInfo.shipping_address.cpfcnpj);
					if (Global.isCnpjCpfOk(s)) cpfCnpjIdentificado = s;
				}

				if ((cpfCnpjIdentificado.Length == 0) && ((salesOrder.magentoSalesOrderInfo.customer_taxvat ?? "").Trim().Length > 0))
				{
					s = Global.digitos(salesOrder.magentoSalesOrderInfo.customer_taxvat);
					if (Global.isCnpjCpfOk(s)) cpfCnpjIdentificado = s;
				}

				if ((cpfCnpjIdentificado.Length == 0) && ((salesOrder.magentoSalesOrderInfo.billing_address.vat_id ?? "").Trim().Length > 0))
				{
					s = Global.digitos(salesOrder.magentoSalesOrderInfo.billing_address.vat_id);
					if (Global.isCnpjCpfOk(s)) cpfCnpjIdentificado = s;
				}

				if ((cpfCnpjIdentificado.Length == 0) && ((salesOrder.magentoSalesOrderInfo.shipping_address.vat_id ?? "").Trim().Length > 0))
				{
					s = Global.digitos(salesOrder.magentoSalesOrderInfo.shipping_address.vat_id);
					if (Global.isCnpjCpfOk(s)) cpfCnpjIdentificado = s;
				}

				salesOrder.cpfCnpjIdentificado = cpfCnpjIdentificado;
				#endregion

				#region [ Pesquisa o BD para verificar se o pedido Magento já foi cadastrado no sistema anteriormente (bloqueia duplicidade) ]
				listaPedidosERP = PedidoDAO.pesquisaPedidoValidoByNumPedidoMagento(numeroPedidoMagento);
				if (listaPedidosERP.Count > 0)
				{
					pedidoERP = PedidoDAO.getPedido(listaPedidosERP[0]);
					salesOrder.erpSalesOrderJaCadastrado.pedido = pedidoERP.pedido_base;
					salesOrder.erpSalesOrderJaCadastrado.vendedor = pedidoERP.vendedor;
					salesOrder.erpSalesOrderJaCadastrado.usuario_cadastro = pedidoERP.usuario_cadastro;
					salesOrder.erpSalesOrderJaCadastrado.dt_cadastro = Global.formataDataYyyyMmDdComSeparador(pedidoERP.data);
					salesOrder.erpSalesOrderJaCadastrado.dt_cadastro_formatado = Global.formataDataDdMmYyyyComSeparador(pedidoERP.data);
					salesOrder.erpSalesOrderJaCadastrado.dt_hr_cadastro = Global.formataDataYyyyMmDdHhMmSsComSeparador(pedidoERP.data_hora);
					salesOrder.erpSalesOrderJaCadastrado.dt_hr_cadastro_formatado = Global.formataDataDdMmYyyyHhMmSsComSeparador(pedidoERP.data_hora);
				}
				#endregion

				#region [ Dados básicos do cliente ]
				if ((salesOrder.cpfCnpjIdentificado ?? "").Trim().Length > 0)
				{
					cliente = ClienteDAO.getClienteByCpfCnpj(salesOrder.cpfCnpjIdentificado);

					#region [ Cliente novo: cadastra automaticamente ]
					if (cliente == null)
					{
						#region [ Verifica o parâmetro que define se o cliente deve ser cadastrado automaticamente ]
						intParametroFlagCadSemiAutoPedMagentoCadastrarAutomaticamenteClienteNovo = GeralDAO.getCampoInteiroTabelaParametro(Global.Cte.Parametros.ID_T_PARAMETRO.FLAG_CAD_SEMI_AUTO_PED_MAGENTO_CADASTRAR_AUTOMATICAMENTE_CLIENTE_NOVO);
						#endregion

						if (intParametroFlagCadSemiAutoPedMagentoCadastrarAutomaticamenteClienteNovo == 1)
						{
							try
							{
								cliente = new Cliente();

								#region [ Preenche dados para cadastro do cliente ]
								cliente.cnpj_cpf = cpfCnpjIdentificado;
								cliente.tipo = (Global.digitos(cpfCnpjIdentificado).Length == 11 ? Global.Cte.TipoPessoa.PF : Global.Cte.TipoPessoa.PJ);
								if ((salesOrder.magentoSalesOrderInfo.billing_address.ie ?? "").Length > 0) cliente.ie = Global.digitos(salesOrder.magentoSalesOrderInfo.billing_address.ie);
								cliente.nome = Global.ecDadosFormataNome(salesOrder.magentoSalesOrderInfo.customer_firstname, salesOrder.magentoSalesOrderInfo.customer_middlename, salesOrder.magentoSalesOrderInfo.customer_lastname, 60);

								if ((salesOrder.magentoSalesOrderInfo.customer_gender ?? "").Length > 0)
								{
									if (salesOrder.magentoSalesOrderInfo.customer_gender.Trim().Equals("1"))
									{
										cliente.sexo = Global.Cte.Sexo.Masculino;
									}
									else if (salesOrder.magentoSalesOrderInfo.customer_gender.Trim().Equals("2"))
									{
										cliente.sexo = Global.Cte.Sexo.Feminino;
									}
								}

								#region [ Endereço ]
								if (cliente.tipo.Equals(Global.Cte.TipoPessoa.PF))
								{
									#region [ Cliente PF ]
									// Cliente PF: usa o endereço de entrega como sendo o único endereço do cliente
									v = (salesOrder.magentoSalesOrderInfo.shipping_address.street ?? "").Split('\n');
									if (v.Length >= 1) cliente.endereco = v[0].Replace('\r', ' ').Trim();
									if (v.Length >= 2) cliente.endereco_numero = v[1].Replace('\r', ' ').Trim();
									if (v.Length >= 3) cliente.endereco_complemento = v[2].Replace('\r', ' ').Trim();
									if (v.Length >= 4) cliente.bairro = v[3].Replace('\r', ' ').Trim();
									cliente.cidade = (salesOrder.magentoSalesOrderInfo.shipping_address.city ?? "").Trim();
									if (Global.isUfOk(salesOrder.magentoSalesOrderInfo.shipping_address.region))
									{
										cliente.uf = salesOrder.magentoSalesOrderInfo.shipping_address.region.Trim();
									}
									else
									{
										cliente.uf = Global.decodificaUfExtensoParaSigla((salesOrder.magentoSalesOrderInfo.shipping_address.region ?? ""));
									}
									cliente.cep = Global.digitos((salesOrder.magentoSalesOrderInfo.shipping_address.postcode ?? ""));
									#endregion
								}
								else
								{
									#region [ Cliente PJ ]
									v = (salesOrder.magentoSalesOrderInfo.billing_address.street ?? "").Split('\n');
									if (v.Length >= 1) cliente.endereco = v[0].Replace('\r', ' ').Trim();
									if (v.Length >= 2) cliente.endereco_numero = v[1].Replace('\r', ' ').Trim();
									if (v.Length >= 3) cliente.endereco_complemento = v[2].Replace('\r', ' ').Trim();
									if (v.Length >= 4) cliente.bairro = v[3].Replace('\r', ' ').Trim();
									cliente.cidade = (salesOrder.magentoSalesOrderInfo.billing_address.city ?? "").Trim();
									if (Global.isUfOk(salesOrder.magentoSalesOrderInfo.billing_address.region))
									{
										cliente.uf = salesOrder.magentoSalesOrderInfo.billing_address.region.Trim();
									}
									else
									{
										cliente.uf = Global.decodificaUfExtensoParaSigla((salesOrder.magentoSalesOrderInfo.billing_address.region ?? ""));
									}
									cliente.cep = Global.digitos((salesOrder.magentoSalesOrderInfo.billing_address.postcode ?? ""));
									#endregion
								}
								#endregion

								#region [ Telefone ]
								if (cliente.tipo.Equals(Global.Cte.TipoPessoa.PF))
								{
									#region [ Telefones para PF ]
									if ((salesOrder.magentoSalesOrderInfo.shipping_address.telephone ?? "").Length > 0)
									{
										if (Global.ecDadosDecodificaTelefoneFormatado((salesOrder.magentoSalesOrderInfo.shipping_address.telephone ?? ""), out sDDD, out sTelefone))
										{
											cliente.ddd_res = Global.digitos(sDDD);
											cliente.tel_res = Global.digitos(sTelefone);
										}
									}

									if ((salesOrder.magentoSalesOrderInfo.shipping_address.celular ?? "").Length > 0)
									{
										if (Global.ecDadosDecodificaTelefoneFormatado((salesOrder.magentoSalesOrderInfo.shipping_address.celular ?? ""), out sDDD, out sTelefone))
										{
											cliente.ddd_cel = Global.digitos(sDDD);
											cliente.tel_cel = Global.digitos(sTelefone);
										}
									}

									if ((salesOrder.magentoSalesOrderInfo.shipping_address.fax ?? "").Length > 0)
									{
										if (Global.ecDadosDecodificaTelefoneFormatado((salesOrder.magentoSalesOrderInfo.shipping_address.fax ?? ""), out sDDD, out sTelefone))
										{
											cliente.ddd_com = Global.digitos(sDDD);
											cliente.tel_com = Global.digitos(sTelefone);
										}
									}
									#endregion
								}
								else
								{
									#region [ Telefones para PJ ]
									if ((salesOrder.magentoSalesOrderInfo.billing_address.telephone ?? "").Length > 0)
									{
										if (Global.ecDadosDecodificaTelefoneFormatado((salesOrder.magentoSalesOrderInfo.billing_address.telephone ?? ""), out sDDD, out sTelefone))
										{
											cliente.ddd_com = Global.digitos(sDDD);
											cliente.tel_com = Global.digitos(sTelefone);
										}
									}

									if ((salesOrder.magentoSalesOrderInfo.billing_address.celular ?? "").Length > 0)
									{
										if (Global.ecDadosDecodificaTelefoneFormatado((salesOrder.magentoSalesOrderInfo.billing_address.celular ?? ""), out sDDD, out sTelefone))
										{
											cliente.ddd_com_2 = Global.digitos(sDDD);
											cliente.tel_com_2 = Global.digitos(sTelefone);
										}
									}

									if ((cliente.tel_com_2.Trim().Length == 0) && ((salesOrder.magentoSalesOrderInfo.billing_address.fax ?? "").Length > 0))
									{
										if (Global.ecDadosDecodificaTelefoneFormatado((salesOrder.magentoSalesOrderInfo.billing_address.fax ?? ""), out sDDD, out sTelefone))
										{
											cliente.ddd_com_2 = Global.digitos(sDDD);
											cliente.tel_com_2 = Global.digitos(sTelefone);
										}
									}
									#endregion
								}
								#endregion

								if ((salesOrder.magentoSalesOrderInfo.customer_dob ?? "").Length > 0) cliente.dt_nasc = Global.converteYyyyMmDdHhMmSsParaDateTime(salesOrder.magentoSalesOrderInfo.customer_dob.Trim());
								cliente.email = salesOrder.magentoSalesOrderInfo.customer_email;

								if (cliente.tipo.Equals(Global.Cte.TipoPessoa.PF)) cliente.produtor_rural_status = Global.Cte.ProdutorRural.COD_ST_CLIENTE_PRODUTOR_RURAL_NAO;

								cliente.sistema_responsavel_cadastro = Global.Cte.SistemaResponsavelCadastro.COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP_WEBAPI;
								cliente.sistema_responsavel_atualizacao = Global.Cte.SistemaResponsavelCadastro.COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP_WEBAPI;
								#endregion

								#region [ Grava cliente no banco de dados ]
								if (!ClienteDAO.insere(cliente, loja, usuario, out msg_erro))
								{
									msg = "Falha ao tentar cadastrar o cliente no banco de dados do sistema!";
									if (msg_erro.Length > 0) msg += "\n" + msg_erro;
									Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + msg);
									throw new Exception(msg);
								}
								#endregion
							}
							catch (Exception ex)
							{
								msg = NOME_DESTA_ROTINA + " - Exception: " + ex.ToString();
								Global.gravaLogAtividade(msg);
								throw new Exception(msg);
							}
						} // if (intParametroFlagCadSemiAutoPedMagentoCadastrarAutomaticamenteClienteNovo == 1)
					} // if (cliente == null)
					#endregion

					if (cliente != null)
					{
						salesOrder.erpCliente.id_cliente = cliente.id;
						salesOrder.erpCliente.cnpj_cpf = cliente.cnpj_cpf;
						salesOrder.erpCliente.nome = cliente.nome;
					}
				}
				#endregion

				#region [ Analisa os dados para tentar identificar se é pedido de marketplace e qual é o nº pedido marketplace ]
				sNumPedidoMktpIdentificado = "";
				sNumPedidoMktpCompletoIdentificado = "";
				sOrigemMktpIdentificado = "";
				listaCodigoDescricao = GeralDAO.getCodigoDescricaoByGrupo(Global.Cte.CodigoDescricao.PedidoECommerce_Origem, Global.eFiltroFlagStInativo.FLAG_IGNORADO, out msg_erro);

				if ((salesOrder.magentoSalesOrderInfo.skyhub_code ?? "").Trim().Length > 0)
				{
					#region [ Tenta identificar o nº pedido marketplace através do campo 'skyhub_code' (ao invés do comentário registrado no status history) ]
					foreach (var codDescr in listaCodigoDescricao)
					{
						// Verifica se o flag está configurado para que seja feito o tratamento usando o campo 'skyhub_info.code' (que foi decodificado para o campo 'skyhub_code')
						if (codDescr.parametro_3_campo_flag == 0) continue;
						sParametro = (codDescr.parametro_4_campo_texto ?? "").Trim();
						if (sParametro.Length == 0) continue;
						vMktpOrderDescriptor = sParametro.Split('|');
						for (int k = 0; k < vMktpOrderDescriptor.Length; k++)
						{
							sMktpOrderDescriptor = vMktpOrderDescriptor[k];
							if ((sMktpOrderDescriptor ?? "").Trim().Length == 0) continue;
							if (salesOrder.magentoSalesOrderInfo.skyhub_code.Trim().ToUpper().StartsWith(sMktpOrderDescriptor.ToUpper()))
							{
								// Obtém a parte relativa ao nº pedido marketplace
								sValue = salesOrder.magentoSalesOrderInfo.skyhub_code.Trim().Substring(sMktpOrderDescriptor.Length).Trim();
								if (sValue.Length > 0)
								{
									sNumPedidoMktpCompletoIdentificado = sValue;

									#region [ Tratamento p/ nº marketplace do Walmart (ex: 78381578-1796973) ]
									if (loja.Equals(Global.Cte.Loja.ArClube) && codDescr.codigo.Equals("009"))
									{
										if (sValue.Contains('-'))
										{
											vValue = sValue.Split('-');
											sValue = vValue[0];
										}
									}
									#endregion

									#region [ Tratamento p/ nº marketplace do Carrefour (ex: 2090221380001-A) ]
									if (loja.Equals(Global.Cte.Loja.ArClube) && codDescr.codigo.Equals("016"))
									{
										if (sValue.Contains('-'))
										{
											vValue = sValue.Split('-');
											sValue = vValue[0];
										}
									}
									#endregion

									sNumPedidoMktpIdentificado = sValue;
									sOrigemMktpIdentificado = codDescr.codigo;
									break;
								}
							}
						}
					}
					#endregion
				}

				#region [ Se não conseguiu identificar o nº pedido marketplace através do campo 'skyhub_code', analisa o status history ]
				if (sNumPedidoMktpIdentificado.Length == 0)
				{
					for (int i = (salesOrder.magentoSalesOrderInfo.status_history.Count - 1); i >= 0; i--)
					{
						statusHistory = salesOrder.magentoSalesOrderInfo.status_history[i];
						if (statusHistory == null) continue;

						sComment = (statusHistory.comment ?? "").Trim();
						if (sComment.Length == 0) continue;

						// Normaliza quebra de linha, se houver, para que sempre seja o \n
						if (sComment.Contains('\r') && sComment.Contains('\n'))
						{
							sComment = sComment.Replace("\r", String.Empty);
						}
						else if (sComment.Contains('\r') && (!sComment.Contains('\n')))
						{
							sComment = sComment.Replace('\r', '\n');
						}
						// Tratamento caso a quebra de linha seja através de tag HTML
						sComment = sComment.Replace("<br>", "\n");
						sComment = sComment.Replace("<br />", "\n");
						sComment = sComment.Replace("<br/>", "\n");
						sComment = sComment.Replace("<BR>", "\n");
						sComment = sComment.Replace("<BR />", "\n");
						sComment = sComment.Replace("<BR/>", "\n");

						vComment = sComment.Split('\n');
						for (int j = 0; j < vComment.Length; j++)
						{
							if (vComment[j].Trim().Length == 0) continue;

							foreach (var codDescr in listaCodigoDescricao)
							{
								if ((codDescr.parametro_2_campo_flag == 1) && ((codDescr.parametro_3_campo_texto ?? "").Trim().Length > 0))
								{
									#region [ É necessário verificar se o identificador especificado em parametro_3_campo_texto está presente no texto do comentário ]
									if (!sComment.Contains((codDescr.parametro_3_campo_texto ?? "").Trim())) continue;
									#endregion
								}

								sParametro = (codDescr.parametro_campo_texto ?? "").Trim();
								if (sParametro.Length == 0) continue;
								vMktpOrderDescriptor = sParametro.Split('|');
								for (int k = 0; k < vMktpOrderDescriptor.Length; k++)
								{
									sMktpOrderDescriptor = vMktpOrderDescriptor[k];
									if ((sMktpOrderDescriptor ?? "").Trim().Length == 0) continue;
									if (vComment[j].ToUpper().StartsWith(sMktpOrderDescriptor.ToUpper()))
									{
										// Obtém a parte relativa ao nº pedido marketplace
										sValue = vComment[j].Substring(sMktpOrderDescriptor.Length).Trim();
										if (sValue.Length > 0)
										{
											sNumPedidoMktpCompletoIdentificado = sValue;

											#region [ Tratamento p/ nº marketplace do Walmart (ex: 78381578-1796973) ]
											if (loja.Equals(Global.Cte.Loja.ArClube) && codDescr.codigo.Equals("009"))
											{
												if (sValue.Contains('-'))
												{
													vValue = sValue.Split('-');
													sValue = vValue[0];
												}
											}
											#endregion

											#region [ Tratamento p/ nº marketplace do Carrefour (ex: 2090221380001-A) ]
											if (loja.Equals(Global.Cte.Loja.ArClube) && codDescr.codigo.Equals("016"))
											{
												if (sValue.Contains('-'))
												{
													vValue = sValue.Split('-');
													sValue = vValue[0];
												}
											}
											#endregion

											#region [ Tratamento p/ nº marketplace do Leroy Merlin (ex: 0004536570-A) ]
											if (loja.Equals(Global.Cte.Loja.ArClube) && codDescr.codigo.Equals("017"))
											{
												if (sValue.Contains('-'))
												{
													vValue = sValue.Split('-');
													sValue = vValue[0];
												}
											}
											#endregion


											sNumPedidoMktpIdentificado = sValue;
											sOrigemMktpIdentificado = codDescr.codigo;
											break;
										}
									}
								}
								if (sNumPedidoMktpIdentificado.Length > 0) break;
							}
							if (sNumPedidoMktpIdentificado.Length > 0) break;
						}
						if (sNumPedidoMktpIdentificado.Length > 0) break;
					}
				}
				#endregion
				#endregion

				#region [ Grava o XML do pedido no BD ]
				insertPedidoXml = new MagentoErpPedidoXml();
				insertPedidoXml.pedido_magento = numeroPedidoMagento;
				insertPedidoXml.pedido_erp = (salesOrder.erpSalesOrderJaCadastrado.pedido ?? "");
				insertPedidoXml.operationControlTicket = operationControlTicket;
				insertPedidoXml.loja = loja;
				insertPedidoXml.usuario_cadastro = usuario;
				insertPedidoXml.magento_api_versao = Global.Cte.MagentoApiIntegracao.VERSAO_API_MAGENTO_V2_REST_JSON;
				insertPedidoXml.pedido_json = sJson;
				insertPedidoXml.cpfCnpjIdentificado = cpfCnpjIdentificado;
				insertPedidoXml.increment_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.increment_id ?? ""));
				insertPedidoXml.created_at = (salesOrder.magentoSalesOrderInfo.created_at ?? "");
				insertPedidoXml.updated_at = (salesOrder.magentoSalesOrderInfo.updated_at ?? "");
				insertPedidoXml.customer_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.customer_id ?? ""));
				insertPedidoXml.billing_address_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.billing_address_id ?? ""));
				insertPedidoXml.shipping_address_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.shipping_address_id ?? ""));
				insertPedidoXml.status = (salesOrder.magentoSalesOrderInfo.status ?? "");
				insertPedidoXml.status_descricao = Global.retornaEcDescricaoStatus(insertPedidoXml.status, loja);
				insertPedidoXml.state = (salesOrder.magentoSalesOrderInfo.state ?? "");
				insertPedidoXml.state_descricao = Global.retornaEcDescricaoState(insertPedidoXml.state, loja);
				insertPedidoXml.customer_email = (salesOrder.magentoSalesOrderInfo.customer_email ?? "");
				insertPedidoXml.customer_firstname = (salesOrder.magentoSalesOrderInfo.customer_firstname ?? "");
				insertPedidoXml.customer_lastname = (salesOrder.magentoSalesOrderInfo.customer_lastname ?? "");
				insertPedidoXml.customer_middlename = (salesOrder.magentoSalesOrderInfo.customer_middlename ?? "");
				insertPedidoXml.quote_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.quote_id ?? ""));
				insertPedidoXml.customer_group_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.customer_group_id ?? ""));
				insertPedidoXml.order_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.order_id ?? ""));
				insertPedidoXml.customer_dob = (salesOrder.magentoSalesOrderInfo.customer_dob ?? "");
				insertPedidoXml.clearsale_status_code = (salesOrder.magentoSalesOrderInfo.clearsale_status_code ?? "");
				insertPedidoXml.clearSale_status = (salesOrder.magentoSalesOrderInfo.clearSale_status ?? "");
				insertPedidoXml.clearSale_score = (salesOrder.magentoSalesOrderInfo.clearSale_score ?? "");
				insertPedidoXml.clearSale_packageID = (salesOrder.magentoSalesOrderInfo.clearSale_packageID ?? "");
				insertPedidoXml.shipping_amount = Global.converteNumeroDecimal((salesOrder.magentoSalesOrderInfo.shipping_amount ?? ""));
				insertPedidoXml.shipping_discount_amount = Global.converteNumeroDecimal((salesOrder.magentoSalesOrderInfo.shipping_discount_amount ?? ""));
				insertPedidoXml.discount_amount = Global.converteNumeroDecimal((salesOrder.magentoSalesOrderInfo.discount_amount ?? ""));
				insertPedidoXml.subtotal = Global.converteNumeroDecimal((salesOrder.magentoSalesOrderInfo.subtotal ?? ""));
				insertPedidoXml.grand_total = Global.converteNumeroDecimal((salesOrder.magentoSalesOrderInfo.grand_total ?? ""));
				insertPedidoXml.installer_document = (salesOrder.magentoSalesOrderInfo.installer_document ?? "");
				insertPedidoXml.installer_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.installer_id ?? ""));
				insertPedidoXml.commission_value = Global.converteNumeroDecimal((salesOrder.magentoSalesOrderInfo.commission_value ?? ""));
				insertPedidoXml.commission_discount = Global.converteNumeroDecimal((salesOrder.magentoSalesOrderInfo.commission_discount ?? ""));
				insertPedidoXml.commission_final_discount = Global.converteNumeroDecimal((salesOrder.magentoSalesOrderInfo.commission_final_discount ?? ""));
				insertPedidoXml.commission_final_value = Global.converteNumeroDecimal((salesOrder.magentoSalesOrderInfo.commission_final_value ?? ""));
				insertPedidoXml.commission_discount_type = (salesOrder.magentoSalesOrderInfo.commission_discount_type ?? "");

				if (sNumPedidoMktpIdentificado.Length > 0)
				{
					insertPedidoXml.pedido_marketplace = sNumPedidoMktpIdentificado;
					insertPedidoXml.pedido_marketplace_completo = sNumPedidoMktpCompletoIdentificado;
					insertPedidoXml.marketplace_codigo_origem = sOrigemMktpIdentificado;
				}

				blnInserted = MagentoApiDAO.insertMagentoPedidoXml(insertPedidoXml, out msg_erro);
				if (!blnInserted)
				{
					msg = "Falha ao tentar gravar no BD os dados do pedido Magento!";
					if (msg_erro.Length > 0) msg += "\n" + msg_erro;
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + msg);
					throw new Exception(msg);
				}
				#endregion
			}
			#endregion

			#region [ Grava os dados decodificados no BD ]
			if (blnInserted)
			{
				#region [ Endereço de cobrança ]
				decodeEndereco = new MagentoErpPedidoXmlDecodeEndereco();
				decodeEndereco.id_magento_api_pedido_xml = insertPedidoXml.id;
				decodeEndereco.tipo_endereco = Global.Cte.MagentoSoapApi.TIPO_ENDERECO__COBRANCA;
				v = (salesOrder.magentoSalesOrderInfo.billing_address.street ?? "").Split('\n');
				if (v.Length >= 1) decodeEndereco.endereco = v[0].Replace('\r', ' ').Trim();
				if (v.Length >= 2) decodeEndereco.endereco_numero = v[1].Replace('\r', ' ').Trim();
				if (v.Length >= 3) decodeEndereco.endereco_complemento = v[2].Replace('\r', ' ').Trim();
				if (v.Length >= 4) decodeEndereco.bairro = v[3].Replace('\r', ' ').Trim();
				decodeEndereco.cidade = (salesOrder.magentoSalesOrderInfo.billing_address.city ?? "");
				if (Global.isUfOk(salesOrder.magentoSalesOrderInfo.billing_address.region))
				{
					decodeEndereco.uf = salesOrder.magentoSalesOrderInfo.billing_address.region;
				}
				else
				{
					decodeEndereco.uf = Global.decodificaUfExtensoParaSigla((salesOrder.magentoSalesOrderInfo.billing_address.region ?? ""));
				}
				decodeEndereco.cep = Global.digitos((salesOrder.magentoSalesOrderInfo.billing_address.postcode ?? ""));
				decodeEndereco.address_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.billing_address.address_id ?? ""));
				decodeEndereco.parent_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.billing_address.parent_id ?? ""));
				decodeEndereco.customer_address_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.billing_address.customer_address_id ?? ""));
				decodeEndereco.quote_address_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.billing_address.quote_address_id ?? ""));
				decodeEndereco.region_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.billing_address.region_id ?? ""));
				decodeEndereco.address_type = (salesOrder.magentoSalesOrderInfo.billing_address.address_type ?? "");
				decodeEndereco.street = (salesOrder.magentoSalesOrderInfo.billing_address.street ?? "");
				decodeEndereco.city = (salesOrder.magentoSalesOrderInfo.billing_address.city ?? "");
				decodeEndereco.region = (salesOrder.magentoSalesOrderInfo.billing_address.region ?? "");
				decodeEndereco.postcode = (salesOrder.magentoSalesOrderInfo.billing_address.postcode ?? "");
				decodeEndereco.country_id = (salesOrder.magentoSalesOrderInfo.billing_address.country_id ?? "");
				decodeEndereco.firstname = (salesOrder.magentoSalesOrderInfo.billing_address.firstname ?? "");
				decodeEndereco.middlename = (salesOrder.magentoSalesOrderInfo.billing_address.middlename ?? "");
				decodeEndereco.lastname = (salesOrder.magentoSalesOrderInfo.billing_address.lastname ?? "");
				decodeEndereco.email = (salesOrder.magentoSalesOrderInfo.billing_address.email ?? "");
				// Retorna os telefones somente se não forem fictícios ou se não estiverem com os dígitos mascarados
				if (Global.ecDadosDecodificaTelefoneFormatado((salesOrder.magentoSalesOrderInfo.billing_address.telephone ?? ""), out sDDD, out sTelefone)) decodeEndereco.telephone = (salesOrder.magentoSalesOrderInfo.billing_address.telephone ?? "");
				if (Global.ecDadosDecodificaTelefoneFormatado((salesOrder.magentoSalesOrderInfo.billing_address.celular ?? ""), out sDDD, out sTelefone)) decodeEndereco.celular = (salesOrder.magentoSalesOrderInfo.billing_address.celular ?? "");
				if (Global.ecDadosDecodificaTelefoneFormatado((salesOrder.magentoSalesOrderInfo.billing_address.fax ?? ""), out sDDD, out sTelefone)) decodeEndereco.fax = (salesOrder.magentoSalesOrderInfo.billing_address.fax ?? "");
				decodeEndereco.tipopessoa = (salesOrder.magentoSalesOrderInfo.billing_address.tipopessoa ?? "");
				decodeEndereco.rg = (salesOrder.magentoSalesOrderInfo.billing_address.rg ?? "");
				decodeEndereco.ie = (salesOrder.magentoSalesOrderInfo.billing_address.ie ?? "");
				decodeEndereco.cpfcnpj = (salesOrder.magentoSalesOrderInfo.billing_address.cpfcnpj ?? "");
				decodeEndereco.empresa = (salesOrder.magentoSalesOrderInfo.billing_address.empresa ?? "");
				decodeEndereco.nomefantasia = (salesOrder.magentoSalesOrderInfo.billing_address.nomefantasia ?? "");
				decodeEndereco.street_detail = (salesOrder.magentoSalesOrderInfo.billing_address.street_detail ?? "");
				if (!MagentoApiDAO.insertMagentoPedidoXmlDecodeEndereco(decodeEndereco, out msg_erro))
				{
					msg = "Falha ao tentar gravar no BD os dados do endereço de cobrança!";
					if (msg_erro.Length > 0) msg += "\n" + msg_erro;
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + msg);
					throw new Exception(msg);
				}
				#endregion

				#region [ Endereço de entrega ]
				decodeEndereco = new MagentoErpPedidoXmlDecodeEndereco();
				decodeEndereco.id_magento_api_pedido_xml = insertPedidoXml.id;
				decodeEndereco.tipo_endereco = Global.Cte.MagentoSoapApi.TIPO_ENDERECO__ENTREGA;
				v = (salesOrder.magentoSalesOrderInfo.shipping_address.street ?? "").Split('\n');
				if (v.Length >= 1) decodeEndereco.endereco = v[0].Replace('\r', ' ').Trim();
				if (v.Length >= 2) decodeEndereco.endereco_numero = v[1].Replace('\r', ' ').Trim();
				if (v.Length >= 3) decodeEndereco.endereco_complemento = v[2].Replace('\r', ' ').Trim();
				if (v.Length >= 4) decodeEndereco.bairro = v[3].Replace('\r', ' ').Trim();
				decodeEndereco.cidade = (salesOrder.magentoSalesOrderInfo.shipping_address.city ?? "");
				if (Global.isUfOk(salesOrder.magentoSalesOrderInfo.shipping_address.region))
				{
					decodeEndereco.uf = salesOrder.magentoSalesOrderInfo.shipping_address.region;
				}
				else
				{
					decodeEndereco.uf = Global.decodificaUfExtensoParaSigla((salesOrder.magentoSalesOrderInfo.shipping_address.region ?? ""));
				}
				decodeEndereco.cep = Global.digitos((salesOrder.magentoSalesOrderInfo.shipping_address.postcode ?? ""));
				decodeEndereco.address_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.shipping_address.address_id ?? ""));
				decodeEndereco.parent_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.shipping_address.parent_id ?? ""));
				decodeEndereco.customer_address_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.shipping_address.customer_address_id ?? ""));
				decodeEndereco.quote_address_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.shipping_address.quote_address_id ?? ""));
				decodeEndereco.region_id = (int)Global.converteInteiro((salesOrder.magentoSalesOrderInfo.shipping_address.region_id ?? ""));
				decodeEndereco.address_type = (salesOrder.magentoSalesOrderInfo.shipping_address.address_type ?? "");
				decodeEndereco.street = (salesOrder.magentoSalesOrderInfo.shipping_address.street ?? "");
				decodeEndereco.city = (salesOrder.magentoSalesOrderInfo.shipping_address.city ?? "");
				decodeEndereco.region = (salesOrder.magentoSalesOrderInfo.shipping_address.region ?? "");
				decodeEndereco.postcode = (salesOrder.magentoSalesOrderInfo.shipping_address.postcode ?? "");
				decodeEndereco.country_id = (salesOrder.magentoSalesOrderInfo.shipping_address.country_id ?? "");
				decodeEndereco.firstname = (salesOrder.magentoSalesOrderInfo.shipping_address.firstname ?? "");
				decodeEndereco.middlename = (salesOrder.magentoSalesOrderInfo.shipping_address.middlename ?? "");
				decodeEndereco.lastname = (salesOrder.magentoSalesOrderInfo.shipping_address.lastname ?? "");
				decodeEndereco.email = (salesOrder.magentoSalesOrderInfo.shipping_address.email ?? "");
				// Retorna os telefones somente se não forem fictícios ou se não estiverem com os dígitos mascarados
				if (Global.ecDadosDecodificaTelefoneFormatado((salesOrder.magentoSalesOrderInfo.shipping_address.telephone ?? ""), out sDDD, out sTelefone)) decodeEndereco.telephone = (salesOrder.magentoSalesOrderInfo.shipping_address.telephone ?? "");
				if (Global.ecDadosDecodificaTelefoneFormatado((salesOrder.magentoSalesOrderInfo.shipping_address.celular ?? ""), out sDDD, out sTelefone)) decodeEndereco.celular = (salesOrder.magentoSalesOrderInfo.shipping_address.celular ?? "");
				if (Global.ecDadosDecodificaTelefoneFormatado((salesOrder.magentoSalesOrderInfo.shipping_address.fax ?? ""), out sDDD, out sTelefone)) decodeEndereco.fax = (salesOrder.magentoSalesOrderInfo.shipping_address.fax ?? "");
				decodeEndereco.tipopessoa = (salesOrder.magentoSalesOrderInfo.shipping_address.tipopessoa ?? "");
				decodeEndereco.rg = (salesOrder.magentoSalesOrderInfo.shipping_address.rg ?? "");
				decodeEndereco.ie = (salesOrder.magentoSalesOrderInfo.shipping_address.ie ?? "");
				decodeEndereco.cpfcnpj = (salesOrder.magentoSalesOrderInfo.shipping_address.cpfcnpj ?? "");
				decodeEndereco.empresa = (salesOrder.magentoSalesOrderInfo.shipping_address.empresa ?? "");
				decodeEndereco.nomefantasia = (salesOrder.magentoSalesOrderInfo.shipping_address.nomefantasia ?? "");
				decodeEndereco.street_detail = (salesOrder.magentoSalesOrderInfo.shipping_address.street_detail ?? "");
				if (!MagentoApiDAO.insertMagentoPedidoXmlDecodeEndereco(decodeEndereco, out msg_erro))
				{
					msg = "Falha ao tentar gravar no BD os dados do endereço de entrega!";
					if (msg_erro.Length > 0) msg += "\n" + msg_erro;
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + msg);
					throw new Exception(msg);
				}
				#endregion

				#region [ Itens (produtos) ]
				foreach (MagentoSoapApiSalesOrderItem item in salesOrder.magentoSalesOrderInfo.items)
				{
					decodeItem = new MagentoErpPedidoXmlDecodeItem();
					decodeItem.id_magento_api_pedido_xml = insertPedidoXml.id;
					decodeItem.sku = (item.sku ?? "");
					decodeItem.qty_ordered = Global.converteNumeroDecimal((item.qty_ordered ?? ""));
					decodeItem.product_id = (int)Global.converteInteiro((item.product_id ?? ""));
					decodeItem.item_id = (int)Global.converteInteiro((item.item_id ?? ""));
					decodeItem.order_id = (int)Global.converteInteiro((item.order_id ?? ""));
					decodeItem.quote_item_id = (int)Global.converteInteiro((item.quote_item_id ?? ""));
					decodeItem.price = Global.converteNumeroDecimal((item.price ?? ""));
					decodeItem.base_price = Global.converteNumeroDecimal((item.base_price ?? ""));
					decodeItem.original_price = Global.converteNumeroDecimal((item.original_price ?? ""));
					decodeItem.base_original_price = Global.converteNumeroDecimal((item.base_original_price ?? ""));
					decodeItem.discount_percent = Global.converteNumeroDecimal((item.discount_percent ?? ""));
					decodeItem.discount_amount = Global.converteNumeroDecimal((item.discount_amount ?? ""));
					decodeItem.base_discount_amount = Global.converteNumeroDecimal((item.base_discount_amount ?? ""));
					decodeItem.name = (item.name ?? "");
					decodeItem.product_type = (item.product_type ?? "");
					decodeItem.has_children = (item.has_children ?? "");
					decodeItem.parent_item_id = (int)Global.converteInteiro((item.parent_item_id ?? ""));
					decodeItem.weight = Global.converteDouble(item.weight, SEPARADOR_DECIMAL_NUM_REAL);
					decodeItem.is_virtual = (int)Global.converteInteiro((item.is_virtual ?? ""));
					decodeItem.free_shipping = (int)Global.converteInteiro((item.free_shipping ?? ""));
					decodeItem.is_qty_decimal = (int)Global.converteInteiro((item.is_qty_decimal ?? ""));
					decodeItem.no_discount = (int)Global.converteInteiro((item.no_discount ?? ""));
					decodeItem.qty_canceled = Global.converteNumeroDecimal((item.qty_canceled ?? ""));
					decodeItem.qty_invoiced = Global.converteNumeroDecimal((item.qty_invoiced ?? ""));
					decodeItem.qty_refunded = Global.converteNumeroDecimal((item.qty_refunded ?? ""));
					decodeItem.qty_shipped = Global.converteNumeroDecimal((item.qty_shipped ?? ""));
					decodeItem.tax_percent = Global.converteDouble(item.tax_percent, SEPARADOR_DECIMAL_NUM_REAL);
					decodeItem.tax_amount = Global.converteNumeroDecimal((item.tax_amount ?? ""));
					decodeItem.base_tax_amount = Global.converteNumeroDecimal((item.base_tax_amount ?? ""));
					decodeItem.tax_invoiced = Global.converteNumeroDecimal((item.tax_invoiced ?? ""));
					decodeItem.base_tax_invoiced = Global.converteNumeroDecimal((item.base_tax_invoiced ?? ""));
					decodeItem.discount_invoiced = Global.converteNumeroDecimal((item.discount_invoiced ?? ""));
					decodeItem.base_discount_invoiced = Global.converteNumeroDecimal((item.base_discount_invoiced ?? ""));
					decodeItem.amount_refunded = Global.converteNumeroDecimal((item.amount_refunded ?? ""));
					decodeItem.base_amount_refunded = Global.converteNumeroDecimal((item.base_amount_refunded ?? ""));
					decodeItem.row_total = Global.converteNumeroDecimal((item.row_total ?? ""));
					decodeItem.base_row_total = Global.converteNumeroDecimal((item.base_row_total ?? ""));
					decodeItem.row_invoiced = Global.converteNumeroDecimal((item.row_invoiced ?? ""));
					decodeItem.base_row_invoiced = Global.converteNumeroDecimal((item.base_row_invoiced ?? ""));
					decodeItem.row_weight = Global.converteDouble(item.row_weight, SEPARADOR_DECIMAL_NUM_REAL);
					decodeItem.price_incl_tax = Global.converteNumeroDecimal((item.price_incl_tax ?? ""));
					decodeItem.base_price_incl_tax = Global.converteNumeroDecimal((item.base_price_incl_tax ?? ""));
					decodeItem.row_total_incl_tax = Global.converteNumeroDecimal((item.row_total_incl_tax ?? ""));
					decodeItem.base_row_total_incl_tax = Global.converteNumeroDecimal((item.base_row_total_incl_tax ?? ""));

					if (!MagentoApiDAO.insertMagentoPedidoXmlDecodeItem(decodeItem, out msg_erro))
					{
						msg = "Falha ao tentar gravar no BD os dados do item do pedido (sku=" + decodeItem.sku + ")!";
						if (msg_erro.Length > 0) msg += "\n" + msg_erro;
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + msg);
						throw new Exception(msg);
					}
				}
				#endregion

				#region [ Status History ]
				foreach (MagentoSoapApiStatusHistory item in salesOrder.magentoSalesOrderInfo.status_history)
				{
					decodeStatusHistory = new MagentoErpPedidoXmlDecodeStatusHistory();
					decodeStatusHistory.id_magento_api_pedido_xml = insertPedidoXml.id;
					decodeStatusHistory.parent_id = (int)Global.converteInteiro((item.parent_id ?? ""));
					decodeStatusHistory.is_customer_notified = (byte)Global.converteInteiro((item.is_customer_notified ?? ""));
					decodeStatusHistory.is_visible_on_front = (byte)Global.converteInteiro((item.is_visible_on_front ?? ""));
					decodeStatusHistory.comment = (item.comment ?? "");
					decodeStatusHistory.status = (item.status ?? "");
					decodeStatusHistory.created_at = (item.created_at ?? "");
					decodeStatusHistory.entity_name = (item.entity_name ?? "");
					decodeStatusHistory.store_id = (int)Global.converteInteiro((item.store_id ?? ""));
					if (!MagentoApiDAO.insertMagentoPedidoXmlDecodeStatusHistory(decodeStatusHistory, out msg_erro))
					{
						msg = "Falha ao tentar gravar no BD os dados do status history do pedido (created_at=" + decodeStatusHistory.created_at + ")!";
						if (msg_erro.Length > 0) msg += "\n" + msg_erro;
						Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - " + msg);
						throw new Exception(msg);
					}
				}
				#endregion
			}
			#endregion

			#region [ Converte objeto em dados JSON ]
			var serializer = new JavaScriptSerializer();
			var serializedResult = serializer.Serialize(salesOrder);
			#endregion

			return serializedResult;
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
}