using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Web;
using ART3WebAPI.Models.Entities;
using System.Xml;

namespace ART3WebAPI.Models.Domains
{
	public static class MagentoSoapApi
	{
		#region[ ReaderWriterLock ]
		// Para garantir que os acessos à API do Magento sejam thread-safe
		public static ReaderWriterLock rwlMagentoSoapApi = new ReaderWriterLock();
		#endregion

		#region [ enviaRequisicaoComRetry ]
		/// <summary>
		/// Método que executa o enviaRequisicao() dentro de um laço de tentativas até que a execução seja bem sucedida ou a quantidade máxima de tentativas seja atingida.
		/// Importante: este método pode ser utilizado livremente para requisições de consulta, entretanto, para requisições que alteram dados é importante avaliar antes
		/// as possíveis consequências que podem ocorrer no caso da requisição ter sido processada no web service e o erro ter ocorrido em algum estágio posterior durante
		/// o recebimento da resposta. Nesse caso, o uso deste método pode causar múltiplas execuções da requisição.
		/// </summary>
		/// <param name="xmlReqSoap"></param>
		/// <param name="trxParam"></param>
		/// <param name="urlWebService"></param>
		/// <param name="xmlRespSoap"></param>
		/// <param name="msg_erro"></param>
		/// <returns></returns>
		public static bool enviaRequisicaoComRetry(string xmlReqSoap, Global.Cte.MagentoSoapApi.Transacao trxParam, string urlWebService, out string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			const int MAX_TENTATIVAS = 5;
			int qtdeTentativasRealizadas = 0;
			bool blnResposta;
			#endregion

			do
			{
				qtdeTentativasRealizadas++;

				blnResposta = enviaRequisicao(xmlReqSoap, trxParam, urlWebService, out xmlRespSoap, out msg_erro);
				if (blnResposta) break;

				Thread.Sleep(1000);
			} while (qtdeTentativasRealizadas < MAX_TENTATIVAS);

			return blnResposta;
		}
		#endregion

		#region [ enviaRequisicao ]
		public static bool enviaRequisicao(string xmlReqSoap, Global.Cte.MagentoSoapApi.Transacao trxParam, string urlWebService, out string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "enviaRequisicao()";
			string strMsg;
			HttpWebRequest req;
			HttpWebResponse resp;
			#endregion

			xmlRespSoap = "";
			msg_erro = "";

			try
			{
				strMsg = NOME_DESTA_ROTINA + " - TX\n" + xmlReqSoap;
				Global.gravaLogAtividade(strMsg);

				req = (HttpWebRequest)WebRequest.Create(urlWebService);
				// The Timeout applies to the entire request and response, not individually to the GetRequestStream and GetResponse method calls
				req.Timeout = Global.Cte.MagentoSoapApi.REQUEST_TIMEOUT_EM_MS;
				req.Headers.Add("SOAPAction", trxParam.GetSoapAction());
				req.ContentType = "text/xml;charset=\"utf-8\"";
				req.Method = "POST";
				using (Stream reqStm = req.GetRequestStream())
				{
					using (StreamWriter reqStmW = new StreamWriter(reqStm))
					{
						reqStmW.Write(xmlReqSoap);
					}
				}

				resp = (HttpWebResponse)req.GetResponse();
				using (Stream respStm = resp.GetResponseStream())
				{
					using (StreamReader respStmR = new StreamReader(respStm, Encoding.UTF8))
					{
						xmlRespSoap = respStmR.ReadToEnd();
					}
				}

				strMsg = NOME_DESTA_ROTINA + " - RX\n" + xmlRespSoap;
				Global.gravaLogAtividade(strMsg);

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ montaRequisicaoLogin ]
		public static string montaRequisicaoLogin(string userName, string password)
		{
			string xmlRequisicaoSoap;

			xmlRequisicaoSoap = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
								"<soapenv:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:urn=\"urn: Magento\">" +
								"<soapenv:Header/>" +
								"<soapenv:Body>" +
								"<urn:login soapenv:encodingStyle=\"http://schemas.xmlsoap.org/soap/encoding/\">" +
								"<username xsi:type=\"xsd: string\">" + userName + "</username>" +
								"<apiKey xsi:type=\"xsd: string\">" + password + "</apiKey>" +
								"</urn:login>" +
								"</soapenv:Body>" +
								"</soapenv:Envelope>";

			return xmlRequisicaoSoap;
		}
		#endregion

		#region [ montaRequisicaoEndSession ]
		public static string montaRequisicaoEndSession(string sessionId)
		{
			string xmlRequisicaoSoap;

			xmlRequisicaoSoap = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
								"<soapenv:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:urn=\"urn: Magento\">" +
								"<soapenv:Header/>" +
								"<soapenv:Body>" +
								"<urn:endSession soapenv:encodingStyle=\"http://schemas.xmlsoap.org/soap/encoding/\">" +
								"<sessionId xsi:type=\"xsd: string\">" + sessionId + "</sessionId>" +
								"</urn:endSession>" +
								"</soapenv:Body>" +
								"</soapenv:Envelope>";

			return xmlRequisicaoSoap;
		}
		#endregion

		#region [ obtemSessionIdFromLoginResponse ]
		public static string obtemSessionIdFromLoginResponse(string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			string strValue;
			string sessionId = "";
			XmlDocument xmlDoc;
			XmlNamespaceManager nsmgr;
			XmlNode xmlNode;
			#endregion

			msg_erro = "";
			try
			{
				if ((xmlRespSoap ?? "").Trim().Length == 0) return null;

				xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(xmlRespSoap);
				nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
				nsmgr.AddNamespace("ns1", "urn:Magento");
				xmlNode = xmlDoc.SelectSingleNode("//ns1:loginResponse", nsmgr);
				strValue = Global.obtemXmlChildNodeValue(xmlNode, "loginReturn");
				sessionId = (strValue ?? "");

				return sessionId;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				return null;
			}
		}
		#endregion

		#region [ montaRequisicaoCallSalesOrderInfo ]
		public static string montaRequisicaoCallSalesOrderInfo(string sessionId, string orderIncrementId)
		{
			string xmlRequisicaoSoap;

			xmlRequisicaoSoap = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
								"<soapenv:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:urn=\"urn: Magento\">" +
								"<soapenv:Header/>" +
								"<soapenv:Body>" +
								"<urn:call soapenv:encodingStyle=\"http://schemas.xmlsoap.org/soap/encoding/\">" +
								"<sessionId xsi:type=\"xsd: string\">" + sessionId + "</sessionId>" +
								"<resourcePath xsi:type=\"xsd: string\">sales_order.info</resourcePath>" +
								"<args xsi:type=\"xsd: anyType\">" + orderIncrementId + "</args>" +
								"</urn:call>" +
								"</soapenv:Body>" +
								"</soapenv:Envelope>";

			return xmlRequisicaoSoap;
		}
		#endregion

		#region [ decodificaXmlSalesOrderInfoResponse ]
		public static MagentoSoapApiSalesOrderInfo decodificaXmlSalesOrderInfoResponse(string xmlRespSoap, out string msg_erro)
		{
			string strKey;
			string strValue;
			XmlDocument xmlDoc;
			XmlNamespaceManager nsmgr;
			XmlNodeList xmlNodeListN1;
			MagentoSoapApiSalesOrderInfo orderInfo = new MagentoSoapApiSalesOrderInfo();
			MagentoSoapApiSalesOrderItem orderItem;
			MagentoSoapApiStatusHistory statusHistory;

			msg_erro = "";

			try
			{
				if ((xmlRespSoap ?? "").Trim().Length == 0) return null;

				#region [ Decodifica resposta ]
				xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(xmlRespSoap);
				nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
				nsmgr.AddNamespace("ns1", "urn:Magento");
				xmlNodeListN1 = xmlDoc.SelectNodes("//ns1:callResponse/callReturn/item", nsmgr);
				if (xmlNodeListN1 != null)
				{
					foreach (XmlNode node in xmlNodeListN1)
					{
						strKey = (node["key"].InnerText ?? "");
						switch (strKey)
						{
							case "shipping_address":
								#region [ Campos internos do nó 'shipping_address' ]
								foreach (XmlNode nodeN2 in node["value"].ChildNodes)
								{
									strKey = nodeN2["key"].InnerText;
									strValue = (nodeN2["value"].InnerText ?? "");
									switch (strKey)
									{
										case "parent_id":
											orderInfo.shipping_address.parent_id = strValue;
											break;
										case "customer_address_id":
											orderInfo.shipping_address.customer_address_id = strValue;
											break;
										case "quote_address_id":
											orderInfo.shipping_address.quote_address_id = strValue;
											break;
										case "region_id":
											orderInfo.shipping_address.region_id = strValue;
											break;
										case "customer_id":
											orderInfo.shipping_address.customer_id = strValue;
											break;
										case "fax":
											orderInfo.shipping_address.fax = strValue;
											break;
										case "region":
											orderInfo.shipping_address.region = strValue;
											break;
										case "postcode":
											orderInfo.shipping_address.postcode = strValue;
											break;
										case "firstname":
											orderInfo.shipping_address.firstname = strValue;
											break;
										case "middlename":
											orderInfo.shipping_address.middlename = strValue;
											break;
										case "lastname":
											orderInfo.shipping_address.lastname = strValue;
											break;
										case "street":
											orderInfo.shipping_address.street = strValue;
											break;
										case "city":
											orderInfo.shipping_address.city = strValue;
											break;
										case "email":
											orderInfo.shipping_address.email = strValue;
											break;
										case "telephone":
											orderInfo.shipping_address.telephone = strValue;
											break;
										case "country_id":
											orderInfo.shipping_address.country_id = strValue;
											break;
										case "address_type":
											orderInfo.shipping_address.address_type = strValue;
											break;
										case "prefix":
											orderInfo.shipping_address.prefix = strValue;
											break;
										case "suffix":
											orderInfo.shipping_address.suffix = strValue;
											break;
										case "company":
											orderInfo.shipping_address.company = strValue;
											break;
										case "vat_id":
											orderInfo.shipping_address.vat_id = strValue;
											break;
										case "vat_is_valid":
											orderInfo.shipping_address.vat_is_valid = strValue;
											break;
										case "vat_request_id":
											orderInfo.shipping_address.vat_request_id = strValue;
											break;
										case "vat_request_date":
											orderInfo.shipping_address.vat_request_date = strValue;
											break;
										case "vat_request_success":
											orderInfo.shipping_address.vat_request_success = strValue;
											break;
										case "tipopessoa":
											orderInfo.shipping_address.tipopessoa = strValue;
											break;
										case "rg":
											orderInfo.shipping_address.rg = strValue;
											break;
										case "ie":
											orderInfo.shipping_address.ie = strValue;
											break;
										case "cpfcnpj":
											orderInfo.shipping_address.cpfcnpj = strValue;
											break;
										case "celular":
											orderInfo.shipping_address.celular = strValue;
											break;
										case "empresa":
											orderInfo.shipping_address.empresa = strValue;
											break;
										case "nomefantasia":
											orderInfo.shipping_address.nomefantasia = strValue;
											break;
										case "cpf":
											orderInfo.shipping_address.cpf = strValue;
											break;
										case "address_id":
											orderInfo.shipping_address.address_id = strValue;
											break;
										default:
											orderInfo.shipping_address.UnknownFields.Add(new KeyValuePair<string, string>(strKey, strValue));
											break;
									}
								}
								#endregion
								break;
							case "billing_address":
								#region [ Campos internos do nó 'billing_address' ]
								foreach (XmlNode nodeN2 in node["value"].ChildNodes)
								{
									strKey = nodeN2["key"].InnerText;
									strValue = (nodeN2["value"].InnerText ?? "");
									switch (strKey)
									{
										case "parent_id":
											orderInfo.billing_address.parent_id = strValue;
											break;
										case "customer_address_id":
											orderInfo.billing_address.customer_address_id = strValue;
											break;
										case "quote_address_id":
											orderInfo.billing_address.quote_address_id = strValue;
											break;
										case "region_id":
											orderInfo.billing_address.region_id = strValue;
											break;
										case "customer_id":
											orderInfo.billing_address.customer_id = strValue;
											break;
										case "fax":
											orderInfo.billing_address.fax = strValue;
											break;
										case "region":
											orderInfo.billing_address.region = strValue;
											break;
										case "postcode":
											orderInfo.billing_address.postcode = strValue;
											break;
										case "firstname":
											orderInfo.billing_address.firstname = strValue;
											break;
										case "middlename":
											orderInfo.billing_address.middlename = strValue;
											break;
										case "lastname":
											orderInfo.billing_address.lastname = strValue;
											break;
										case "street":
											orderInfo.billing_address.street = strValue;
											break;
										case "city":
											orderInfo.billing_address.city = strValue;
											break;
										case "email":
											orderInfo.billing_address.email = strValue;
											break;
										case "telephone":
											orderInfo.billing_address.telephone = strValue;
											break;
										case "country_id":
											orderInfo.billing_address.country_id = strValue;
											break;
										case "address_type":
											orderInfo.billing_address.address_type = strValue;
											break;
										case "prefix":
											orderInfo.billing_address.prefix = strValue;
											break;
										case "suffix":
											orderInfo.billing_address.suffix = strValue;
											break;
										case "company":
											orderInfo.billing_address.company = strValue;
											break;
										case "vat_id":
											orderInfo.billing_address.vat_id = strValue;
											break;
										case "vat_is_valid":
											orderInfo.billing_address.vat_is_valid = strValue;
											break;
										case "vat_request_id":
											orderInfo.billing_address.vat_request_id = strValue;
											break;
										case "vat_request_date":
											orderInfo.billing_address.vat_request_date = strValue;
											break;
										case "vat_request_success":
											orderInfo.billing_address.vat_request_success = strValue;
											break;
										case "tipopessoa":
											orderInfo.billing_address.tipopessoa = strValue;
											break;
										case "rg":
											orderInfo.billing_address.rg = strValue;
											break;
										case "ie":
											orderInfo.billing_address.ie = strValue;
											break;
										case "cpfcnpj":
											orderInfo.billing_address.cpfcnpj = strValue;
											break;
										case "celular":
											orderInfo.billing_address.celular = strValue;
											break;
										case "empresa":
											orderInfo.billing_address.empresa = strValue;
											break;
										case "nomefantasia":
											orderInfo.billing_address.nomefantasia = strValue;
											break;
										case "cpf":
											orderInfo.billing_address.cpf = strValue;
											break;
										case "address_id":
											orderInfo.billing_address.address_id = strValue;
											break;
										default:
											orderInfo.billing_address.UnknownFields.Add(new KeyValuePair<string, string>(strKey, strValue));
											break;
									}
								}
								#endregion
								break;
							case "items":
								#region [ Campos da coleção 'items' ]
								foreach (XmlNode nodeItem in node["value"].ChildNodes)
								{
									orderItem = new MagentoSoapApiSalesOrderItem();
									foreach (XmlNode nodeN2 in nodeItem.ChildNodes)
									{
										strKey = nodeN2["key"].InnerText;
										strValue = (nodeN2["value"].InnerText ?? "");
										switch (strKey)
										{
											case "item_id":
												orderItem.item_id = strValue;
												break;
											case "order_id":
												orderItem.order_id = strValue;
												break;
											case "parent_item_id":
												orderItem.parent_item_id = strValue;
												break;
											case "quote_item_id":
												orderItem.quote_item_id = strValue;
												break;
											case "store_id":
												orderItem.store_id = strValue;
												break;
											case "created_at":
												orderItem.created_at = strValue;
												break;
											case "updated_at":
												orderItem.updated_at = strValue;
												break;
											case "product_id":
												orderItem.product_id = strValue;
												break;
											case "product_type":
												orderItem.product_type = strValue;
												break;
											case "product_options":
												orderItem.product_options = strValue;
												break;
											case "weight":
												orderItem.weight = strValue;
												break;
											case "is_virtual":
												orderItem.is_virtual = strValue;
												break;
											case "sku":
												orderItem.sku = strValue;
												break;
											case "name":
												orderItem.name = strValue;
												break;
											case "description":
												orderItem.description = strValue;
												break;
											case "applied_rule_ids":
												orderItem.applied_rule_ids = strValue;
												break;
											case "additional_data":
												orderItem.additional_data = strValue;
												break;
											case "free_shipping":
												orderItem.free_shipping = strValue;
												break;
											case "is_qty_decimal":
												orderItem.is_qty_decimal = strValue;
												break;
											case "no_discount":
												orderItem.no_discount = strValue;
												break;
											case "qty_backordered":
												orderItem.qty_backordered = strValue;
												break;
											case "qty_canceled":
												orderItem.qty_canceled = strValue;
												break;
											case "qty_invoiced":
												orderItem.qty_invoiced = strValue;
												break;
											case "qty_ordered":
												orderItem.qty_ordered = strValue;
												break;
											case "qty_refunded":
												orderItem.qty_refunded = strValue;
												break;
											case "qty_shipped":
												orderItem.qty_shipped = strValue;
												break;
											case "base_cost":
												orderItem.base_cost = strValue;
												break;
											case "price":
												orderItem.price = strValue;
												break;
											case "base_price":
												orderItem.base_price = strValue;
												break;
											case "original_price":
												orderItem.original_price = strValue;
												break;
											case "base_original_price":
												orderItem.base_original_price = strValue;
												break;
											case "tax_percent":
												orderItem.tax_percent = strValue;
												break;
											case "tax_amount":
												orderItem.tax_amount = strValue;
												break;
											case "base_tax_amount":
												orderItem.base_tax_amount = strValue;
												break;
											case "tax_invoiced":
												orderItem.tax_invoiced = strValue;
												break;
											case "base_tax_invoiced":
												orderItem.base_tax_invoiced = strValue;
												break;
											case "discount_percent":
												orderItem.discount_percent = strValue;
												break;
											case "discount_amount":
												orderItem.discount_amount = strValue;
												break;
											case "base_discount_amount":
												orderItem.base_discount_amount = strValue;
												break;
											case "discount_invoiced":
												orderItem.discount_invoiced = strValue;
												break;
											case "base_discount_invoiced":
												orderItem.base_discount_invoiced = strValue;
												break;
											case "amount_refunded":
												orderItem.amount_refunded = strValue;
												break;
											case "base_amount_refunded":
												orderItem.base_amount_refunded = strValue;
												break;
											case "row_total":
												orderItem.row_total = strValue;
												break;
											case "base_row_total":
												orderItem.base_row_total = strValue;
												break;
											case "row_invoiced":
												orderItem.row_invoiced = strValue;
												break;
											case "base_row_invoiced":
												orderItem.base_row_invoiced = strValue;
												break;
											case "row_weight":
												orderItem.row_weight = strValue;
												break;
											case "base_tax_before_discount":
												orderItem.base_tax_before_discount = strValue;
												break;
											case "tax_before_discount":
												orderItem.tax_before_discount = strValue;
												break;
											case "ext_order_item_id":
												orderItem.ext_order_item_id = strValue;
												break;
											case "locked_do_invoice":
												orderItem.locked_do_invoice = strValue;
												break;
											case "locked_do_ship":
												orderItem.locked_do_ship = strValue;
												break;
											case "price_incl_tax":
												orderItem.price_incl_tax = strValue;
												break;
											case "base_price_incl_tax":
												orderItem.base_price_incl_tax = strValue;
												break;
											case "row_total_incl_tax":
												orderItem.row_total_incl_tax = strValue;
												break;
											case "base_row_total_incl_tax":
												orderItem.base_row_total_incl_tax = strValue;
												break;
											case "hidden_tax_amount":
												orderItem.hidden_tax_amount = strValue;
												break;
											case "base_hidden_tax_amount":
												orderItem.base_hidden_tax_amount = strValue;
												break;
											case "hidden_tax_invoiced":
												orderItem.hidden_tax_invoiced = strValue;
												break;
											case "base_hidden_tax_invoiced":
												orderItem.base_hidden_tax_invoiced = strValue;
												break;
											case "hidden_tax_refunded":
												orderItem.hidden_tax_refunded = strValue;
												break;
											case "base_hidden_tax_refunded":
												orderItem.base_hidden_tax_refunded = strValue;
												break;
											case "is_nominal":
												orderItem.is_nominal = strValue;
												break;
											case "tax_canceled":
												orderItem.tax_canceled = strValue;
												break;
											case "hidden_tax_canceled":
												orderItem.hidden_tax_canceled = strValue;
												break;
											case "tax_refunded":
												orderItem.tax_refunded = strValue;
												break;
											case "base_tax_refunded":
												orderItem.base_tax_refunded = strValue;
												break;
											case "discount_refunded":
												orderItem.discount_refunded = strValue;
												break;
											case "base_discount_refunded":
												orderItem.base_discount_refunded = strValue;
												break;
											case "gift_message_id":
												orderItem.gift_message_id = strValue;
												break;
											case "gift_message_available":
												orderItem.gift_message_available = strValue;
												break;
											case "base_weee_tax_applied_amount":
												orderItem.base_weee_tax_applied_amount = strValue;
												break;
											case "base_weee_tax_applied_row_amnt":
												orderItem.base_weee_tax_applied_row_amnt = strValue;
												break;
											case "base_weee_tax_applied_row_amount":
												orderItem.base_weee_tax_applied_row_amount = strValue;
												break;
											case "weee_tax_applied_amount":
												orderItem.weee_tax_applied_amount = strValue;
												break;
											case "weee_tax_applied_row_amount":
												orderItem.weee_tax_applied_row_amount = strValue;
												break;
											case "weee_tax_applied":
												orderItem.weee_tax_applied = strValue;
												break;
											case "weee_tax_disposition":
												orderItem.weee_tax_disposition = strValue;
												break;
											case "weee_tax_row_disposition":
												orderItem.weee_tax_row_disposition = strValue;
												break;
											case "base_weee_tax_disposition":
												orderItem.base_weee_tax_disposition = strValue;
												break;
											case "base_weee_tax_row_disposition":
												orderItem.base_weee_tax_row_disposition = strValue;
												break;
											case "installer_document":
												orderItem.installer_document = strValue;
												break;
											case "commission_type":
												orderItem.commission_type = strValue;
												break;
											case "commission_value":
												orderItem.commission_value = strValue;
												break;
											case "has_children":
												orderItem.has_children = strValue;
												break;
											default:
												orderItem.UnknownFields.Add(new KeyValuePair<string, string>(strKey, strValue));
												break;
										}
									}
									orderInfo.items.Add(orderItem);
								}
								#endregion
								break;
							case "payment":
								#region [ Campos internos do nó 'payment' ]
								foreach (XmlNode nodeN2 in node["value"].ChildNodes)
								{
									strKey = nodeN2["key"].InnerText;
									strValue = (nodeN2["value"].InnerText ?? "");
									switch (strKey)
									{
										case "parent_id":
											orderInfo.payment.parent_id = strValue;
											break;
										case "base_shipping_captured":
											orderInfo.payment.base_shipping_captured = strValue;
											break;
										case "shipping_captured":
											orderInfo.payment.shipping_captured = strValue;
											break;
										case "amount_refunded":
											orderInfo.payment.amount_refunded = strValue;
											break;
										case "base_amount_paid":
											orderInfo.payment.base_amount_paid = strValue;
											break;
										case "amount_canceled":
											orderInfo.payment.amount_canceled = strValue;
											break;
										case "base_amount_authorized":
											orderInfo.payment.base_amount_authorized = strValue;
											break;
										case "base_amount_paid_online":
											orderInfo.payment.base_amount_paid_online = strValue;
											break;
										case "base_amount_refunded_online":
											orderInfo.payment.base_amount_refunded_online = strValue;
											break;
										case "base_shipping_amount":
											orderInfo.payment.base_shipping_amount = strValue;
											break;
										case "shipping_amount":
											orderInfo.payment.shipping_amount = strValue;
											break;
										case "amount_paid":
											orderInfo.payment.amount_paid = strValue;
											break;
										case "amount_authorized":
											orderInfo.payment.amount_authorized = strValue;
											break;
										case "base_amount_ordered":
											orderInfo.payment.base_amount_ordered = strValue;
											break;
										case "base_shipping_refunded":
											orderInfo.payment.base_shipping_refunded = strValue;
											break;
										case "shipping_refunded":
											orderInfo.payment.shipping_refunded = strValue;
											break;
										case "base_amount_refunded":
											orderInfo.payment.base_amount_refunded = strValue;
											break;
										case "amount_ordered":
											orderInfo.payment.amount_ordered = strValue;
											break;
										case "base_amount_canceled":
											orderInfo.payment.base_amount_canceled = strValue;
											break;
										case "quote_payment_id":
											orderInfo.payment.quote_payment_id = strValue;
											break;
										case "additional_data":
											orderInfo.payment.additional_data = strValue;
											break;
										case "cc_exp_month":
											orderInfo.payment.cc_exp_month = strValue;
											break;
										case "cc_ss_start_year":
											orderInfo.payment.cc_ss_start_year = strValue;
											break;
										case "echeck_bank_name":
											orderInfo.payment.echeck_bank_name = strValue;
											break;
										case "method":
											orderInfo.payment.method = strValue;
											break;
										case "cc_debug_request_body":
											orderInfo.payment.cc_debug_request_body = strValue;
											break;
										case "cc_secure_verify":
											orderInfo.payment.cc_secure_verify = strValue;
											break;
										case "protection_eligibility":
											orderInfo.payment.protection_eligibility = strValue;
											break;
										case "cc_approval":
											orderInfo.payment.cc_approval = strValue;
											break;
										case "cc_last4":
											orderInfo.payment.cc_last4 = strValue;
											break;
										case "cc_status_description":
											orderInfo.payment.cc_status_description = strValue;
											break;
										case "echeck_type":
											orderInfo.payment.echeck_type = strValue;
											break;
										case "cc_debug_response_serialized":
											orderInfo.payment.cc_debug_response_serialized = strValue;
											break;
										case "cc_ss_start_month":
											orderInfo.payment.cc_ss_start_month = strValue;
											break;
										case "echeck_account_type":
											orderInfo.payment.echeck_account_type = strValue;
											break;
										case "last_trans_id":
											orderInfo.payment.last_trans_id = strValue;
											break;
										case "cc_cid_status":
											orderInfo.payment.cc_cid_status = strValue;
											break;
										case "cc_owner":
											orderInfo.payment.cc_owner = strValue;
											break;
										case "cc_type":
											orderInfo.payment.cc_type = strValue;
											break;
										case "po_number":
											orderInfo.payment.po_number = strValue;
											break;
										case "cc_exp_year":
											orderInfo.payment.cc_exp_year = strValue;
											break;
										case "cc_status":
											orderInfo.payment.cc_status = strValue;
											break;
										case "echeck_routing_number":
											orderInfo.payment.echeck_routing_number = strValue;
											break;
										case "account_status":
											orderInfo.payment.account_status = strValue;
											break;
										case "anet_trans_method":
											orderInfo.payment.anet_trans_method = strValue;
											break;
										case "cc_debug_response_body":
											orderInfo.payment.cc_debug_response_body = strValue;
											break;
										case "cc_ss_issue":
											orderInfo.payment.cc_ss_issue = strValue;
											break;
										case "echeck_account_name":
											orderInfo.payment.echeck_account_name = strValue;
											break;
										case "cc_avs_status":
											orderInfo.payment.cc_avs_status = strValue;
											break;
										case "cc_number_enc":
											orderInfo.payment.cc_number_enc = strValue;
											break;
										case "cc_trans_id":
											orderInfo.payment.cc_trans_id = strValue;
											break;
										case "paybox_request_number":
											orderInfo.payment.paybox_request_number = strValue;
											break;
										case "address_status":
											orderInfo.payment.address_status = strValue;
											break;
										case "cc_parcelamento":
											orderInfo.payment.cc_parcelamento = strValue;
											break;
										case "additional_information":
											#region [ Campos internos do nó 'additional_information' ]
											foreach (XmlNode nodeN3 in nodeN2["value"].ChildNodes)
											{
												strKey = nodeN3["key"].InnerText;
												strValue = (nodeN3["value"].InnerText ?? "");
												switch (strKey)
												{
													case "PaymentMethod":
														orderInfo.payment.additional_information.PaymentMethod = strValue;
														break;
													case "InstallmentsCount":
														orderInfo.payment.additional_information.InstallmentsCount = strValue;
														break;
													case "BraspagOrderId":
														orderInfo.payment.additional_information.BraspagOrderId = strValue;
														break;
													case "ErrorDescription":
														orderInfo.payment.additional_information.ErrorDescription = strValue;
														break;
													default:
														orderInfo.payment.additional_information.UnknownFields.Add(new KeyValuePair<string, string>(strKey, strValue));
														break;
												}
											}
											#endregion
											break;
										case "cc_type2":
											orderInfo.payment.cc_type2 = strValue;
											break;
										case "cc_owner2":
											orderInfo.payment.cc_owner2 = strValue;
											break;
										case "cc_last42":
											orderInfo.payment.cc_last42 = strValue;
											break;
										case "cc_number_enc2":
											orderInfo.payment.cc_number_enc2 = strValue;
											break;
										case "cc_exp_month2":
											orderInfo.payment.cc_exp_month2 = strValue;
											break;
										case "cc_exp_year2":
											orderInfo.payment.cc_exp_year2 = strValue;
											break;
										case "cc_ss_issue2":
											orderInfo.payment.cc_ss_issue2 = strValue;
											break;
										case "cc_cid2":
											orderInfo.payment.cc_cid2 = strValue;
											break;
										case "cc_parcelamento2":
											orderInfo.payment.cc_parcelamento2 = strValue;
											break;
										case "additional_information2":
											#region [ Campos internos do nó 'additional_information2' ]
											foreach (XmlNode nodeN3 in nodeN2["value"].ChildNodes)
											{
												strKey = nodeN3["key"].InnerText;
												strValue = (nodeN3["value"].InnerText ?? "");
												switch (strKey)
												{
													case "PaymentMethod":
														orderInfo.payment.additional_information2.PaymentMethod = strValue;
														break;
													case "InstallmentsCount":
														orderInfo.payment.additional_information2.InstallmentsCount = strValue;
														break;
													case "BraspagOrderId":
														orderInfo.payment.additional_information2.BraspagOrderId = strValue;
														break;
													case "ErrorDescription":
														orderInfo.payment.additional_information2.ErrorDescription = strValue;
														break;
													default:
														orderInfo.payment.additional_information2.UnknownFields.Add(new KeyValuePair<string, string>(strKey, strValue));
														break;
												}
											}
											#endregion
											break;
										case "bseller_payment_in_cash":
											orderInfo.payment.bseller_payment_in_cash = strValue;
											break;
										case "bseller_payment_installment":
											orderInfo.payment.bseller_payment_installment = strValue;
											break;
										case "payment_id":
											orderInfo.payment.payment_id = strValue;
											break;
										case "integracommerce_name":
											orderInfo.payment.integracommerce_name = strValue;
											break;
										case "integracommerce_installments":
											orderInfo.payment.integracommerce_installments = strValue;
											break;
										default:
											orderInfo.payment.UnknownFields.Add(new KeyValuePair<string, string>(strKey, strValue));
											break;
									}
								}
								#endregion
								break;
							case "status_history":
								#region [ Campos da coleção 'status_history' ]
								foreach (XmlNode nodeSH in node["value"].ChildNodes)
								{
									statusHistory = new MagentoSoapApiStatusHistory();
									foreach (XmlNode nodeN2 in nodeSH.ChildNodes)
									{
										strKey = nodeN2["key"].InnerText;
										strValue = (nodeN2["value"].InnerText ?? "");
										switch (strKey)
										{
											case "parent_id":
												statusHistory.parent_id = strValue;
												break;
											case "is_customer_notified":
												statusHistory.is_customer_notified = strValue;
												break;
											case "is_visible_on_front":
												statusHistory.is_visible_on_front = strValue;
												break;
											case "comment":
												statusHistory.comment = strValue;
												break;
											case "status":
												statusHistory.status = strValue;
												break;
											case "created_at":
												statusHistory.created_at = strValue;
												break;
											case "entity_name":
												statusHistory.entity_name = strValue;
												break;
											case "store_id":
												statusHistory.store_id = strValue;
												break;
											default:
												statusHistory.UnknownFields.Add(new KeyValuePair<string, string>(strKey, strValue));
												break;
										}
									}
									orderInfo.status_history.Add(statusHistory);
								}
								#endregion
								break;
							default:
								#region [ Decodifica campos da resposta que não são coleções ]
								strValue = (node["value"].InnerText ?? "");
								switch (strKey)
								{
									case "increment_id":
										orderInfo.increment_id = strValue;
										break;
									case "parent_id":
										orderInfo.parent_id = strValue;
										break;
									case "store_id":
										orderInfo.store_id = strValue;
										break;
									case "created_at":
										orderInfo.created_at = strValue;
										break;
									case "updated_at":
										orderInfo.updated_at = strValue;
										break;
									case "is_active":
										orderInfo.is_active = strValue;
										break;
									case "customer_id":
										orderInfo.customer_id = strValue;
										break;
									case "shipping_amount":
										orderInfo.shipping_amount = strValue;
										break;
									case "shipping_canceled":
										orderInfo.shipping_canceled = strValue;
										break;
									case "shipping_invoiced":
										orderInfo.shipping_invoiced = strValue;
										break;
									case "shipping_refunded":
										orderInfo.shipping_refunded = strValue;
										break;
									case "shipping_tax_amount":
										orderInfo.shipping_tax_amount = strValue;
										break;
									case "shipping_tax_refunded":
										orderInfo.shipping_tax_refunded = strValue;
										break;
									case "shipping_discount_amount":
										orderInfo.shipping_discount_amount = strValue;
										break;
									case "discount_amount":
										orderInfo.discount_amount = strValue;
										break;
									case "discount_canceled":
										orderInfo.discount_canceled = strValue;
										break;
									case "discount_invoiced":
										orderInfo.discount_invoiced = strValue;
										break;
									case "discount_refunded":
										orderInfo.discount_refunded = strValue;
										break;
									case "subtotal":
										orderInfo.subtotal = strValue;
										break;
									case "subtotal_canceled":
										orderInfo.subtotal_canceled = strValue;
										break;
									case "subtotal_invoiced":
										orderInfo.subtotal_invoiced = strValue;
										break;
									case "subtotal_refunded":
										orderInfo.subtotal_refunded = strValue;
										break;
									case "subtotal_incl_tax":
										orderInfo.subtotal_incl_tax = strValue;
										break;
									case "tax_amount":
										orderInfo.tax_amount = strValue;
										break;
									case "tax_canceled":
										orderInfo.tax_canceled = strValue;
										break;
									case "tax_invoiced":
										orderInfo.tax_invoiced = strValue;
										break;
									case "tax_refunded":
										orderInfo.tax_refunded = strValue;
										break;
									case "grand_total":
										orderInfo.grand_total = strValue;
										break;
									case "total_paid":
										orderInfo.total_paid = strValue;
										break;
									case "total_refunded":
										orderInfo.total_refunded = strValue;
										break;
									case "total_qty_ordered":
										orderInfo.total_qty_ordered = strValue;
										break;
									case "total_canceled":
										orderInfo.total_canceled = strValue;
										break;
									case "total_invoiced":
										orderInfo.total_invoiced = strValue;
										break;
									case "total_due":
										orderInfo.total_due = strValue;
										break;
									case "total_online_refunded":
										orderInfo.total_online_refunded = strValue;
										break;
									case "total_offline_refunded":
										orderInfo.total_offline_refunded = strValue;
										break;
									case "base_tax_amount":
										orderInfo.base_tax_amount = strValue;
										break;
									case "base_tax_canceled":
										orderInfo.base_tax_canceled = strValue;
										break;
									case "base_tax_invoiced":
										orderInfo.base_tax_invoiced = strValue;
										break;
									case "base_tax_refunded":
										orderInfo.base_tax_refunded = strValue;
										break;
									case "base_shipping_amount":
										orderInfo.base_shipping_amount = strValue;
										break;
									case "base_shipping_canceled":
										orderInfo.base_shipping_canceled = strValue;
										break;
									case "base_shipping_invoiced":
										orderInfo.base_shipping_invoiced = strValue;
										break;
									case "base_shipping_refunded":
										orderInfo.base_shipping_refunded = strValue;
										break;
									case "base_shipping_tax_amount":
										orderInfo.base_shipping_tax_amount = strValue;
										break;
									case "base_shipping_tax_refunded":
										orderInfo.base_shipping_tax_refunded = strValue;
										break;
									case "base_discount_amount":
										orderInfo.base_discount_amount = strValue;
										break;
									case "base_discount_canceled":
										orderInfo.base_discount_canceled = strValue;
										break;
									case "base_discount_invoiced":
										orderInfo.base_discount_invoiced = strValue;
										break;
									case "base_discount_refunded":
										orderInfo.base_discount_refunded = strValue;
										break;
									case "base_subtotal":
										orderInfo.base_subtotal = strValue;
										break;
									case "base_subtotal_canceled":
										orderInfo.base_subtotal_canceled = strValue;
										break;
									case "base_subtotal_invoiced":
										orderInfo.base_subtotal_invoiced = strValue;
										break;
									case "base_subtotal_refunded":
										orderInfo.base_subtotal_refunded = strValue;
										break;
									case "base_grand_total":
										orderInfo.base_grand_total = strValue;
										break;
									case "base_total_paid":
										orderInfo.base_total_paid = strValue;
										break;
									case "base_total_refunded":
										orderInfo.base_total_refunded = strValue;
										break;
									case "base_total_qty_ordered":
										orderInfo.base_total_qty_ordered = strValue;
										break;
									case "base_total_canceled":
										orderInfo.base_total_canceled = strValue;
										break;
									case "base_total_invoiced":
										orderInfo.base_total_invoiced = strValue;
										break;
									case "base_total_invoiced_cost":
										orderInfo.base_total_invoiced_cost = strValue;
										break;
									case "base_total_online_refunded":
										orderInfo.base_total_online_refunded = strValue;
										break;
									case "base_total_offline_refunded":
										orderInfo.base_total_offline_refunded = strValue;
										break;
									case "billing_address_id":
										orderInfo.billing_address_id = strValue;
										break;
									case "billing_firstname":
										orderInfo.billing_firstname = strValue;
										break;
									case "billing_lastname":
										orderInfo.billing_lastname = strValue;
										break;
									case "shipping_address_id":
										orderInfo.shipping_address_id = strValue;
										break;
									case "shipping_firstname":
										orderInfo.shipping_firstname = strValue;
										break;
									case "shipping_lastname":
										orderInfo.shipping_lastname = strValue;
										break;
									case "billing_name":
										orderInfo.billing_name = strValue;
										break;
									case "shipping_name":
										orderInfo.shipping_name = strValue;
										break;
									case "store_to_base_rate":
										orderInfo.store_to_base_rate = strValue;
										break;
									case "store_to_order_rate":
										orderInfo.store_to_order_rate = strValue;
										break;
									case "base_to_global_rate":
										orderInfo.base_to_global_rate = strValue;
										break;
									case "base_to_order_rate":
										orderInfo.base_to_order_rate = strValue;
										break;
									case "weight":
										orderInfo.weight = strValue;
										break;
									case "store_name":
										orderInfo.store_name = strValue;
										break;
									case "remote_ip":
										orderInfo.remote_ip = strValue;
										break;
									case "status":
										orderInfo.status = strValue;
										break;
									case "state":
										orderInfo.state = strValue;
										break;
									case "applied_rule_ids":
										orderInfo.applied_rule_ids = strValue;
										break;
									case "global_currency_code":
										orderInfo.global_currency_code = strValue;
										break;
									case "base_currency_code":
										orderInfo.base_currency_code = strValue;
										break;
									case "store_currency_code":
										orderInfo.store_currency_code = strValue;
										break;
									case "order_currency_code":
										orderInfo.order_currency_code = strValue;
										break;
									case "shipping_method":
										orderInfo.shipping_method = strValue;
										break;
									case "shipping_description":
										orderInfo.shipping_description = strValue;
										break;
									case "customer_email":
										orderInfo.customer_email = strValue;
										break;
									case "customer_firstname":
										orderInfo.customer_firstname = strValue;
										break;
									case "customer_lastname":
										orderInfo.customer_lastname = strValue;
										break;
									case "customer_middlename":
										orderInfo.customer_middlename = strValue;
										break;
									case "customer_prefix":
										orderInfo.customer_prefix = strValue;
										break;
									case "customer_suffix":
										orderInfo.customer_suffix = strValue;
										break;
									case "customer_taxvat":
										orderInfo.customer_taxvat = strValue;
										break;
									case "quote_id":
										orderInfo.quote_id = strValue;
										break;
									case "is_virtual":
										orderInfo.is_virtual = strValue;
										break;
									case "customer_group_id":
										orderInfo.customer_group_id = strValue;
										break;
									case "customer_note":
										orderInfo.customer_note = strValue;
										break;
									case "customer_note_notify":
										orderInfo.customer_note_notify = strValue;
										break;
									case "customer_is_guest":
										orderInfo.customer_is_guest = strValue;
										break;
									case "email_sent":
										orderInfo.email_sent = strValue;
										break;
									case "order_id":
										orderInfo.order_id = strValue;
										break;
									case "gift_message_id":
										orderInfo.gift_message_id = strValue;
										break;
									case "gift_message":
										orderInfo.gift_message = strValue;
										break;
									case "coupon_code":
										orderInfo.coupon_code = strValue;
										break;
									case "protect_code":
										orderInfo.protect_code = strValue;
										break;
									case "can_ship_partially":
										orderInfo.can_ship_partially = strValue;
										break;
									case "can_ship_partially_item":
										orderInfo.can_ship_partially_item = strValue;
										break;
									case "edit_increment":
										orderInfo.edit_increment = strValue;
										break;
									case "forced_shipment_with_invoice":
										orderInfo.forced_shipment_with_invoice = strValue;
										break;
									case "forced_do_shipment_with_invoice":
										orderInfo.forced_do_shipment_with_invoice = strValue;
										break;
									case "payment_auth_expiration":
										orderInfo.payment_auth_expiration = strValue;
										break;
									case "quote_address_id":
										orderInfo.quote_address_id = strValue;
										break;
									case "adjustment_negative":
										orderInfo.adjustment_negative = strValue;
										break;
									case "adjustment_positive":
										orderInfo.adjustment_positive = strValue;
										break;
									case "base_adjustment_negative":
										orderInfo.base_adjustment_negative = strValue;
										break;
									case "base_adjustment_positive":
										orderInfo.base_adjustment_positive = strValue;
										break;
									case "base_shipping_discount_amount":
										orderInfo.base_shipping_discount_amount = strValue;
										break;
									case "base_subtotal_incl_tax":
										orderInfo.base_subtotal_incl_tax = strValue;
										break;
									case "base_total_due":
										orderInfo.base_total_due = strValue;
										break;
									case "payment_authorization_amount":
										orderInfo.payment_authorization_amount = strValue;
										break;
									case "customer_dob":
										orderInfo.customer_dob = strValue;
										break;
									case "discount_description":
										orderInfo.discount_description = strValue;
										break;
									case "ext_customer_id":
										orderInfo.ext_customer_id = strValue;
										break;
									case "ext_order_id":
										orderInfo.ext_order_id = strValue;
										break;
									case "hold_before_state":
										orderInfo.hold_before_state = strValue;
										break;
									case "hold_before_status":
										orderInfo.hold_before_status = strValue;
										break;
									case "original_increment_id":
										orderInfo.original_increment_id = strValue;
										break;
									case "relation_child_id":
										orderInfo.relation_child_id = strValue;
										break;
									case "relation_child_real_id":
										orderInfo.relation_child_real_id = strValue;
										break;
									case "relation_parent_id":
										orderInfo.relation_parent_id = strValue;
										break;
									case "relation_parent_real_id":
										orderInfo.relation_parent_real_id = strValue;
										break;
									case "x_forwarded_for":
										orderInfo.x_forwarded_for = strValue;
										break;
									case "total_item_count":
										orderInfo.total_item_count = strValue;
										break;
									case "customer_gender":
										orderInfo.customer_gender = strValue;
										break;
									case "hidden_tax_amount":
										orderInfo.hidden_tax_amount = strValue;
										break;
									case "base_hidden_tax_amount":
										orderInfo.base_hidden_tax_amount = strValue;
										break;
									case "shipping_hidden_tax_amount":
										orderInfo.shipping_hidden_tax_amount = strValue;
										break;
									case "base_shipping_hidden_tax_amnt":
										orderInfo.base_shipping_hidden_tax_amnt = strValue;
										break;
									case "hidden_tax_invoiced":
										orderInfo.hidden_tax_invoiced = strValue;
										break;
									case "base_hidden_tax_invoiced":
										orderInfo.base_hidden_tax_invoiced = strValue;
										break;
									case "hidden_tax_refunded":
										orderInfo.hidden_tax_refunded = strValue;
										break;
									case "base_hidden_tax_refunded":
										orderInfo.base_hidden_tax_refunded = strValue;
										break;
									case "shipping_incl_tax":
										orderInfo.shipping_incl_tax = strValue;
										break;
									case "base_shipping_incl_tax":
										orderInfo.base_shipping_incl_tax = strValue;
										break;
									case "coupon_rule_name":
										orderInfo.coupon_rule_name = strValue;
										break;
									case "paypal_ipn_customer_notified":
										orderInfo.paypal_ipn_customer_notified = strValue;
										break;
									case "firecheckout_delivery_date":
										orderInfo.firecheckout_delivery_date = strValue;
										break;
									case "firecheckout_delivery_timerange":
										orderInfo.firecheckout_delivery_timerange = strValue;
										break;
									case "firecheckout_customer_comment":
										orderInfo.firecheckout_customer_comment = strValue;
										break;
									case "tm_field1":
										orderInfo.tm_field1 = strValue;
										break;
									case "tm_field2":
										orderInfo.tm_field2 = strValue;
										break;
									case "tm_field3":
										orderInfo.tm_field3 = strValue;
										break;
									case "tm_field4":
										orderInfo.tm_field4 = strValue;
										break;
									case "tm_field5":
										orderInfo.tm_field5 = strValue;
										break;
									case "from_lengow":
										orderInfo.from_lengow = strValue;
										break;
									case "order_id_lengow":
										orderInfo.order_id_lengow = strValue;
										break;
									case "fees_lengow":
										orderInfo.fees_lengow = strValue;
										break;
									case "xml_node_lengow":
										orderInfo.xml_node_lengow = strValue;
										break;
									case "feed_id_lengow":
										orderInfo.feed_id_lengow = strValue;
										break;
									case "message_lengow":
										orderInfo.message_lengow = strValue;
										break;
									case "marketplace_lengow":
										orderInfo.marketplace_lengow = strValue;
										break;
									case "total_paid_lengow":
										orderInfo.total_paid_lengow = strValue;
										break;
									case "carrier_lengow":
										orderInfo.carrier_lengow = strValue;
										break;
									case "carrier_method_lengow":
										orderInfo.carrier_method_lengow = strValue;
										break;
									case "clearsale_status_code":
										orderInfo.clearsale_status_code = strValue;
										break;
									case "session_id":
										orderInfo.session_id = strValue;
										break;
									case "skyhub_code":
										orderInfo.skyhub_code = strValue;
										break;
									case "commission_value":
										orderInfo.commission_value = strValue;
										break;
									case "installer_document":
										orderInfo.installer_document = strValue;
										break;
									case "installer_id":
										orderInfo.installer_id = strValue;
										break;
									case "commission_discount":
										orderInfo.commission_discount = strValue;
										break;
									case "commission_final_discount":
										orderInfo.commission_final_discount = strValue;
										break;
									case "commission_discount_type":
										orderInfo.commission_discount_type = strValue;
										break;
									case "commission_final_value":
										orderInfo.commission_final_value = strValue;
										break;
									case "base_bseller_payment_total_tax_rate":
										orderInfo.base_bseller_payment_total_tax_rate = strValue;
										break;
									case "bseller_payment_total_tax_rate":
										orderInfo.bseller_payment_total_tax_rate = strValue;
										break;
									case "payment_authorization_expiration":
										orderInfo.payment_authorization_expiration = strValue;
										break;
									case "base_shipping_hidden_tax_amount":
										orderInfo.base_shipping_hidden_tax_amount = strValue;
										break;
									case "clearSale_status":
										orderInfo.clearSale_status = strValue;
										break;
									case "clearSale_score":
										orderInfo.clearSale_score = strValue;
										break;
									case "clearSale_packageID":
										orderInfo.clearSale_packageID = strValue;
										break;
									case "clearSale_fingerPrintSessionId":
										orderInfo.clearSale_fingerPrintSessionId = strValue;
										break;
									case "integracommerce_id":
										orderInfo.integracommerce_id = strValue;
										break;
									case "bseller_skyhub":
										orderInfo.bseller_skyhub = strValue;
										break;
									case "bseller_skyhub_code":
										orderInfo.bseller_skyhub_code = strValue;
										break;
									case "bseller_skyhub_channel":
										orderInfo.bseller_skyhub_channel = strValue;
										break;
									case "bseller_skyhub_invoice_key":
										orderInfo.bseller_skyhub_invoice_key = strValue;
										break;
									case "bseller_skyhub_interest":
										orderInfo.bseller_skyhub_interest = strValue;
										break;
									case "bseller_skyhub_json":
										orderInfo.bseller_skyhub_json = strValue;
										break;
									default:
										orderInfo.UnknownFields.Add(new KeyValuePair<string, string>(strKey, strValue));
										break;
								}
								#endregion
								break;
						}
					}
				}
				#endregion

				#region [ Decodifica resposta de erro? ]
				if (xmlRespSoap.Contains(":Fault>") && xmlRespSoap.Contains("<faultcode>"))
				{
					orderInfo.faultResponse.isFaultResponse = true;

					xmlDoc = new XmlDocument();
					xmlDoc.LoadXml(xmlRespSoap);
					nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
					nsmgr.AddNamespace("SOAP-ENV", "http://schemas.xmlsoap.org/soap/envelope/");
					xmlNodeListN1 = xmlDoc.SelectNodes("//SOAP-ENV:Fault", nsmgr);
					foreach (XmlNode nodeN1 in xmlNodeListN1)
					{
						if (nodeN1.HasChildNodes)
						{
							foreach (XmlNode nodeN2 in nodeN1)
							{
								strKey = (nodeN2.Name ?? "");
								strValue = (nodeN2.InnerText ?? "");
								switch (strKey)
								{
									case "faultcode":
										orderInfo.faultResponse.faultcode = strValue;
										break;
									case "faultstring":
										orderInfo.faultResponse.faultstring = strValue;
										break;
									default:
										break;
								}
							}
						}
					}
				}
				#endregion

				return orderInfo;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				return null;
			}
		}
		#endregion

		#region [ getSalesOrderInfo ]
		public static string getSalesOrderInfo(string numeroPedidoMagento, MagentoApiLoginParameters loginParameters, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "getSalesOrderInfo()";
			bool blnEnviouOk;
			string msg;
			string msg_erro_aux;
			string xmlReqSoap;
			string xmlRespSoap;
			string sessionId;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Consistências ]
				if ((numeroPedidoMagento ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o número do pedido Magento!";
					return null;
				}

				if ((loginParameters.urlWebService ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o endereço do web service da API do Magento!";
					return null;
				}

				if ((loginParameters.username ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informado o nome do usuário para login no web service da API do Magento!";
					return null;
				}

				if ((loginParameters.password ?? "").Trim().Length == 0)
				{
					msg_erro = "Não foi informada a senha para login no web service da API do Magento!";
					return null;
				}
				#endregion

				try
				{
					rwlMagentoSoapApi.AcquireWriterLock(Global.Cte.MagentoSoapApi.TIMEOUT_READER_WRITER_LOCK_EM_MS);
					try // finally: rwlMagentoSoapApi.ReleaseWriterLock();
					{
						xmlReqSoap = montaRequisicaoLogin(loginParameters.username, loginParameters.password);

						blnEnviouOk = enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.MagentoSoapApi.Transacao.login, loginParameters.urlWebService, out xmlRespSoap, out msg_erro_aux);
						if (!blnEnviouOk)
						{
							msg_erro = msg_erro_aux;
							return null;
						}

						sessionId = obtemSessionIdFromLoginResponse(xmlRespSoap, out msg_erro_aux);

						if ((sessionId ?? "").Length == 0)
						{
							msg_erro = "Falha ao tentar obter o SessionId!!";
							return null;
						}

						try // Finally: Encerra sessão
						{
							xmlReqSoap = montaRequisicaoCallSalesOrderInfo(sessionId, numeroPedidoMagento);
							blnEnviouOk = MagentoSoapApi.enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.MagentoSoapApi.Transacao.call, loginParameters.urlWebService, out xmlRespSoap, out msg_erro_aux);
							if (!blnEnviouOk)
							{
								msg_erro = "Falha ao tentar consultar o pedido Magento " + numeroPedidoMagento + " através da API!";
								if (msg_erro_aux.Length > 0) msg_erro += "\n" + msg_erro_aux;
								return null;
							}

							return xmlRespSoap;
						}
						finally
						{
							xmlReqSoap = montaRequisicaoEndSession(sessionId);
							blnEnviouOk = MagentoSoapApi.enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.MagentoSoapApi.Transacao.endSession, loginParameters.urlWebService, out xmlRespSoap, out msg_erro_aux);
						}
					}
					finally
					{
						rwlMagentoSoapApi.ReleaseWriterLock();
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
	}
}