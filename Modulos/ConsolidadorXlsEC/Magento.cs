using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace ConsolidadorXlsEC
{
	class Magento
	{
		#region [ enviaRequisicaoComRetry ]
		/// <summary>
		/// Método que executa o enviaRequisicao() dentro de um laço de tentativas até que a execução seja bem sucedida ou a quantidade máxima de tentativas seja atingida.
		/// Importante: este método pode ser utilizado livremente para requisições de consulta, entretanto, para requisições que alteram dados é importante avaliar antes
		/// as possíveis consequências que podem ocorrer no caso da requisição ter sido processada no web service e o erro ter ocorrido em algum estágio posterior durante
		/// o recebimento da resposta. Nesse caso, o uso deste método pode causar múltiplas execuções da requisição.
		/// </summary>
		/// <param name="xmlReqSoap"></param>
		/// <param name="trxParam"></param>
		/// <param name="xmlRespSoap"></param>
		/// <param name="msg_erro"></param>
		/// <returns></returns>
		public static bool enviaRequisicaoComRetry(string xmlReqSoap, Global.Cte.Magento.Transacao trxParam, out string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			const int MAX_TENTATIVAS = 5;
			int qtdeTentativasRealizadas = 0;
			bool blnResposta;
			#endregion

			do
			{
				qtdeTentativasRealizadas++;

				blnResposta = enviaRequisicao(xmlReqSoap, trxParam, out xmlRespSoap, out msg_erro);
				if (blnResposta) break;

				Thread.Sleep(1000);
			} while (qtdeTentativasRealizadas < MAX_TENTATIVAS);

			return blnResposta;
		}
		#endregion

		#region [ enviaRequisicao ]
		public static bool enviaRequisicao(string xmlReqSoap, Global.Cte.Magento.Transacao trxParam, out string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "enviaRequisicao()";
			HttpWebRequest req;
			HttpWebResponse resp;
			#endregion

			xmlRespSoap = "";
			msg_erro = "";

			try
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + trxParam.GetMethodName() + " - XML (TX)\n" + xmlReqSoap);

				req = (HttpWebRequest)WebRequest.Create(trxParam.GetEnderecoWebService());
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

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + trxParam.GetMethodName() + " - XML (RX)\n" + xmlRespSoap);

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + trxParam.GetMethodName() + " - Exception\n" + ex.ToString());
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

		#region [ montaRequisicaoCallCatalogProductList ]
		public static string montaRequisicaoCallCatalogProductList(string sessionId)
		{
			string xmlRequisicaoSoap;

			xmlRequisicaoSoap = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
								"<soapenv:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:urn=\"urn: Magento\">" +
								"<soapenv:Header/>" +
								"<soapenv:Body>" +
								"<urn:call soapenv:encodingStyle=\"http://schemas.xmlsoap.org/soap/encoding/\">" +
								"<sessionId xsi:type=\"xsd: string\">" + sessionId + "</sessionId>" +
								"<resourcePath xsi:type=\"xsd: string\">catalog_product.list</resourcePath>" +
								"</urn:call>" +
								"</soapenv:Body>" +
								"</soapenv:Envelope>";

			return xmlRequisicaoSoap;
		}
		#endregion

		#region [ montaRequisicaoCallCatalogProductInfo ]
		public static string montaRequisicaoCallCatalogProductInfo(string sessionId, string product_id)
		{
			string xmlRequisicaoSoap;

			xmlRequisicaoSoap = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
								"<soapenv:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:urn=\"urn: Magento\">" +
								"<soapenv:Header/>" +
								"<soapenv:Body>" +
								"<urn:call soapenv:encodingStyle=\"http://schemas.xmlsoap.org/soap/encoding/\">" +
								"<sessionId xsi:type=\"xsd: string\">" + sessionId + "</sessionId>" +
								"<resourcePath xsi:type=\"xsd: string\">catalog_product.info</resourcePath>" +
								"<args xsi:type=\"xsd: anyType\">" + product_id + "</args>" +
								"</urn:call>" +
								"</soapenv:Body>" +
								"</soapenv:Envelope>";

			return xmlRequisicaoSoap;
		}
        #endregion

        #region [ montaRequisicaoSalesOrderAddComment ]
        public static string montaRequisicaoSalesOrderAddComment(string sessionId, SalesOrderAddCommentRequest addCommentRequest)
        {
            StringBuilder xmlRequisicaoSoap = new StringBuilder("");

            xmlRequisicaoSoap.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                                "<SOAP-ENV:Envelope xmlns:SOAP-ENV=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:ns1=\"urn:Magento\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:ns2=\"http://xml.apache.org/xml-soap\" xmlns:SOAP-ENC=\"http://schemas.xmlsoap.org/soap/encoding/\" SOAP-ENV:encodingStyle=\"http://schemas.xmlsoap.org/soap/encoding/\">" +
                                "  <SOAP-ENV:Body>" +
                                    "<ns1:call>" +
                                     "<sessionId xsi:type=\"xsd:string\">" + sessionId + "</sessionId>" +
                                     "<resourcePath xsi:type=\"xsd:string\">sales_order.addComment</resourcePath>" +
                                     "<args xsi:type=\"ns2:Map\">" +
                                        "<item>" +
                                            "<key xsi:type=\"xsd:string\">orderIncrementId</key>" +
                                            "<value xsi:type=\"xsd:string\">" + addCommentRequest.orderIncrementId + "</value>" +
                                        "</item>" +
                                        "<item>" +
                                            "<key xsi:type=\"xsd:string\">status</key>" +
                                            "<value xsi:type=\"xsd:string\">" + addCommentRequest.status + "</value>" +
                                        "</item>");

            if ((addCommentRequest.comment ?? "").Length > 0)
            {
                xmlRequisicaoSoap.Append("<item>" +
                                            "<key xsi:type=\"xsd:string\">comment</key>" +
                                            "<value xsi:type=\"xsd:string\">" + addCommentRequest.comment + "</value>" +
                                        "</item>");
            }

            if ((addCommentRequest.notify ?? "").Length > 0)
            {
                xmlRequisicaoSoap.Append("<item>" +
                                            "<key xsi:type=\"xsd:string\">notify</key>" +
                                            "<value xsi:type=\"xsd:string\">" + addCommentRequest.notify + "</value>" +
                                        "</item>");
            }

            xmlRequisicaoSoap.Append("</args>" +
                                    "</ns1:call>" +
                                  "</SOAP-ENV:Body>" +
                                "</SOAP-ENV:Envelope>");

            return xmlRequisicaoSoap.ToString();
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
			string strValue;
			string sessionId = "";
			XmlDocument xmlDoc;
			XmlNamespaceManager nsmgr;
			XmlNode xmlNode;

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

		#region [ decodificaXmlCatalogProductListResponse ]
		public static List<ProductList> decodificaXmlCatalogProductListResponse(string xmlRespSoap, out string msg_erro)
		{
			string strKey;
			string strValue;
			XmlDocument xmlDoc;
			XmlNamespaceManager nsmgr;
			XmlNodeList xmlNodeListN1;
			XmlNodeList xmlNodeListN2;
			ProductList productList;
			List<ProductList> vProductList = new List<ProductList>();

			msg_erro = "";

			try
			{
				if ((xmlRespSoap ?? "").Trim().Length == 0) return null;

				xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(xmlRespSoap);
				nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
				nsmgr.AddNamespace("ns1", "urn:Magento");
				xmlNodeListN1 = xmlDoc.SelectNodes("//ns1:callResponse/callReturn/item", nsmgr);
				foreach (XmlNode nodesN1 in xmlNodeListN1)
				{
					productList = new ProductList();
					xmlNodeListN2 = nodesN1.ChildNodes;
					foreach (XmlNode node in xmlNodeListN2)
					{
						strKey = (node["key"].InnerText ?? "");
						switch (strKey)
						{
							case "category_ids":
								foreach (XmlNode nodeN2 in node["value"].ChildNodes)
								{
									strValue = (nodeN2.InnerText ?? "");
									productList.category_ids.Add(strValue);
								}
								break;
							case "website_ids":
								foreach (XmlNode nodeN2 in node["value"].ChildNodes)
								{
									strValue = (nodeN2.InnerText ?? "");
									productList.website_ids.Add(strValue);
								}
								break;
							default:
								strValue = (node["value"].InnerText ?? "");
								switch (strKey)
								{
									case "product_id":
										productList.product_id = strValue;
										break;
									case "sku":
										productList.sku = strValue;
										break;
									case "name":
										productList.name = strValue;
										break;
									case "set":
										productList.set = strValue;
										break;
									case "type":
										productList.type = strValue;
										break;
									default:
										productList.UnknownFields.Add(strKey);
										break;
								}
								break;
						}
					}

					vProductList.Add(productList);
				}

				return vProductList;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				return null;
			}
		}
		#endregion

		#region [ decodificaXmlCatalogProductInfoResponse ]
		public static ProductInfo decodificaXmlCatalogProductInfoResponse(string xmlRespSoap, out string msg_erro)
		{
			string strKey;
			string strValue;
			XmlDocument xmlDoc;
			XmlNamespaceManager nsmgr;
			XmlNodeList xmlNodeListN1;
			ProductInfo productInfo = new ProductInfo();

			msg_erro = "";

			try
			{
				if ((xmlRespSoap ?? "").Trim().Length == 0) return null;

				xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(xmlRespSoap);
				nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
				nsmgr.AddNamespace("ns1", "urn:Magento");
				xmlNodeListN1 = xmlDoc.SelectNodes("//ns1:callResponse/callReturn/item", nsmgr);
				foreach (XmlNode node in xmlNodeListN1)
				{
					strKey = (node["key"].InnerText ?? "");
					switch (strKey)
					{
						case "categories":
							foreach (XmlNode nodeN2 in node["value"].ChildNodes)
							{
								strValue = (nodeN2.InnerText ?? "");
								productInfo.categories.Add(strValue);
							}
							break;
						case "websites":
							foreach (XmlNode nodeN2 in node["value"].ChildNodes)
							{
								strValue = (nodeN2.InnerText ?? "");
								productInfo.websites.Add(strValue);
							}
							break;
						case "category_ids":
							foreach (XmlNode nodeN2 in node["value"].ChildNodes)
							{
								strValue = (nodeN2.InnerText ?? "");
								productInfo.category_ids.Add(strValue);
							}
							break;
						default:
							strValue = (node["value"].InnerText ?? "");
							switch (strKey)
							{
								case "product_id":
									productInfo.product_id = strValue;
									break;
								case "sku":
									productInfo.sku = strValue;
									break;
								case "set":
									productInfo.set = strValue;
									break;
								case "type":
									productInfo.type = strValue;
									break;
								case "type_id":
									productInfo.type_id = strValue;
									break;
								case "name":
									productInfo.name = strValue;
									break;
								case "titulo_ml":
									productInfo.titulo_ml = strValue;
									break;
								case "weight":
									productInfo.weight = strValue;
									break;
								case "cubagem":
									productInfo.cubagem = strValue;
									break;
								case "old_id":
									productInfo.old_id = strValue;
									break;
								case "news_from_date":
									productInfo.news_from_date = strValue;
									break;
								case "news_to_date":
									productInfo.news_to_date = strValue;
									break;
								case "status":
									productInfo.status = strValue;
									break;
								case "url_key":
									productInfo.url_key = strValue;
									break;
								case "visibility":
									productInfo.visibility = strValue;
									break;
								case "url_path":
									productInfo.url_path = strValue;
									break;
								case "country_of_manufacture":
									productInfo.country_of_manufacture = strValue;
									break;
								case "volume_comprimento":
									productInfo.volume_comprimento = strValue;
									break;
								case "volume_altura":
									productInfo.volume_altura = strValue;
									break;
								case "required_options":
									productInfo.required_options = strValue;
									break;
								case "volume_largura":
									productInfo.volume_largura = strValue;
									break;
								case "vender_buscape":
									productInfo.vender_buscape = strValue;
									break;
								case "has_options":
									productInfo.has_options = strValue;
									break;
								case "image_label":
									productInfo.image_label = strValue;
									break;
								case "enviado_buscape":
									productInfo.enviado_buscape = strValue;
									break;
								case "small_image_label":
									productInfo.small_image_label = strValue;
									break;
								case "fretegratis":
									productInfo.fretegratis = strValue;
									break;
								case "thumbnail_label":
									productInfo.thumbnail_label = strValue;
									break;
								case "created_at":
									productInfo.created_at = strValue;
									break;
								case "updated_at":
									productInfo.updated_at = strValue;
									break;
								case "price":
									productInfo.price = strValue;
									break;
								case "parcelamento_cartao":
									productInfo.parcelamento_cartao = strValue;
									break;
								case "parcela_cartao":
									productInfo.parcela_cartao = strValue;
									break;
								case "forma_pagamento":
									productInfo.forma_pagamento = strValue;
									break;
								case "group_price":
									productInfo.group_price = strValue;
									break;
								case "special_price":
									productInfo.special_price = strValue;
									break;
								case "special_from_date":
									productInfo.special_from_date = strValue;
									break;
								case "special_to_date":
									productInfo.special_to_date = strValue;
									break;
								case "minimal_price":
									productInfo.minimal_price = strValue;
									break;
								case "tier_price":
									productInfo.tier_price = strValue;
									break;
								case "msrp_enabled":
									productInfo.msrp_enabled = strValue;
									break;
								case "msrp_display_actual_price_type":
									productInfo.msrp_display_actual_price_type = strValue;
									break;
								case "msrp":
									productInfo.msrp = strValue;
									break;
								case "tax_class_id":
									productInfo.tax_class_id = strValue;
									break;
								case "markup":
									productInfo.markup = strValue;
									break;
								case "preco_calculado":
									productInfo.preco_calculado = strValue;
									break;
								case "procel":
									productInfo.procel = strValue;
									break;
								case "marca":
									productInfo.marca = strValue;
									break;
								case "codigo_fabricante":
									productInfo.codigo_fabricante = strValue;
									break;
								case "ean":
									productInfo.ean = strValue;
									break;
								case "inverter":
									productInfo.inverter = strValue;
									break;
								case "voltagem":
									productInfo.voltagem = strValue;
									break;
								case "temperatura":
									productInfo.temperatura = strValue;
									break;
								case "inmetro":
									productInfo.inmetro = strValue;
									break;
								case "consumo_energia":
									productInfo.consumo_energia = strValue;
									break;
								case "trifasico":
									productInfo.trifasico = strValue;
									break;
								case "medida_unidade_interma":
									productInfo.medida_unidade_interma = strValue;
									break;
								case "medida_unidade_externa":
									productInfo.medida_unidade_externa = strValue;
									break;
								case "short_description":
									productInfo.short_description = strValue;
									break;
								case "description":
									productInfo.description = strValue;
									break;
								case "detalhes":
									productInfo.detalhes = strValue;
									break;
								case "additional_1":
									productInfo.additional_1 = strValue;
									break;
								case "additional_2":
									productInfo.additional_2 = strValue;
									break;
								case "detalhes_tecnicos_comparacao":
									productInfo.detalhes_tecnicos_comparacao = strValue;
									break;
								case "multi_split":
									productInfo.multi_split = strValue;
									break;
								case "tipo":
									productInfo.tipo = strValue;
									break;
								case "capacidade":
									productInfo.capacidade = strValue;
									break;
								case "meta_title":
									productInfo.meta_title = strValue;
									break;
								case "meta_keyword":
									productInfo.meta_keyword = strValue;
									break;
								case "meta_description":
									productInfo.meta_description = strValue;
									break;
								case "cjm_imageswitcher":
									productInfo.cjm_imageswitcher = strValue;
									break;
								case "cjm_moreviews":
									productInfo.cjm_moreviews = strValue;
									break;
								case "cjm_useimages":
									productInfo.cjm_useimages = strValue;
									break;
								case "is_recurring":
									productInfo.is_recurring = strValue;
									break;
								case "recurring_profile":
									productInfo.recurring_profile = strValue;
									break;
								case "custom_design":
									productInfo.custom_design = strValue;
									break;
								case "custom_design_from":
									productInfo.custom_design_from = strValue;
									break;
								case "custom_design_to":
									productInfo.custom_design_to = strValue;
									break;
								case "custom_layout_update":
									productInfo.custom_layout_update = strValue;
									break;
								case "page_layout":
									productInfo.page_layout = strValue;
									break;
								case "options_container":
									productInfo.options_container = strValue;
									break;
								case "gift_message_available":
									productInfo.gift_message_available = strValue;
									break;
								case "package_height":
									productInfo.package_height = strValue;
									break;
								case "package_width":
									productInfo.package_width = strValue;
									break;
								case "package_length":
									productInfo.package_length = strValue;
									break;
								case "integra_anymarket":
									productInfo.integra_anymarket = strValue;
									break;
								case "id_anymarket":
									productInfo.id_anymarket = strValue;
									break;
								case "garantia":
									productInfo.garantia = strValue;
									break;
								case "tempo_garantia":
									productInfo.tempo_garantia = strValue;
									break;
								case "categoria_anymarket":
									productInfo.categoria_anymarket = strValue;
									break;
								case "origem":
									productInfo.origem = strValue;
									break;
								case "modelo":
									productInfo.modelo = strValue;
									break;
								case "nbm":
									productInfo.nbm = strValue;
									break;
								case "intelipost_altura":
									productInfo.intelipost_altura = strValue;
									break;
								case "intelipost_largura":
									productInfo.intelipost_largura = strValue;
									break;
								case "intelipost_comprimento":
									productInfo.intelipost_comprimento = strValue;
									break;
								case "intelipost_peso":
									productInfo.intelipost_peso = strValue;
									break;
								case "intelipost_prazo_produto":
									productInfo.intelipost_prazo_produto = strValue;
									break;
								default:
									productInfo.UnknownFields.Add(strKey);
									break;
							}
							break;
					}
				}

				return productInfo;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				return null;
			}
		}
        #endregion

        #region [ decodificaXmlSalesOrderAddCommentResponse ]
        public static SalesOrderAddCommentResponse decodificaXmlSalesOrderAddCommentResponse(string xmlRespSoap, out string msg_erro)
        {
            #region [ Declarações ]
            string strKey;
            string strValue;
            XmlDocument xmlDoc;
            XmlNamespaceManager nsmgr;
            XmlNode xmlNode;
            XmlNodeList xmlNodeListN1;
            SalesOrderAddCommentResponse addCommentResponse = new SalesOrderAddCommentResponse();
            #endregion

            msg_erro = "";
            try
            {
                if ((xmlRespSoap ?? "").Trim().Length == 0) return null;

                #region [ Decodifica resposta ]
                xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xmlRespSoap);
                nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                nsmgr.AddNamespace("ns1", "urn:Magento");
                xmlNode = xmlDoc.SelectSingleNode("//ns1:callResponse", nsmgr);
                if (xmlNode != null)
                {
                    strValue = Global.obtemXmlChildNodeValue(xmlNode, "callReturn");
                    if (strValue != null) addCommentResponse.callReturn = strValue;
                }
                #endregion

                #region [ Decodifica resposta de erro? ]
                if (xmlRespSoap.Contains(":Fault>") && xmlRespSoap.Contains("<faultcode>"))
                {
                    addCommentResponse.faultResponse.isFaultResponse = true;

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
                                        addCommentResponse.faultResponse.faultcode = strValue;
                                        break;
                                    case "faultstring":
                                        addCommentResponse.faultResponse.faultstring = strValue;
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                    }
                }
                #endregion

                return addCommentResponse;
            }
            catch (Exception ex)
            {
                msg_erro = ex.ToString();
                return null;
            }
        }
        #endregion

        #region [ decodificaXmlSalesOrderInfoResponse ]
        public static SalesOrderInfo decodificaXmlSalesOrderInfoResponse(string xmlRespSoap, out string msg_erro)
        {
            #region [ Declarações ]
            string strKey;
            string strValue;
            XmlDocument xmlDoc;
            XmlNamespaceManager nsmgr;
            XmlNodeList xmlNodeListN1;
            SalesOrderInfo orderInfo = new SalesOrderInfo();
            SalesOrderItem orderItem;
            StatusHistory statusHistory;
            #endregion

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
                                    orderItem = new SalesOrderItem();
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
                                    statusHistory = new StatusHistory();
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
    }
}
