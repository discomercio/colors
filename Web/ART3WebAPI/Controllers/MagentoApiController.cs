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
using System.Xml;
using Newtonsoft.Json;
using System.Text;
using System.Web.Script.Serialization;

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
			bool blnInserted = false;
			string msg;
			string msg_erro;
			string sXml = null;
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
			Usuario usuarioBD;
			Cliente cliente;
			MagentoErpPedidoXml readPedidoXml = null;
			MagentoErpPedidoXml insertPedidoXml = null;
			MagentoApiLoginParameters loginParameters;
			MagentoErpSalesOrder salesOrder = new MagentoErpSalesOrder();
			List<string> listaPedidosERP;
			Pedido pedidoERP;
			MagentoErpPedidoXmlDecodeEndereco decodeEndereco;
			MagentoErpPedidoXmlDecodeItem decodeItem;
			MagentoErpPedidoXmlDecodeStatusHistory decodeStatusHistory;
			MagentoSoapApiStatusHistory statusHistory;
			HttpResponseMessage result;
			List<CodigoDescricao> listaCodigoDescricao;
			string[] v;
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

			#region [ Verifica se o pedido já foi consultado e a resposta se encontra gravada no BD, desde que se trate da mesma operação ]
			if ((operationControlTicket ?? "").Trim().Length > 0)
			{
				if ((numeroPedidoMagento ?? "").Trim().Length == 0)
				{
					throw new Exception("O número do pedido Magento não foi informado!");
				}

				readPedidoXml = MagentoApiDAO.getMagentoPedidoXmlByTicket(numeroPedidoMagento, operationControlTicket, out msg_erro);
				if (readPedidoXml != null)
				{
					msg = "Pedido Magento nº " + numeroPedidoMagento + " localizado no BD";
					Global.gravaLogAtividade(msg);

					salesOrder.cpfCnpjIdentificado = readPedidoXml.cpfCnpjIdentificado;

					if ((readPedidoXml.pedido_xml ?? "").Trim().Length > 0)
					{
						sXml = readPedidoXml.pedido_xml;

						#region [ Converte XML da resposta do Magento em objeto ]
						salesOrder.magentoSalesOrderInfo = MagentoSoapApi.decodificaXmlSalesOrderInfoResponse(sXml, out msg_erro);
						if (salesOrder.magentoSalesOrderInfo == null)
						{
							msg = "Falha ao tentar decodificar o XML de resposta da API do Magento!";
							if (msg_erro.Length > 0) msg += "\n" + msg_erro;
							throw new Exception(msg);
						}
						#endregion
					}
				}
			}
			#endregion

			#region [ Não encontrou os dados do pedido armazenados no BD, executa consulta via API ]
			if ((readPedidoXml == null) || ((sXml ?? "").Trim().Length == 0))
			{
				if ((loja ?? "").Trim().Length == 0)
				{
					throw new Exception("O número da loja não foi informado!");
				}

				loginParameters = MagentoApiDAO.getLoginParameters(loja, out msg_erro);
				if (loginParameters == null)
				{
					msg = "Falha ao tentar recuperar os parâmetros de login da API do Magento para a loja " + loja + "!";
					if (msg_erro.Length > 0) msg += "\n" + msg_erro;
					throw new Exception(msg);
				}

				#region [ Há parâmetros de login cadastrados para a loja? ]
				if ((loginParameters.urlWebService ?? "").Trim().Length == 0)
				{
					msg = "Falha ao tentar recuperar os parâmetros de login da API do Magento: a URL da API não está cadastrada para a loja " + loja + "!";
					if (msg_erro.Length > 0) msg += "\n" + msg_erro;
					throw new Exception(msg);
				}

				if ((loginParameters.username ?? "").Trim().Length == 0)
				{
					msg = "Falha ao tentar recuperar os parâmetros de login da API do Magento: o usuário para login não está cadastrado para a loja " + loja + "!";
					if (msg_erro.Length > 0) msg += "\n" + msg_erro;
					throw new Exception(msg);
				}

				if ((loginParameters.password ?? "").Trim().Length == 0)
				{
					msg = "Falha ao tentar recuperar os parâmetros de login da API do Magento: a senha para login não está cadastrada para a loja " + loja + "!";
					if (msg_erro.Length > 0) msg += "\n" + msg_erro;
					throw new Exception(msg);
				}
				#endregion

				#region [ Executa a consulta via API ]
				msg = "Consulta do pedido Magento nº " + numeroPedidoMagento + " via API";
				Global.gravaLogAtividade(msg);
				sXml = MagentoSoapApi.getSalesOrderInfo(numeroPedidoMagento, loginParameters, out msg_erro);
				#endregion

				#region [ Falha ao obter os dados do pedido Magento ]
				if ((sXml ?? "").Trim().Length == 0)
				{
					msg = "Falha desconhecida ao tentar recuperar os dados do pedido Magento " + numeroPedidoMagento + "!";
					if (msg_erro.Length > 0) msg += "\n" + msg_erro;
					throw new Exception(msg);
				}
				#endregion

				#region [ Converte XML da resposta do Magento em objeto ]
				salesOrder.magentoSalesOrderInfo = MagentoSoapApi.decodificaXmlSalesOrderInfoResponse(sXml, out msg_erro);
				if (salesOrder.magentoSalesOrderInfo == null)
				{
					msg = "Falha ao tentar decodificar o XML de resposta da API do Magento!";
					if (msg_erro.Length > 0) msg += "\n" + msg_erro;
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

				if ((salesOrder.magentoSalesOrderInfo.bseller_skyhub ?? "").Equals("1") && ((salesOrder.magentoSalesOrderInfo.bseller_skyhub_code ?? "").Trim().Length > 0))
				{
					#region [ Tenta identificar o nº pedido marketplace através do campo 'bseller_skyhub_code' (ao invés do comentário registrado no status history) ]
					foreach (var codDescr in listaCodigoDescricao)
					{
						// Verifica se o flag está configurado para que seja feito o tratamento usando o campo 'bseller_skyhub_code'
						if (codDescr.parametro_1_campo_flag == 0) continue;
						sParametro = (codDescr.parametro_2_campo_texto ?? "").Trim();
						if (sParametro.Length == 0) continue;
						vMktpOrderDescriptor = sParametro.Split('|');
						for (int k = 0; k < vMktpOrderDescriptor.Length; k++)
						{
							sMktpOrderDescriptor = vMktpOrderDescriptor[k];
							if ((sMktpOrderDescriptor ?? "").Trim().Length == 0) continue;
							if (salesOrder.magentoSalesOrderInfo.bseller_skyhub_code.Trim().ToUpper().StartsWith(sMktpOrderDescriptor.ToUpper()))
							{
								// Obtém a parte relativa ao nº pedido marketplace
								sValue = salesOrder.magentoSalesOrderInfo.bseller_skyhub_code.Trim().Substring(sMktpOrderDescriptor.Length).Trim();
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

				#region [ Se não conseguiu identificar o nº pedido marketplace através do campo 'bseller_skyhub_code', analisa o status history ]
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
				insertPedidoXml.pedido_xml = sXml;
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
				if (v.Length >= 1) decodeEndereco.endereco = v[0];
				if (v.Length >= 2) decodeEndereco.endereco_numero = v[1];
				if (v.Length >= 3) decodeEndereco.endereco_complemento = v[2];
				if (v.Length >= 4) decodeEndereco.bairro = v[3];
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
				decodeEndereco.telephone = (salesOrder.magentoSalesOrderInfo.billing_address.telephone ?? "");
				decodeEndereco.celular = (salesOrder.magentoSalesOrderInfo.billing_address.celular ?? "");
				decodeEndereco.fax = (salesOrder.magentoSalesOrderInfo.billing_address.fax ?? "");
				decodeEndereco.tipopessoa = (salesOrder.magentoSalesOrderInfo.billing_address.tipopessoa ?? "");
				decodeEndereco.rg = (salesOrder.magentoSalesOrderInfo.billing_address.rg ?? "");
				decodeEndereco.ie = (salesOrder.magentoSalesOrderInfo.billing_address.ie ?? "");
				decodeEndereco.cpfcnpj = (salesOrder.magentoSalesOrderInfo.billing_address.cpfcnpj ?? "");
				decodeEndereco.empresa = (salesOrder.magentoSalesOrderInfo.billing_address.empresa ?? "");
				decodeEndereco.nomefantasia = (salesOrder.magentoSalesOrderInfo.billing_address.nomefantasia ?? "");
				if (!MagentoApiDAO.insertMagentoPedidoXmlDecodeEndereco(decodeEndereco, out msg_erro))
				{
					msg = "Falha ao tentar gravar no BD os dados do endereço de cobrança!";
					if (msg_erro.Length > 0) msg += "\n" + msg_erro;
					throw new Exception(msg);
				}
				#endregion

				#region [ Endereço de entrega ]
				decodeEndereco = new MagentoErpPedidoXmlDecodeEndereco();
				decodeEndereco.id_magento_api_pedido_xml = insertPedidoXml.id;
				decodeEndereco.tipo_endereco = Global.Cte.MagentoSoapApi.TIPO_ENDERECO__ENTREGA;
				v = (salesOrder.magentoSalesOrderInfo.shipping_address.street ?? "").Split('\n');
				if (v.Length >= 1) decodeEndereco.endereco = v[0];
				if (v.Length >= 2) decodeEndereco.endereco_numero = v[1];
				if (v.Length >= 3) decodeEndereco.endereco_complemento = v[2];
				if (v.Length >= 4) decodeEndereco.bairro = v[3];
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
				decodeEndereco.telephone = (salesOrder.magentoSalesOrderInfo.shipping_address.telephone ?? "");
				decodeEndereco.celular = (salesOrder.magentoSalesOrderInfo.shipping_address.celular ?? "");
				decodeEndereco.fax = (salesOrder.magentoSalesOrderInfo.shipping_address.fax ?? "");
				decodeEndereco.tipopessoa = (salesOrder.magentoSalesOrderInfo.shipping_address.tipopessoa ?? "");
				decodeEndereco.rg = (salesOrder.magentoSalesOrderInfo.shipping_address.rg ?? "");
				decodeEndereco.ie = (salesOrder.magentoSalesOrderInfo.shipping_address.ie ?? "");
				decodeEndereco.cpfcnpj = (salesOrder.magentoSalesOrderInfo.shipping_address.cpfcnpj ?? "");
				decodeEndereco.empresa = (salesOrder.magentoSalesOrderInfo.shipping_address.empresa ?? "");
				decodeEndereco.nomefantasia = (salesOrder.magentoSalesOrderInfo.shipping_address.nomefantasia ?? "");
				if (!MagentoApiDAO.insertMagentoPedidoXmlDecodeEndereco(decodeEndereco, out msg_erro))
				{
					msg = "Falha ao tentar gravar no BD os dados do endereço de entrega!";
					if (msg_erro.Length > 0) msg += "\n" + msg_erro;
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
					if (!MagentoApiDAO.insertMagentoPedidoXmlDecodeItem(decodeItem, out msg_erro))
					{
						msg = "Falha ao tentar gravar no BD os dados do item do pedido (sku=" + decodeItem.sku + ")!";
						if (msg_erro.Length > 0) msg += "\n" + msg_erro;
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

			result = Request.CreateResponse(HttpStatusCode.OK);
			result.Content = new StringContent(serializedResult, Encoding.UTF8, "text/html");
			return result;
		}
		#endregion
	}
}
