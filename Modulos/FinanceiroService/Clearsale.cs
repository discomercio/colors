#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Net;
using System.IO;
using System.Xml;
using System.Web;
using System.Threading;
#endregion

namespace FinanceiroService
{
	static class Clearsale
	{
		#region [ Atributos ]
		private static int qtdeFalhasConsecutivasMetodoGetReturnAnalysis = 0;
		#endregion

		#region [ bandeiraCodificaParaPadraoClearsale ]
		public static string bandeiraCodificaParaPadraoClearsale(string bandeira)
		{
			#region [ Declarações ]
			string strResp = "";
			#endregion

			bandeira = bandeira.Trim().ToUpper();
			if (bandeira.Equals(Global.Cte.Braspag.Bandeira.DINERS.GetValue().ToUpper()))
			{
				strResp = "1";
			}
			else if (bandeira.Equals(Global.Cte.Braspag.Bandeira.MASTERCARD.GetValue().ToUpper()))
			{
				strResp = "2";
			}
			else if (bandeira.Equals(Global.Cte.Braspag.Bandeira.VISA.GetValue().ToUpper()))
			{
				strResp = "3";
			}
			else if (bandeira.Equals(Global.Cte.Braspag.Bandeira.AMEX.GetValue().ToUpper()))
			{
				strResp = "5";
			}
			else if (bandeira.Equals(Global.Cte.Braspag.Bandeira.AURA.GetValue().ToUpper()))
			{
				strResp = "7";
			}
			else
			{
				// Outros
				strResp = "4";
			}

			return strResp;
		}
		#endregion

		#region [ isAFStatusAprovado ]
		private static bool isAFStatusAprovado(string status)
		{
			if (status == null) return false;
			if (status.Trim().Length == 0) return false;
			if (status.Equals(Global.Cte.Clearsale.StatusAF.APROVACAO_AUTOMATICA.GetValue())) return true;
			if (status.Equals(Global.Cte.Clearsale.StatusAF.APROVACAO_MANUAL.GetValue())) return true;
			if (status.Equals(Global.Cte.Clearsale.StatusAF.APROVACAO_POR_POLITICA.GetValue())) return true;
			return false;
		}
		#endregion

		#region [ isAFStatusReprovado ]
		private static bool isAFStatusReprovado(string status)
		{
			if (status == null) return false;
			if (status.Trim().Length == 0) return false;
			if (status.Equals(Global.Cte.Clearsale.StatusAF.REPROVADO_SEM_SUSPEITA.GetValue())) return true;
			if (status.Equals(Global.Cte.Clearsale.StatusAF.FRAUDE_CONFIRMADA.GetValue())) return true;
			if (status.Equals(Global.Cte.Clearsale.StatusAF.SUSPENSAO_MANUAL.GetValue())) return true;
			if (status.Equals(Global.Cte.Clearsale.StatusAF.REPROVACAO_POR_POLITICA.GetValue())) return true;
			if (status.Equals(Global.Cte.Clearsale.StatusAF.REPROVACAO_AUTOMATICA.GetValue())) return true;
			if (status.Equals(Global.Cte.Clearsale.StatusAF.CANCELADO_PELO_CLIENTE.GetValue())) return true;
			return false;
		}
		#endregion

		#region [ executaProcessamentoAntifraude ]
		public static bool executaProcessamentoAntifraude(out int qtdePedidosNovosEnviados, out int qtdePedidosFalhaEnvio, out int qtdePedidosResultadoProcessado, out string strMsgInformativa, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Clearsale.executaProcessamentoAntifraude()";
			string strMsg;
			string strMsgInformativaAux;
			string msg_erro_aux;
			bool blnEnviaNovasTransacoes;
			bool blnProcessaResultadoAntifraude;
			#endregion

			qtdePedidosNovosEnviados = 0;
			qtdePedidosFalhaEnvio = 0;
			qtdePedidosResultadoProcessado = 0;
			strMsgInformativa = "";
			msg_erro = "";

			try
			{
				strMsg = "Rotina " + NOME_DESTA_ROTINA + " iniciada";
				Global.gravaLogAtividade(strMsg);

				blnEnviaNovasTransacoes = enviaNovasTransacoes(out qtdePedidosNovosEnviados, out qtdePedidosFalhaEnvio, out strMsgInformativaAux, out msg_erro_aux);
				if (!blnEnviaNovasTransacoes)
				{
					msg_erro = msg_erro_aux;
				}
				if ((strMsgInformativaAux ?? "").Trim().Length == 0) strMsgInformativaAux = "Pedidos novos enviados: " + qtdePedidosNovosEnviados.ToString() + "; Pedidos com falha no envio: " + qtdePedidosFalhaEnvio.ToString();
				strMsgInformativa = strMsgInformativaAux;

				blnProcessaResultadoAntifraude = processaResultadoAntifraude(out qtdePedidosResultadoProcessado, out strMsgInformativaAux, out msg_erro_aux);
				if (!blnProcessaResultadoAntifraude)
				{
					if (msg_erro.Length > 0) msg_erro += "\r\n\r\n";
					msg_erro += msg_erro_aux;
				}
				if ((strMsgInformativaAux ?? "").Trim().Length == 0) strMsgInformativaAux = "Resultado AF (pedidos processados): " + qtdePedidosResultadoProcessado.ToString();
				if (strMsgInformativa.Length > 0) strMsgInformativa += "; ";
				strMsgInformativa += strMsgInformativaAux;

				return (blnEnviaNovasTransacoes && blnProcessaResultadoAntifraude);
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ enviaNovasTransacoes ]
		private static bool enviaNovasTransacoes(out int qtdePedidosNovosEnviados, out int qtdePedidosFalhaEnvio, out string strMsgInformativa, out string msg_erro)
		{
			#region Declarações
			const String NOME_DESTA_ROTINA = "Clearsale.enviaNovasTransacoes()";
			const decimal VL_MARGEM_ERRO = 1m;
			bool blnAchou;
			bool blnEnviarAF;
			bool blnExcedeuLimiteTentativas;
			bool blnSucesso;
			bool blnGerouNsu;
			bool blnEnviarEmail;
			bool blnEnviouOk;
			bool blnExecutar;
			bool blnXmlInvalido;
			bool blnStatusRespSucesso;
			short shortValue;
			byte byteVazioStatus;
			int intValue;
			int intQtyInstallments;
			int intOwner;
			int idPagtoGwAf = 0;
			int idPagtoGwPagPayment;
			int intQtdeAnulados;
			int intQtdeTentativasFalhaTX;
			int id_emailsndsvc_mensagem;
			int nsuSufixo;
			DateTime dtHrTrxPag;
			DateTime dtHrUltTentativaTXFalha;
			String strMsg;
			String strSql;
			String strValue;
			String strOrder_ID;
			String strOrder_FingerPrint_SessionID;
			String strOrder_IP;
			String strOrder_Origin;
			String strPackageStatus_StatusCode;
			String strUsuario;
			String strLoja;
			String numeroPedido;
			String strTrxErroMensagem;
			String strSubject;
			String strBody;
			String xmlReqSoap;
			String xmlRespSoap;
			String msg_erro_aux;
			StringBuilder sbPedidosEnviadosOk = new StringBuilder("");
			StringBuilder sbPedidosFalhaEnvio = new StringBuilder("");
			StringBuilder sbMsgInformativa = new StringBuilder("");
			decimal vl_pagador;
			decimal vl_pedido;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbTrxPag = new DataTable();
			DataTable dtbPagPayment;
			List<string> listaPedidoTrxPag = new List<string>();
			Pedido pedido;
			Cliente cliente;
			StringBuilder sbXml;
			ClearsaleAF clearsaleAF;
			ClearsaleAFPhone afPhone;
			ClearsaleAFItem afItem;
			ClearsaleAFPayment afPayment;
			ClearsaleAFXml clearsaleAFXml;
			EmailCtrl emailCtrl;
			XmlDocument xmlDoc;
			XmlDocument xmlDocSendOrdersResult;
			XmlNode xmlNode;
			XmlNodeList xmlNodeList;
			XmlNamespaceManager nsmgr;
			ClearsaleSendOrdersResponse sendOrdersResponse;
			ClearsaleSendOrdersResponseOrder responseOrder;
			#endregion

			qtdePedidosNovosEnviados = 0;
			qtdePedidosFalhaEnvio = 0;
			strMsgInformativa = "";
			msg_erro = "";

			try
			{
				strMsg = "Rotina " + NOME_DESTA_ROTINA + " iniciada";
				Global.gravaLogAtividade(strMsg);

				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				#region [ Obtém relação de pedidos com transação a enviar ]
				// Importante: o cliente pode realizar o pagamento usando mais do que um cartão.
				// =========== Entretanto, caso um dos cartões não seja autorizado, o cliente precisará fazer novamente o
				// processo de checkout para enviar nova transação ao pagador.
				// Esta rotina tenta prever essa situação confrontando o valor a ser pago com o valor já autorizado pelo pagador.
				// Caso o valor autorizado seja inferior, esta rotina considera que o cliente poderá ainda tentar realizar nova
				// transação c/ o pagador p/ completar o valor. Portanto, esta rotina irá aguardar um determinado intervalo de
				// tempo na expectativa de que possa enviar para o antifraude, de uma única vez, todas as transações de pagamento
				// referentes a um mesmo pedido.
				strSql = "SELECT" +
							" pedido" +
							", data" +
							", data_hora" +
						" FROM t_PAGTO_GW_PAG t_PAG" +
							" INNER JOIN t_PAGTO_GW_PAG_PAYMENT t_PAYMENT ON (t_PAG.id = t_PAYMENT.id_pagto_gw_pag)" +
						" WHERE" +
							" (st_enviado_analise_AF = 0)" +
							" AND (st_cancelado_envio_analise_AF = 0)" +
							" AND (ult_GlobalStatus IN (" +
										"'" + Global.Cte.Braspag.Pagador.GlobalStatus.AUTORIZADA.GetValue() + "'," +
										"'" + Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue() + "'))" +
						" ORDER BY" +
							" data_hora, t_PAG.id";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbTrxPag);
				for (int i = 0; i < dtbTrxPag.Rows.Count; i++)
				{
					blnAchou = false;
					for (int j = 0; j < listaPedidoTrxPag.Count; j++)
					{
						if (BD.readToString(dtbTrxPag.Rows[i]["pedido"]).Equals(listaPedidoTrxPag[j]))
						{
							blnAchou = true;
							break;
						}
					}
					if (!blnAchou)
					{
						listaPedidoTrxPag.Add(BD.readToString(dtbTrxPag.Rows[i]["pedido"]));
					}
				}
				#endregion

				#region [ Mensagem informativa ]
				if (listaPedidoTrxPag.Count == 0)
				{
					if (sbMsgInformativa.Length > 0) sbMsgInformativa.Append("; ");
					sbMsgInformativa.Append("Não há pedidos para enviar");
				}
				#endregion

				#region [ Laço para processamento dos pedidos ]
				// VERIFICA SE O CLIENTE JÁ EFETUOU TRANSAÇÕES DE PAGAMENTO SUFICIENTES P/ ATINGIR
				// O VALOR ESPERADO CONFORME CONSTA NA FORMA DE PAGAMENTO DO PEDIDO.
				// CASO SIM, ENVIA OS DADOS P/ A ANÁLISE ANTIFRAUDE.
				// CASO NÃO, AGUARDA O CLIENTE REALIZAR MAIS TRANSAÇÕES DE PAGAMENTO, ATÉ O LIMITE DE TEMPO ESTABELECIDO.
				for (int iLTP = 0; iLTP < listaPedidoTrxPag.Count; iLTP++)
				{
					#region [ Serviço deve parar? ]
					if (FinanceiroService.isOnShutdownAcionado)
					{
						if (sbMsgInformativa.Length > 0) sbMsgInformativa.Append("; ");
						sbMsgInformativa.Append("Rotina interrompida devido a shutdown do serviço");
						return true;
					}
					if (FinanceiroService.isOnStopAcionado)
					{
						if (sbMsgInformativa.Length > 0) sbMsgInformativa.Append("; ");
						sbMsgInformativa.Append("Rotina interrompida devido a parada do serviço");
						return true;
					}
					#endregion

					blnEnviarAF = false;
					blnExcedeuLimiteTentativas = false;
					numeroPedido = listaPedidoTrxPag[iLTP];
					try
					{
						pedido = PedidoDAO.getPedidoConsolidadoFamilia(numeroPedido);
						cliente = ClienteDAO.getCliente(pedido.id_cliente);
						clearsaleAF = new ClearsaleAF();
					}
					catch (Exception ex)
					{
						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						strMsg = "Falha no envio do pedido " + numeroPedido + " para a Clearsale!\n" +
								 ex.ToString();
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = strMsg;
						svcLog.complemento_1 = numeroPedido;
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						#region [ Envia email de alerta ]
						strSubject = Global.montaIdInstanciaServicoEmailSubject() +
									" Clearsale: Falha ao processar o envio do pedido " + numeroPedido +
									" [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Falha ao processar o envio do pedido " + numeroPedido + " para a Clearsale!" +
								  ex.ToString();
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}
						#endregion

						// Pula para o próximo pedido
						continue;
					}

					#region [ Verifica se houve tentativa anterior com falha e se já passou o intervalo mínimo entre tentativas ]
					dtHrUltTentativaTXFalha = ClearsaleDAO.obtemDataHoraUltTentativaFalhaTX(numeroPedido);
					if (dtHrUltTentativaTXFalha != null)
					{
						if (dtHrUltTentativaTXFalha != DateTime.MinValue)
						{
							if (Global.calculaTimeSpanSegundos(DateTime.Now - dtHrUltTentativaTXFalha) < Global.Parametros.Clearsale.TempoMinEntreTentativasEmSeg)
							{
								// Ainda não passou tempo suficiente para atender ao requisito do intervalo mínimo
								// Não deve enviar para a Clearsale, pula p/ o próximo pedido
								continue;
							}
						}
					}
					#endregion

					#region [ Verifica se já excedeu a quantidade de tentativas ]
					intQtdeTentativasFalhaTX = ClearsaleDAO.contagemTentativasFalhaTX(numeroPedido);
					if (intQtdeTentativasFalhaTX >= Global.Parametros.Clearsale.MaxTentativasEnvioTransacao)
					{
						blnExcedeuLimiteTentativas = true;

						#region [ Envia email de alerta, já que pode se tratar de algum problema inesperado ]
						// Já enviou este mesmo alerta hoje?
						blnEnviarEmail = true;
						emailCtrl = EmailCtrlDAO.getLastEmailCtrlByPedidoFilteredByCodigoMsg(numeroPedido, Global.Cte.EmailCtrl.CodigoMsg.CLEARSALE_EXCEDEU_LIMITE_TENTATIVAS_TX, out msg_erro_aux);
						if (emailCtrl != null)
						{
							if (emailCtrl.data.Equals(DateTime.Now.Date)) blnEnviarEmail = false;
						}

						if (blnEnviarEmail)
						{
							#region [ Envia o email ]
							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Clearsale: pedido " + numeroPedido + " excedeu o limite de " + Global.Parametros.Clearsale.MaxTentativasEnvioTransacao.ToString() + " tentativas de envio [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nPedido " + numeroPedido + " excedeu o limite de " + Global.Parametros.Clearsale.MaxTentativasEnvioTransacao.ToString() + " tentativas de envio de transação para a Clearsale";
							if (EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								emailCtrl = new EmailCtrl();
								emailCtrl.id_emailsndsvc_mensagem = id_emailsndsvc_mensagem;
								emailCtrl.pedido = numeroPedido;
								emailCtrl.id_cliente = pedido.id_cliente;
								emailCtrl.cnpj_cpf_cliente = cliente.cnpj_cpf;
								emailCtrl.tipo_destinatario = Global.Cte.EmailCtrl.TipoDestinatario.ADMINISTRADOR_SISTEMA.GetValue();
								emailCtrl.modulo = Global.Cte.EmailCtrl.Modulo.CLEARSALE.GetValue();
								emailCtrl.tipo_msg = Global.Cte.EmailCtrl.TipoMsg.FALHA.GetValue();
								emailCtrl.codigo_msg = Global.Cte.EmailCtrl.CodigoMsg.CLEARSALE_EXCEDEU_LIMITE_TENTATIVAS_TX.GetValue();
								emailCtrl.rotina = NOME_DESTA_ROTINA;
								emailCtrl.remetente = Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA;
								emailCtrl.destinatario = Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA;
								if (!EmailCtrlDAO.insere(emailCtrl, out msg_erro_aux))
								{
									msg_erro = msg_erro_aux;
									return false;
								}
							}
							#endregion
						}
						#endregion
					}

					// Não deve enviar para a Clearsale, pula p/ o próximo pedido
					if (blnExcedeuLimiteTentativas) continue;
					#endregion

					#region [ Verifica se cliente já realizou transações que totalizam o valor esperado ]
					if (!calculaValorPagador(numeroPedido, out vl_pagador, out msg_erro_aux))
					{
						msg_erro = msg_erro_aux;
						return false;
					}

					if (vl_pagador >= (pedido.vlPagtoEmCartao - VL_MARGEM_ERRO))
					{
						// Transações aprovadas com o pagador totalizam o valor esperado
						blnEnviarAF = true;
					}
					else
					{
						// Verifica se deve aguardar o cliente realizar mais transações ou se já excedeu o timeout
						if (obtemDataHoraTrxPagadorMaisAntiga(numeroPedido, out dtHrTrxPag, out msg_erro_aux))
						{
							if (dtHrTrxPag != DateTime.MinValue)
							{
								if (Global.calculaTimeSpanSegundos(DateTime.Now - dtHrTrxPag) > Global.Parametros.Clearsale.TempoMaxClienteTotalizarPagtoEmSeg) blnEnviarAF = true;
							}
						}
					}

					// Não deve enviar para a Clearsale, pula p/ o próximo pedido
					if (!blnEnviarAF) continue;
					#endregion

					#region [ Envia transação para Clearsale ]

					#region [ Monta dados do XML ]

					#region [ Obtém dados do(s) pagamento(s) ]
					dtbPagPayment = ClearsaleDAO.getPagPaymentRowsParaEnvioAF(numeroPedido, out msg_erro_aux);
					#endregion

					sbXml = new StringBuilder();

					sbXml.Append("<ClearSale>");
					sbXml.Append("<Orders>");
					sbXml.Append("<Order>");

					#region [ Seção: Order ]

					#region [ Gera o campo Order/ID ]
					// É FUNDAMENTAL QUE CADA REQUISIÇÃO FEITA À CLEARSALE POSSUA UM NÚMERO ÚNICO DE ORDER/ID
					if (ClearsaleDAO.geraSufixoPedidoNsuAf(numeroPedido, out nsuSufixo, out msg_erro_aux))
					{
						if (nsuSufixo > 1)
						{
							strOrder_ID = numeroPedido + "-" + nsuSufixo.ToString();
						}
						else
						{
							strOrder_ID = numeroPedido;
						}
					}
					else
					{
						if (msg_erro_aux.Length > 0) msg_erro_aux = "\n" + msg_erro_aux;
						msg_erro_aux = "Falha ao tentar gerar o NSU do sufixo do pedido " + numeroPedido + "!\n" + msg_erro_aux;
						Global.gravaLogAtividade(msg_erro_aux);
						// PROSSEGUE P/ O PRÓXIMO PEDIDO
						continue;
					}
					#endregion

					#region [ Obtém dados ]
					strOrder_FingerPrint_SessionID = "";
					strOrder_IP = "";
					strUsuario = "";
					strLoja = "";
					strOrder_Origin = "";
					intOwner = 0;
					vl_pedido = 0m;
					if (dtbPagPayment != null)
					{
						if (dtbPagPayment.Rows.Count > 0)
						{
							intOwner = BD.readToInt(dtbPagPayment.Rows[dtbPagPayment.Rows.Count - 1]["owner"]);
							strOrder_FingerPrint_SessionID = BD.readToString(dtbPagPayment.Rows[dtbPagPayment.Rows.Count - 1]["FingerPrint_SessionID"]);
							strOrder_IP = BD.readToString(dtbPagPayment.Rows[dtbPagPayment.Rows.Count - 1]["origem_endereco_IP"]);
							strUsuario = BD.readToString(dtbPagPayment.Rows[dtbPagPayment.Rows.Count - 1]["usuario"]);
							strLoja = BD.readToString(dtbPagPayment.Rows[dtbPagPayment.Rows.Count - 1]["loja"]);
							shortValue = BD.readToByte(dtbPagPayment.Rows[dtbPagPayment.Rows.Count - 1]["executado_pelo_cliente_status"]);
							strOrder_Origin = (shortValue == 0 ? "televendas" : "web");
							vl_pedido = BD.readToDecimal(dtbPagPayment.Rows[dtbPagPayment.Rows.Count - 1]["valor_pedido"]);
						}
					}
					#endregion

					#region [ Armazena dados p/ gravação no BD ]
					clearsaleAF.pedido = numeroPedido;
					// LEMBRANDO QUE O SUFIXO NSU DO PEDIDO USADO NA TABELA t_PAGTO_GW_PAG É UMA SEQUÊNCIA INDEPENDENTE DA SEQUÊNCIA USADA NA TABELA t_PAGTO_GW_AF
					clearsaleAF.pedido_com_sufixo_nsu = strOrder_ID;
					clearsaleAF.id_cliente = pedido.id_cliente;
					clearsaleAF.owner = intOwner;
					clearsaleAF.usuario = strUsuario;
					clearsaleAF.loja = strLoja;
					clearsaleAF.valor_pedido = vl_pedido;
					#endregion

					#region [ entityCode ]
					clearsaleAF.req_entityCode = Global.Cte.Clearsale.CS_ENTITY_CODE;
					#endregion

					#region [ Campo: Order/ID ]
					if (strOrder_ID.Length == 0) strOrder_ID = numeroPedido;
					strValue = strOrder_ID;
					sbXml.Append(xmlMontaCampo(strValue, "ID"));
					clearsaleAF.req_Order_ID = strValue;
					#endregion

					#region [ Campo: Order/FingerPrint/SessionID ]
					if (strOrder_FingerPrint_SessionID.Length > 0)
					{
						sbXml.Append("<FingerPrint>");
						strValue = strOrder_FingerPrint_SessionID;
						sbXml.Append(xmlMontaCampo(strValue, "SessionID"));
						clearsaleAF.req_Order_FingerPrint_SessionID = strValue;
						sbXml.Append("</FingerPrint>");
					}
					#endregion

					#region [ Campo: Order/Date ]
					strValue = Global.formataDataHoraYyyyMmDdTHhMmSs(pedido.data_hora);
					sbXml.Append(xmlMontaCampo(strValue, "Date"));
					clearsaleAF.req_Order_Date = strValue;
					#endregion

					#region [ Campo: Order/Email ]
					if (dtbPagPayment != null)
					{
						if (dtbPagPayment.Rows.Count > 0)
						{
							strValue = BD.readToString(dtbPagPayment.Rows[0]["checkout_email"]);
						}
					}
					else
					{
						strValue = cliente.email;
					}
					sbXml.Append(xmlMontaCampo(strValue, "Email"));
					clearsaleAF.req_Order_Email = strValue;
					#endregion

					#region [ Campo: Order/B2B_B2C ]
					strValue = "B2C";
					sbXml.Append(xmlMontaCampo(strValue, "B2B_B2C"));
					clearsaleAF.req_Order_B2B_B2C = strValue;
					#endregion

					#region [ Campo: Order/TotalItems ]
					strValue = Global.formataMoedaClearsale(pedido.vl_total_NF);
					sbXml.Append(xmlMontaCampo(strValue, "TotalItems"));
					clearsaleAF.req_Order_TotalItems = strValue;
					#endregion

					#region [ Campo: Order/TotalOrder ]
					strValue = Global.formataMoedaClearsale(pedido.vl_total_NF);
					sbXml.Append(xmlMontaCampo(strValue, "TotalOrder"));
					clearsaleAF.req_Order_TotalOrder = strValue;
					#endregion

					#region [ Campo: Order/QtyInstallments ]
					// Soma da quantidade de parcelas de todos os cartões utilizados no pagamento
					intQtyInstallments = 0;
					for (int i = 0; i < dtbPagPayment.Rows.Count; i++)
					{
						strValue = BD.readToString(dtbPagPayment.Rows[i]["req_PaymentDataRequest_NumberOfPayments"]);
						intQtyInstallments += (int)Global.converteInteiro(strValue);
					}
					strValue = intQtyInstallments.ToString();
					sbXml.Append(xmlMontaCampo(strValue, "QtyInstallments"));
					clearsaleAF.req_Order_QtyInstallments = strValue;
					#endregion

					#region [ Campo: Order/QtyItems ]
					intValue = 0;
					for (int i = 0; i < pedido.listaPedidoItem.Count; i++)
					{
						intValue += pedido.listaPedidoItem[i].qtde;
					}
					strValue = intValue.ToString();
					sbXml.Append(xmlMontaCampo(strValue, "QtyItems"));
					clearsaleAF.req_Order_QtyItems = strValue;
					#endregion

					#region [ Campo: Order/QtyPaymentTypes ]
					strValue = dtbPagPayment.Rows.Count.ToString();
					sbXml.Append(xmlMontaCampo(strValue, "QtyPaymentTypes"));
					clearsaleAF.req_Order_QtyPaymentTypes = strValue;
					#endregion

					#region [ Campo: Order/IP ]
					strValue = strOrder_IP;
					sbXml.Append(xmlMontaCampo(strValue, "IP"));
					clearsaleAF.req_Order_IP = strValue;
					#endregion

					#region [ Campo: Order/Origin ]
					strValue = strOrder_Origin;
					sbXml.Append(xmlMontaCampo(strValue, "Origin"));
					clearsaleAF.req_Order_Origin = strValue;
					#endregion

					#endregion

					#region [ Seção Order/BillingData ]
					sbXml.Append("<BillingData>");

					#region [ Campo: Order/BillingData/ID ]
					strValue = cliente.id;
					sbXml.Append(xmlMontaCampo(strValue, "ID"));
					clearsaleAF.req_Order_BillingData_ID = strValue;
					#endregion

					#region [ Campo: Order/BillingData/Type ]
					strValue = (cliente.tipo.Equals(Global.Cte.TipoPessoa.PF) ? "1" : "2");  // 1=Pessoa Física 2=Pessoa Jurídica
					sbXml.Append(xmlMontaCampo(strValue, "Type"));
					clearsaleAF.req_Order_BillingData_Type = strValue;
					#endregion

					#region [ Campo: Order/BillingData/LegalDocument1 ]
					strValue = cliente.cnpj_cpf;
					sbXml.Append(xmlMontaCampo(strValue, "LegalDocument1"));
					clearsaleAF.req_Order_BillingData_LegalDocument1 = strValue;
					#endregion

					#region [ Campo: Order/BillingData/LegalDocument2 ]
					strValue = "";
					if (cliente.tipo.Equals(Global.Cte.TipoPessoa.PF))
					{
						strValue = cliente.rg;
					}
					else
					{
						if (cliente.contribuinte_icms_status == 2) strValue = cliente.ie;  // contribuinte_icms_status = 2 -> Contribuinte de ICMS
					}

					if (strValue.Length > 0)
					{
						sbXml.Append(xmlMontaCampo(strValue, "LegalDocument2"));
						clearsaleAF.req_Order_BillingData_LegalDocument2 = strValue;
					}
					#endregion

					#region [ Campo: Order/BillingData/Name ]
					strValue = Global.filtraAmpersand(cliente.nome);
					sbXml.Append(xmlMontaCampo(strValue, "Name"));
					clearsaleAF.req_Order_BillingData_Name = strValue;
					#endregion

					#region [ Campo: Order/BillingData/BirthDate ]
					strValue = "";
					if (cliente.tipo.Equals(Global.Cte.TipoPessoa.PF))
					{
						if (cliente.dt_nasc != null)
						{
							if (cliente.dt_nasc != DateTime.MinValue)
							{
								strValue = Global.formataDataHoraYyyyMmDdTHhMmSs(cliente.dt_nasc);
							}
						}
					}
					if (strValue.Length > 0)
					{
						sbXml.Append(xmlMontaCampo(strValue, "BirthDate"));
						clearsaleAF.req_Order_BillingData_BirthDate = strValue;
					}
					#endregion

					#region [ Cammpo: Order/BillingData/Email ]
					strValue = cliente.email;
					sbXml.Append(xmlMontaCampo(strValue, "Email"));
					clearsaleAF.req_Order_BillingData_Email = strValue;
					#endregion

					#region [ Campo: Order/BillingData/Gender ]
					strValue = "";
					if (cliente.tipo.Equals(Global.Cte.TipoPessoa.PF))
					{
						strValue = cliente.sexo;
					}
					if (strValue.Length > 0)
					{
						sbXml.Append(xmlMontaCampo(strValue, "Gender"));
						clearsaleAF.req_Order_BillingData_Gender = strValue;
					}
					#endregion

					#region [ Seção: Order/BillingData/Address ]
					sbXml.Append("<Address>");

					#region [ Campo: Order/BillingData/Address/Street ]
					strValue = Global.filtraAmpersand(cliente.endereco);
					sbXml.Append(xmlMontaCampo(strValue, "Street"));
					clearsaleAF.req_Order_BillingData_Address_Street = strValue;
					#endregion

					#region [ Campo: Order/BillingData/Address/Number ]
					strValue = Global.filtraAmpersand(cliente.endereco_numero);
					sbXml.Append(xmlMontaCampo(strValue, "Number"));
					clearsaleAF.req_Order_BillingData_Address_Number = strValue;
					#endregion

					#region [ Campo: Order/BillingData/Address/Comp ]
					if (cliente.endereco_complemento.Length > 0)
					{
						strValue = Global.filtraAmpersand(cliente.endereco_complemento);
						sbXml.Append(xmlMontaCampo(strValue, "Comp"));
						clearsaleAF.req_Order_BillingData_Address_Comp = strValue;
					}
					#endregion

					#region [ Campo: Order/BillingData/Address/County ]
					strValue = Global.filtraAmpersand(cliente.bairro);
					sbXml.Append(xmlMontaCampo(strValue, "County"));
					clearsaleAF.req_Order_BillingData_Address_County = strValue;
					#endregion

					#region [ Campo: Order/BillingData/Address/City ]
					strValue = Global.filtraAmpersand(cliente.cidade);
					sbXml.Append(xmlMontaCampo(strValue, "City"));
					clearsaleAF.req_Order_BillingData_Address_City = strValue;
					#endregion

					#region [ Campo: Order/BillingData/Address/State ]
					strValue = cliente.uf;
					sbXml.Append(xmlMontaCampo(strValue, "State"));
					clearsaleAF.req_Order_BillingData_Address_State = strValue;
					#endregion

					#region [ Campo: Order/BillingData/Address/ZipCode ]
					strValue = cliente.cep;
					sbXml.Append(xmlMontaCampo(strValue, "ZipCode"));
					clearsaleAF.req_Order_BillingData_Address_ZipCode = strValue;
					#endregion

					sbXml.Append("</Address>");
					#endregion

					#region [ Seção: Order/BillingData/Phones ]
					sbXml.Append("<Phones>");

					// Tipo de telefone:
					//		0 = Não definido
					//		1 = Residencial
					//		2 = Comercial
					//		3 = Recados
					//		4 = Cobrança
					//		5 = Temporário
					//		6 = Celular
					if (cliente.tel_cel.Length > 0)
					{
						afPhone = new ClearsaleAFPhone();
						afPhone.idBlocoXml = Global.Cte.Clearsale.T_PAGTO_GW_AF_PHONE_IdBlocoXml.Order_BillingData_Phones.GetValue();
						sbXml.Append("<Phone>");

						strValue = "6";
						sbXml.Append(xmlMontaCampo(strValue, "Type"));
						afPhone.af_Type = strValue;

						strValue = cliente.ddd_cel;
						sbXml.Append(xmlMontaCampo(strValue, "DDD"));
						afPhone.af_DDD = strValue;

						strValue = cliente.tel_cel;
						sbXml.Append(xmlMontaCampo(strValue, "Number"));
						afPhone.af_Number = strValue;

						sbXml.Append("</Phone>");
						clearsaleAF.Order_BillingData_Phones.Add(afPhone);
					}

					if (cliente.tel_res.Length > 0)
					{
						afPhone = new ClearsaleAFPhone();
						afPhone.idBlocoXml = Global.Cte.Clearsale.T_PAGTO_GW_AF_PHONE_IdBlocoXml.Order_BillingData_Phones.GetValue();
						sbXml.Append("<Phone>");

						strValue = "1";
						sbXml.Append(xmlMontaCampo(strValue, "Type"));
						afPhone.af_Type = strValue;

						strValue = cliente.ddd_res;
						sbXml.Append(xmlMontaCampo(strValue, "DDD"));
						afPhone.af_DDD = strValue;

						strValue = cliente.tel_res;
						sbXml.Append(xmlMontaCampo(strValue, "Number"));
						afPhone.af_Number = strValue;

						sbXml.Append("</Phone>");
						clearsaleAF.Order_BillingData_Phones.Add(afPhone);
					}

					if (cliente.tel_com.Length > 0)
					{
						afPhone = new ClearsaleAFPhone();
						afPhone.idBlocoXml = Global.Cte.Clearsale.T_PAGTO_GW_AF_PHONE_IdBlocoXml.Order_BillingData_Phones.GetValue();
						sbXml.Append("<Phone>");

						strValue = "2";
						sbXml.Append(xmlMontaCampo(strValue, "Type"));
						afPhone.af_Type = strValue;

						strValue = cliente.ddd_com;
						sbXml.Append(xmlMontaCampo(strValue, "DDD"));
						afPhone.af_DDD = strValue;

						strValue = cliente.tel_com;
						sbXml.Append(xmlMontaCampo(strValue, "Number"));
						afPhone.af_Number = strValue;

						if (cliente.ramal_com.Length > 0)
						{
							strValue = cliente.ramal_com;
							sbXml.Append(xmlMontaCampo(strValue, "Extension"));
							afPhone.af_Extension = strValue;
						}
						sbXml.Append("</Phone>");
						clearsaleAF.Order_BillingData_Phones.Add(afPhone);
					}

					if (cliente.tel_com_2.Length > 0)
					{
						afPhone = new ClearsaleAFPhone();
						afPhone.idBlocoXml = Global.Cte.Clearsale.T_PAGTO_GW_AF_PHONE_IdBlocoXml.Order_BillingData_Phones.GetValue();
						sbXml.Append("<Phone>");

						strValue = "2";
						sbXml.Append(xmlMontaCampo(strValue, "Type"));
						afPhone.af_Type = strValue;

						strValue = cliente.ddd_com_2;
						sbXml.Append(xmlMontaCampo(strValue, "DDD"));
						afPhone.af_DDD = strValue;

						strValue = cliente.tel_com_2;
						sbXml.Append(xmlMontaCampo(strValue, "Number"));
						afPhone.af_Number = strValue;

						if (cliente.ramal_com_2.Length > 0)
						{
							strValue = cliente.ramal_com_2;
							sbXml.Append(xmlMontaCampo(strValue, "Extension"));
							afPhone.af_Extension = strValue;
						}

						sbXml.Append("</Phone>");
						clearsaleAF.Order_BillingData_Phones.Add(afPhone);
					}

					sbXml.Append("</Phones>");
					#endregion

					sbXml.Append("</BillingData>");
					#endregion

					#region [ Seção Order/ShippingData ]
					sbXml.Append("<ShippingData>");

					#region [ Campo: Order/ShippingData/ID ]
					strValue = cliente.id;
					sbXml.Append(xmlMontaCampo(strValue, "ID"));
					clearsaleAF.req_Order_ShippingData_ID = strValue;
					#endregion

					#region [ Campo: Order/ShippingData/Type ]
					strValue = (cliente.tipo.Equals(Global.Cte.TipoPessoa.PF) ? "1" : "2");  // 1=Pessoa Física 2=Pessoa Jurídica
					sbXml.Append(xmlMontaCampo(strValue, "Type"));
					clearsaleAF.req_Order_ShippingData_Type = strValue;
					#endregion

					#region [ Campo: Order/ShippingData/LegalDocument1 ]
					strValue = cliente.cnpj_cpf;
					sbXml.Append(xmlMontaCampo(cliente.cnpj_cpf, "LegalDocument1"));
					clearsaleAF.req_Order_ShippingData_LegalDocument1 = strValue;
					#endregion

					#region [ Campo: Order/ShippingData/LegalDocument2 ]
					strValue = "";
					if (cliente.tipo.Equals(Global.Cte.TipoPessoa.PF))
					{
						strValue = cliente.rg;
					}
					else
					{
						if (cliente.contribuinte_icms_status == 2) strValue = cliente.ie;  // contribuinte_icms_status = 2 -> Contribuinte de ICMS
					}

					if (strValue.Length > 0)
					{
						sbXml.Append(xmlMontaCampo(strValue, "LegalDocument2"));
						clearsaleAF.req_Order_ShippingData_LegalDocument2 = strValue;
					}
					#endregion

					#region [ Campo: Order/ShippingData/Name ]
					strValue = Global.filtraAmpersand(cliente.nome);
					sbXml.Append(xmlMontaCampo(strValue, "Name"));
					clearsaleAF.req_Order_ShippingData_Name = strValue;
					#endregion

					#region [ Campo: Order/ShippingData/BirthDate ]
					strValue = "";
					if (cliente.tipo.Equals(Global.Cte.TipoPessoa.PF))
					{
						if (cliente.dt_nasc != null)
						{
							if (cliente.dt_nasc != DateTime.MinValue)
							{
								strValue = Global.formataDataHoraYyyyMmDdTHhMmSs(cliente.dt_nasc);
							}
						}
					}
					if (strValue.Length > 0)
					{
						sbXml.Append(xmlMontaCampo(strValue, "BirthDate"));
						clearsaleAF.req_Order_ShippingData_BirthDate = strValue;
					}
					#endregion

					#region [ Cammpo: Order/ShippingData/Email ]
					strValue = cliente.email;
					sbXml.Append(xmlMontaCampo(strValue, "Email"));
					clearsaleAF.req_Order_ShippingData_Email = strValue;
					#endregion

					#region [ Campo: Order/ShippingData/Gender ]
					strValue = "";
					if (cliente.tipo.Equals(Global.Cte.TipoPessoa.PF))
					{
						strValue = cliente.sexo;
					}
					if (strValue.Length > 0)
					{
						sbXml.Append(xmlMontaCampo(strValue, "Gender"));
						clearsaleAF.req_Order_ShippingData_Gender = strValue;
					}
					#endregion

					#region [ Seção: Order/ShippingData/Address ]
					sbXml.Append("<Address>");

					if (pedido.st_end_entrega != 0)
					{
						#region [ Campo: Order/ShippingData/Address/Street ]
						strValue = Global.filtraAmpersand(pedido.endEtg_endereco);
						sbXml.Append(xmlMontaCampo(strValue, "Street"));
						clearsaleAF.req_Order_ShippingData_Address_Street = strValue;
						#endregion

						#region [ Campo: Order/ShippingData/Address/Number ]
						strValue = Global.filtraAmpersand(pedido.endEtg_endereco_numero);
						sbXml.Append(xmlMontaCampo(strValue, "Number"));
						clearsaleAF.req_Order_ShippingData_Address_Number = strValue;
						#endregion

						#region [ Campo: Order/ShippingData/Address/Comp ]
						if (pedido.endEtg_endereco_complemento.Length > 0)
						{
							strValue = Global.filtraAmpersand(pedido.endEtg_endereco_complemento);
							sbXml.Append(xmlMontaCampo(strValue, "Comp"));
							clearsaleAF.req_Order_ShippingData_Address_Comp = strValue;
						}
						#endregion

						#region [ Campo: Order/ShippingData/Address/County ]
						strValue = Global.filtraAmpersand(pedido.endEtg_bairro);
						sbXml.Append(xmlMontaCampo(strValue, "County"));
						clearsaleAF.req_Order_ShippingData_Address_County = strValue;
						#endregion

						#region [ Campo: Order/ShippingData/Address/City ]
						strValue = Global.filtraAmpersand(pedido.endEtg_cidade);
						sbXml.Append(xmlMontaCampo(strValue, "City"));
						clearsaleAF.req_Order_ShippingData_Address_City = strValue;
						#endregion

						#region [ Campo: Order/ShippingData/Address/State ]
						strValue = pedido.endEtg_uf;
						sbXml.Append(xmlMontaCampo(strValue, "State"));
						clearsaleAF.req_Order_ShippingData_Address_State = strValue;
						#endregion

						#region [ Campo: Order/ShippingData/Address/ZipCode ]
						strValue = pedido.endEtg_cep;
						sbXml.Append(xmlMontaCampo(strValue, "ZipCode"));
						clearsaleAF.req_Order_ShippingData_Address_ZipCode = strValue;
						#endregion
					}
					else
					{
						#region [ Campo: Order/ShippingData/Address/Street ]
						strValue = Global.filtraAmpersand(cliente.endereco);
						sbXml.Append(xmlMontaCampo(strValue, "Street"));
						clearsaleAF.req_Order_ShippingData_Address_Street = strValue;
						#endregion

						#region [ Campo: Order/ShippingData/Address/Number ]
						strValue = Global.filtraAmpersand(cliente.endereco_numero);
						sbXml.Append(xmlMontaCampo(strValue, "Number"));
						clearsaleAF.req_Order_ShippingData_Address_Number = strValue;
						#endregion

						#region [ Campo: Order/ShippingData/Address/Comp ]
						if (cliente.endereco_complemento.Length > 0)
						{
							strValue = Global.filtraAmpersand(cliente.endereco_complemento);
							sbXml.Append(xmlMontaCampo(strValue, "Comp"));
							clearsaleAF.req_Order_ShippingData_Address_Comp = strValue;
						}
						#endregion

						#region [ Campo: Order/ShippingData/Address/County ]
						strValue = Global.filtraAmpersand(cliente.bairro);
						sbXml.Append(xmlMontaCampo(strValue, "County"));
						clearsaleAF.req_Order_ShippingData_Address_County = strValue;
						#endregion

						#region [ Campo: Order/ShippingData/Address/City ]
						strValue = Global.filtraAmpersand(cliente.cidade);
						sbXml.Append(xmlMontaCampo(strValue, "City"));
						clearsaleAF.req_Order_ShippingData_Address_City = strValue;
						#endregion

						#region [ Campo: Order/ShippingData/Address/State ]
						strValue = cliente.uf;
						sbXml.Append(xmlMontaCampo(strValue, "State"));
						clearsaleAF.req_Order_ShippingData_Address_State = strValue;
						#endregion

						#region [ Campo: Order/ShippingData/Address/ZipCode ]
						strValue = cliente.cep;
						sbXml.Append(xmlMontaCampo(strValue, "ZipCode"));
						clearsaleAF.req_Order_ShippingData_Address_ZipCode = strValue;
						#endregion
					}

					sbXml.Append("</Address>");
					#endregion

					#region [ Seção: Order/ShippingData/Phones ]
					sbXml.Append("<Phones>");

					// Tipo de telefone:
					//		0 = Não definido
					//		1 = Residencial
					//		2 = Comercial
					//		3 = Recados
					//		4 = Cobrança
					//		5 = Temporário
					//		6 = Celular
					if (cliente.tel_cel.Length > 0)
					{
						afPhone = new ClearsaleAFPhone();
						afPhone.idBlocoXml = Global.Cte.Clearsale.T_PAGTO_GW_AF_PHONE_IdBlocoXml.Order_ShippingData_Phones.GetValue();
						sbXml.Append("<Phone>");

						strValue = "6";
						sbXml.Append(xmlMontaCampo(strValue, "Type"));
						afPhone.af_Type = strValue;

						strValue = cliente.ddd_cel;
						sbXml.Append(xmlMontaCampo(strValue, "DDD"));
						afPhone.af_DDD = strValue;

						strValue = cliente.tel_cel;
						sbXml.Append(xmlMontaCampo(strValue, "Number"));
						afPhone.af_Number = strValue;

						sbXml.Append("</Phone>");
						clearsaleAF.Order_ShippingData_Phones.Add(afPhone);
					}

					if (cliente.tel_res.Length > 0)
					{
						afPhone = new ClearsaleAFPhone();
						afPhone.idBlocoXml = Global.Cte.Clearsale.T_PAGTO_GW_AF_PHONE_IdBlocoXml.Order_ShippingData_Phones.GetValue();
						sbXml.Append("<Phone>");

						strValue = "1";
						sbXml.Append(xmlMontaCampo(strValue, "Type"));
						afPhone.af_Type = strValue;

						strValue = cliente.ddd_res;
						sbXml.Append(xmlMontaCampo(strValue, "DDD"));
						afPhone.af_DDD = strValue;

						strValue = cliente.tel_res;
						sbXml.Append(xmlMontaCampo(strValue, "Number"));
						afPhone.af_Number = strValue;

						sbXml.Append("</Phone>");
						clearsaleAF.Order_ShippingData_Phones.Add(afPhone);
					}

					if (cliente.tel_com.Length > 0)
					{
						afPhone = new ClearsaleAFPhone();
						afPhone.idBlocoXml = Global.Cte.Clearsale.T_PAGTO_GW_AF_PHONE_IdBlocoXml.Order_ShippingData_Phones.GetValue();
						sbXml.Append("<Phone>");

						strValue = "2";
						sbXml.Append(xmlMontaCampo(strValue, "Type"));
						afPhone.af_Type = strValue;

						strValue = cliente.ddd_com;
						sbXml.Append(xmlMontaCampo(strValue, "DDD"));
						afPhone.af_DDD = strValue;

						strValue = cliente.tel_com;
						sbXml.Append(xmlMontaCampo(strValue, "Number"));
						afPhone.af_Number = strValue;

						if (cliente.ramal_com.Length > 0)
						{
							strValue = cliente.ramal_com;
							sbXml.Append(xmlMontaCampo(strValue, "Extension"));
							afPhone.af_Extension = strValue;
						}

						sbXml.Append("</Phone>");
						clearsaleAF.Order_ShippingData_Phones.Add(afPhone);
					}

					if (cliente.tel_com_2.Length > 0)
					{
						afPhone = new ClearsaleAFPhone();
						afPhone.idBlocoXml = Global.Cte.Clearsale.T_PAGTO_GW_AF_PHONE_IdBlocoXml.Order_ShippingData_Phones.GetValue();
						sbXml.Append("<Phone>");

						strValue = "2";
						sbXml.Append(xmlMontaCampo(strValue, "Type"));
						afPhone.af_Type = strValue;

						strValue = cliente.ddd_com_2;
						sbXml.Append(xmlMontaCampo(strValue, "DDD"));
						afPhone.af_DDD = strValue;

						strValue = cliente.tel_com_2;
						sbXml.Append(xmlMontaCampo(strValue, "Number"));
						afPhone.af_Number = strValue;

						if (cliente.ramal_com_2.Length > 0)
						{
							strValue = cliente.ramal_com_2;
							sbXml.Append(xmlMontaCampo(strValue, "Extension"));
							afPhone.af_Extension = strValue;
						}

						sbXml.Append("</Phone>");
						clearsaleAF.Order_ShippingData_Phones.Add(afPhone);
					}

					sbXml.Append("</Phones>");
					#endregion

					sbXml.Append("</ShippingData>");
					#endregion

					#region [ Seção: Order/Payments ]
					// Importante: é fundamental que os dados de 't_PAGTO_GW_PAG_PAYMENT.id' dos pagamentos
					// enviados p/ análise antifraude sejam "memorizados" p/ posterior atualização de status.
					// Caso o cliente realize um pagamento durante este processamento, isso evita que o
					// novo pagamento seja assinalado indevidamente como tendo sido enviado p/ análise antifraude.
					sbXml.Append("<Payments>");
					if (dtbPagPayment != null)
					{
						for (int i = 0; i < dtbPagPayment.Rows.Count; i++)
						{
							sbXml.Append(montaXmlPaymentAF((i + 1), dtbPagPayment.Rows[i], out afPayment));
							clearsaleAF.Payments.Add(afPayment);
						}
					}
					sbXml.Append("</Payments>");
					#endregion

					#region [ Seção: Order/Items ]
					sbXml.Append("<Items>");
					for (int i = 0; i < pedido.listaPedidoItem.Count; i++)
					{
						sbXml.Append(montaXmlItem(pedido.listaPedidoItem[i], out afItem));
						clearsaleAF.Items.Add(afItem);
					}
					sbXml.Append("</Items>");
					#endregion

					sbXml.Append("</Order>");
					sbXml.Append("</Orders>");
					sbXml.Append("</ClearSale>");

					#endregion

					#region [ Gera o NSU do registro principal ]
					blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.FIN.NSU.T_PAGTO_GW_AF, out idPagtoGwAf, out msg_erro_aux);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro principal dos dados da análise antifraude!!\n" + msg_erro_aux;
						return false;
					}
					clearsaleAF.id = idPagtoGwAf;
					#endregion

					#region [ Prepara os dados do XML da transação de envio p/ armazenar no BD ]
					xmlReqSoap = montaRequisicaoSoapSendOrders(clearsaleAF.req_entityCode, sbXml.ToString());
					clearsaleAFXml = new ClearsaleAFXml();
					clearsaleAFXml.id_pagto_gw_af = idPagtoGwAf;
					clearsaleAFXml.tipo_transacao = Global.Cte.Clearsale.Transacao.SendOrders.GetCodOpLog();
					clearsaleAFXml.fluxo_xml = Global.Cte.FluxoXml.TX.GetValue();
					clearsaleAFXml.xml = xmlReqSoap;
					#endregion

					#region [ Se houver dados de AF de tentativa anterior, anula os registros ]
					// Esta situação ocorre se em um processamento anterior os dados foram gravados no BD com sucesso, mas ocorreu falha no envio para a Clearsale
					// Importante: a abordagem de anular os registros anteriores e gravar registros novos nesta execução é para garantir que os dados gravados
					// =========== no BD sejam exatamente iguais aos enviados à Clearsale, já que entre as execuções desta rotina é possível que algum dos registros
					// de t_PAGTO_GW_PAG_PAYMENT pode ter sido cancelado (operação 'Void' ou 'Refund') quando o pagamento é feito usando mais do que um cartão.
					if (!ClearsaleDAO.anulaRegistroAFTentativaAnterior(numeroPedido, idPagtoGwAf, out intQtdeAnulados, out msg_erro_aux))
					{
						msg_erro = "Falha ao tentar anular dados de tentativa anterior!!\n" + msg_erro_aux;
						return false;
					}
					#endregion

					#region [ Grava dados no BD ]
					blnSucesso = false;
					try
					{
						BD.iniciaTransacao();

						if (!ClearsaleDAO.insereAF(clearsaleAF, out msg_erro_aux))
						{
							msg_erro = msg_erro_aux;
							return false;
						}

						if (!ClearsaleDAO.insereAFXml(clearsaleAFXml, out msg_erro_aux))
						{
							msg_erro = msg_erro_aux;
							return false;
						}
						else
						{
							#region [ Armazena o ID do registro do XML em t_PAGTO_GW_AF.trx_TX_id_pagto_gw_af_xml ]
							ClearsaleDAO.updateRegistroAFTxXml(idPagtoGwAf, clearsaleAFXml.id, out msg_erro_aux);
							#endregion
						}

						blnSucesso = true;
					}
					catch (Exception ex)
					{
						Global.gravaLogAtividade(ex.ToString());
						msg_erro = ex.ToString();
						blnSucesso = false;
					}
					finally
					{
						if (blnSucesso)
						{
							BD.commitTransacao();
						}
						else
						{
							BD.rollbackTransacao();
						}
					}
					#endregion

					#region [ Transmite XML para Clearsale ]
					blnEnviouOk = enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.Clearsale.Transacao.SendOrders, out xmlRespSoap, out msg_erro_aux);
					if (!blnEnviouOk)
					{
						strTrxErroMensagem = msg_erro_aux;
						// Registra o erro em log, mas prossegue com as próximas transações
						msg_erro_aux = "Falha ao tentar enviar transação para a Clearsale: " + Global.Cte.Clearsale.Transacao.SendOrders.GetMethodName() + "!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(msg_erro_aux);

						#region [ Grava o erro no registro da transação no BD ]
						byteVazioStatus = 1;
						if (xmlRespSoap != null)
						{
							if (xmlRespSoap.Trim().Length > 0) byteVazioStatus = 0;
						}
						ClearsaleDAO.updateRegistroAFRespostaErro(idPagtoGwAf, byteVazioStatus, 1, null, strTrxErroMensagem, out msg_erro_aux);
						#endregion
					}
					#endregion

					#region [ Tratamento para transação enviada com sucesso ]
					if (blnEnviouOk)
					{
						sendOrdersResponse = new ClearsaleSendOrdersResponse();

						#region [ Armazena o XML de retorno no BD ]
						clearsaleAFXml = new ClearsaleAFXml();
						clearsaleAFXml.id_pagto_gw_af = idPagtoGwAf;
						clearsaleAFXml.tipo_transacao = Global.Cte.Clearsale.Transacao.SendOrders.GetCodOpLog();
						clearsaleAFXml.fluxo_xml = Global.Cte.FluxoXml.RX.GetValue();
						clearsaleAFXml.xml = xmlRespSoap;
						if (ClearsaleDAO.insereAFXml(clearsaleAFXml, out msg_erro_aux))
						{
							#region [ Armazena o ID do registro do XML em t_PAGTO_GW_AF.trx_RX_id_pagto_gw_af_xml ]
							byteVazioStatus = 1;
							if (xmlRespSoap != null)
							{
								if (xmlRespSoap.Trim().Length > 0) byteVazioStatus = 0;
							}
							ClearsaleDAO.updateRegistroAFRxXml(idPagtoGwAf, clearsaleAFXml.id, byteVazioStatus, out msg_erro_aux);
							#endregion
						}
						#endregion

						#region [ Analisa o StatusCode da resposta ]
						strPackageStatus_StatusCode = "";
						xmlDoc = new XmlDocument();
						xmlDoc.LoadXml(xmlRespSoap);
						nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
						nsmgr.AddNamespace("cs", "http://www.clearsale.com.br/integration");
						xmlNode = xmlDoc.SelectSingleNode("//cs:SendOrdersResult", nsmgr);
						blnXmlInvalido = false;
						if (xmlNode == null) blnXmlInvalido = true;
						if (!blnXmlInvalido)
						{
							if (xmlNode.ChildNodes == null) blnXmlInvalido = true;
						}

						if (!blnXmlInvalido)
						{
							if (xmlNode.ChildNodes.Count > 0)
							{
								strValue = HttpUtility.HtmlDecode(xmlNode.FirstChild.Value);
								xmlDocSendOrdersResult = new XmlDocument();
								xmlDocSendOrdersResult.LoadXml(strValue);
								xmlNode = xmlDocSendOrdersResult.SelectSingleNode("//PackageStatus/TransactionID");
								if (xmlNode != null) sendOrdersResponse.TransactionID = (xmlNode.FirstChild != null ? (xmlNode.FirstChild.Value ?? "") : "");
								xmlNode = xmlDocSendOrdersResult.SelectSingleNode("//PackageStatus/Message");
								if (xmlNode != null) sendOrdersResponse.Message = (xmlNode.FirstChild != null ? (xmlNode.FirstChild.Value ?? "") : "");
								xmlNode = xmlDocSendOrdersResult.SelectSingleNode("//PackageStatus/StatusCode");
								if (xmlNode != null) strPackageStatus_StatusCode = (xmlNode.FirstChild != null ? (xmlNode.FirstChild.Value ?? "") : "");
								sendOrdersResponse.StatusCode = strPackageStatus_StatusCode;
								if (strPackageStatus_StatusCode.Equals(Global.Cte.Clearsale.PackageStatus_StatusCode.TRANSACAO_CONCLUIDA.GetValue()))
								{
									xmlNodeList = xmlDocSendOrdersResult.SelectNodes("//PackageStatus/Orders/Order");
									if (xmlNodeList != null)
									{
										foreach (XmlNode node in xmlNodeList)
										{
											responseOrder = new ClearsaleSendOrdersResponseOrder(Global.obtemXmlChildNodeValue(node, "ID"), Global.obtemXmlChildNodeValue(node, "Status"), Global.obtemXmlChildNodeValue(node, "Score"));
											if ((responseOrder.ID ?? "").Length > 0) sendOrdersResponse.Orders.Add(responseOrder);
										}
									}
								}

								#region [ Armazena os dados no BD ]
								for (int i = 0; i < sendOrdersResponse.Orders.Count; i++)
								{
									if (!ClearsaleDAO.updateRegistroAFSendOrdersResponse(sendOrdersResponse.TransactionID, sendOrdersResponse.StatusCode, sendOrdersResponse.Message, sendOrdersResponse.Orders[i], clearsaleAF.req_entityCode, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar atualizar os dados do pedido " + sendOrdersResponse.Orders[i].ID + " devido à resposta obtida pela requisição GetReturnAnalysis()!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
										strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Clearsale: Falha ao tentar atualizar os dados do pedido " + sendOrdersResponse.Orders[i].ID + " devido à resposta obtida pela requisição GetReturnAnalysis() [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
										strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar os dados do pedido " + sendOrdersResponse.Orders[i].ID + " devido à resposta obtida pela requisição GetReturnAnalysis()\r\n" + msg_erro_aux;
										if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
										{
											strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
											Global.gravaLogAtividade(strMsg);
										}
									}
								}
								#endregion
							}
						}
						#endregion

						#region [ Registra o sucesso do envio em t_PAGTO_GW_PAG_PAYMENT.st_enviado_analise_AF ]
						if (dtbPagPayment != null)
						{
							blnExecutar = false;
							if (strPackageStatus_StatusCode.Equals(Global.Cte.Clearsale.PackageStatus_StatusCode.TRANSACAO_CONCLUIDA.GetValue())) blnExecutar = true;
							// Caso o StatusCode retorne como "Pedido já enviado ou não está em reanálise", marca os pagamentos como já enviados p/ o AF,
							// pois pode ter ocorrido de que em uma tentativa anterior a transação tenha sido enviada com sucesso, mas ter ocorrido falha
							// na atualização dos dados no BD.
							if (strPackageStatus_StatusCode.Equals(Global.Cte.Clearsale.PackageStatus_StatusCode.PEDIDO_JA_ENVIADO.GetValue())) blnExecutar = true;

							if (blnExecutar)
							{
								for (int i = 0; i < dtbPagPayment.Rows.Count; i++)
								{
									idPagtoGwPagPayment = BD.readToInt(dtbPagPayment.Rows[i]["id"]);
									ClearsaleDAO.updatePagPaymentStEnviadoAnaliseAF(idPagtoGwPagPayment, idPagtoGwAf, out msg_erro_aux);
								}
							}
						}
						#endregion

						#region [ Sucesso ou falha no envio? ]
						blnStatusRespSucesso = false;
						if (strPackageStatus_StatusCode.Equals(Global.Cte.Clearsale.PackageStatus_StatusCode.TRANSACAO_CONCLUIDA.GetValue())) blnStatusRespSucesso = true;
						if (strPackageStatus_StatusCode.Equals(Global.Cte.Clearsale.PackageStatus_StatusCode.PEDIDO_JA_ENVIADO.GetValue())) blnStatusRespSucesso = true;

						if (blnStatusRespSucesso)
						{
							qtdePedidosNovosEnviados++;
							if (sbPedidosEnviadosOk.Length > 0) sbPedidosEnviadosOk.Append(", ");
							sbPedidosEnviadosOk.Append(strOrder_ID);
						}
						else
						{
							qtdePedidosFalhaEnvio++;
							if (sbPedidosFalhaEnvio.Length > 0) sbPedidosFalhaEnvio.Append(", ");
							sbPedidosFalhaEnvio.Append(strOrder_ID);
						}
						#endregion

						#region [ Tratamento para caso de resposta informando erro ]
						if (!blnStatusRespSucesso)
						{
							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							strMsg = "Falha no envio do pedido para a Clearsale: " + strOrder_ID +
									((sendOrdersResponse.StatusCode ?? "").Length > 0 ? "\r\nStatusCode: " + sendOrdersResponse.StatusCode : "") +
									((sendOrdersResponse.Message ?? "").Length > 0 ? "\r\nMessage: " + sendOrdersResponse.Message : "");
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = strMsg;
							svcLog.complemento_1 = xmlReqSoap;
							svcLog.complemento_2 = xmlRespSoap;
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							#region [ Envia email de alerta ]
							strSubject = Global.montaIdInstanciaServicoEmailSubject() +
										" Clearsale: Falha ao tentar enviar o pedido " + strOrder_ID +
										" [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Falha ao tentar enviar o pedido " + strOrder_ID + " para a Clearsale!" +
										((sendOrdersResponse.StatusCode ?? "").Length > 0 ? "\r\nStatusCode: " + sendOrdersResponse.StatusCode : "") +
										((sendOrdersResponse.Message ?? "").Length > 0 ? "\r\nMessage: " + sendOrdersResponse.Message : "") +
										((xmlReqSoap ?? "").Length > 0 ? "\r\n\r\nRequisição:\r\n" + xmlReqSoap : "") +
										((xmlRespSoap ?? "").Length > 0 ? "\r\n\r\nResposta:\r\n" + xmlRespSoap : "");
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}
							#endregion
						}
						#endregion
					}  // if (blnEnviouOk)
					#endregion

					#endregion
				}  // for (int iLTP = 0; iLTP < listaPedidoTrxPag.Count; iLTP++)
				#endregion

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				if ((qtdePedidosNovosEnviados + qtdePedidosFalhaEnvio) > 0)
				{
					strMsg = "Envio de pedidos para a Clearsale (sucesso: " + qtdePedidosNovosEnviados.ToString() + ", falha: " + qtdePedidosFalhaEnvio.ToString() + "): " +
							"Sucesso = " + (sbPedidosEnviadosOk.Length > 0 ? sbPedidosEnviadosOk.ToString() : "(nenhum)") +
							"; Falha = " + (sbPedidosFalhaEnvio.Length > 0 ? sbPedidosFalhaEnvio.ToString() : "(nenhum)");
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = strMsg;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
			finally
			{
				if (sbPedidosEnviadosOk.Length > 0)
				{
					if (sbMsgInformativa.Length > 0) sbMsgInformativa.Append("; ");
					sbMsgInformativa.Append("Pedidos enviados com sucesso: " + sbPedidosEnviadosOk.ToString());
				}

				if (sbPedidosFalhaEnvio.Length > 0)
				{
					if (sbMsgInformativa.Length > 0) sbMsgInformativa.Append("; ");
					sbMsgInformativa.Append("Pedidos com falha no envio: " + sbPedidosFalhaEnvio.ToString());
				}

				strMsgInformativa = sbMsgInformativa.ToString();
			}
		}
		#endregion

		#region [ processaResultadoAntifraude ]
		private static bool processaResultadoAntifraude(out int qtdePedidosResultadoProcessado, out string strMsgInformativa, out string msg_erro)
		{
			#region Declarações
			const String NOME_DESTA_ROTINA = "Clearsale.processaResultadoAntifraude()";
			bool blnErroFinalizacaoBraspag;
			bool blnAssinalarPaymentAFFinalizado;
			bool blnResultadoSetOrderAsReturned;
			int id_emailsndsvc_mensagem;
			string msg_erro_aux;
			string msg_erro_temp;
			string entityCode;
			string xmlReqSoap;
			string xmlRespSoap;
			string strSubject;
			string strBody;
			string strValue;
			string strMsg;
			string strBlocoNotas;
			string ult_GlobalStatus_original;
			string ult_GlobalStatus_novo;
			StringBuilder sbMsgAprovados;
			StringBuilder sbMsgReprovados;
			StringBuilder sbMsgOutrosStatus;
			StringBuilder sbAlertaStatusErro = new StringBuilder("");
			StringBuilder sbLogNothingToDo;
			StringBuilder sbErroFinalizacaoBraspag;
			StringBuilder sbMsgInformativa = new StringBuilder("");
			DataTable dtbTrxPag = new DataTable();
			List<ClearsaleGetReturnAnalysisResponse> listaReturnAnalysis;
			ClearsaleGetReturnAnalysisResponse order;
			XmlDocument xmlDoc;
			XmlDocument xmlDocGetReturnAnalysisResult;
			XmlNode xmlNode;
			XmlNodeList xmlNodeList;
			XmlNamespaceManager nsmgr;
			ClearsaleAF clearsaleAF;
			BraspagPag braspagPag;
			BraspagPagPayment braspagPayment;
			BraspagPagPayment braspagPaymentAtualizado;
			BraspagPagPaymentFinalizacao rowFinalizacao;
			List<BraspagPagPaymentFinalizacao> listaBraspagPaymentFinalizacao;
			Global.Cte.Braspag.Pagador.Transacao trxFinalizacaoBraspag;
			FinSvcLog svcLogMsg;
			List<ClearsaleAnalystComments> listaComments;
			List<ClearsaleAnalystComments> listaCommentsOrderByCreateDate;
			#endregion

			qtdePedidosResultadoProcessado = 0;
			strMsgInformativa = "";
			msg_erro = "";

			try
			{
				strMsg = "Rotina " + NOME_DESTA_ROTINA + " iniciada";
				Global.gravaLogAtividade(strMsg);

				#region [ Consulta Clearsale ]
				entityCode = Global.Cte.Clearsale.CS_ENTITY_CODE;
				xmlReqSoap = montaRequisicaoSoapGetReturnAnalysis(entityCode);
				if (!enviaRequisicao(xmlReqSoap, Global.Cte.Clearsale.Transacao.GetReturnAnalysis, out xmlRespSoap, out msg_erro_aux))
				{
					qtdeFalhasConsecutivasMetodoGetReturnAnalysis++;
					strMsg = "Falha ao consultar a Clearsale: método " + Global.Cte.Clearsale.Transacao.GetReturnAnalysis.GetMethodName() + "\n" + "Qtde falhas consecutivas: " + qtdeFalhasConsecutivasMetodoGetReturnAnalysis.ToString() + "\n" + msg_erro_aux;
					Global.gravaLogAtividade(strMsg);
					if (qtdeFalhasConsecutivasMetodoGetReturnAnalysis >= Global.Parametros.Clearsale.MaxQtdeFalhasConsecutivasMetodoGetReturnAnalysis)
					{
						#region [ Envia email alertando sobre o problema ]
						if ((qtdeFalhasConsecutivasMetodoGetReturnAnalysis % Global.Parametros.Clearsale.MaxQtdeFalhasConsecutivasMetodoGetReturnAnalysis) == 0)
						{
							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Clearsale: quantidade de falhas consecutivas na chamada ao método GetReturnAnalysis(): " + qtdeFalhasConsecutivasMetodoGetReturnAnalysis.ToString() + " falhas [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nQuantidade de falhas consecutivas na chamada ao método da Clearsale GetReturnAnalysis(): " + qtdeFalhasConsecutivasMetodoGetReturnAnalysis.ToString() + " falhas";
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}
						}
						#endregion
					}
					msg_erro = "Falha ao tentar consultar o método GetReturnAnalysis()!\n" + msg_erro_aux;
					return false;
				}
				#endregion

				#region [ XML: resposta válida? ]
				if (xmlRespSoap == null)
				{
					msg_erro = "A resposta é nula!";
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + msg_erro);
					return false;
				}
				if (xmlRespSoap.Trim().Length == 0)
				{
					msg_erro = "A resposta está vazia!";
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + msg_erro);
					return false;
				}
				#endregion

				qtdeFalhasConsecutivasMetodoGetReturnAnalysis = 0;

				#region [ Carrega dados do XML de resposta em uma lista ]
				listaReturnAnalysis = new List<ClearsaleGetReturnAnalysisResponse>();
				xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(xmlRespSoap);
				nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
				nsmgr.AddNamespace("cs", "http://www.clearsale.com.br/integration");
				xmlNode = xmlDoc.SelectSingleNode("//cs:GetReturnAnalysisResult", nsmgr);
				if (xmlNode.ChildNodes.Count > 0)
				{
					strValue = HttpUtility.HtmlDecode(xmlNode.FirstChild.Value);
					xmlDocGetReturnAnalysisResult = new XmlDocument();
					xmlDocGetReturnAnalysisResult.LoadXml(strValue);
					xmlNodeList = xmlDocGetReturnAnalysisResult.SelectNodes("//Orders/Order");

					foreach (XmlNode node in xmlNodeList)
					{
						order = new ClearsaleGetReturnAnalysisResponse(Global.obtemXmlChildNodeValue(node, "ID"), Global.obtemXmlChildNodeValue(node, "Status"), Global.obtemXmlChildNodeValue(node, "Score"));
						// Trata somente os pedidos do sistema (ERP), ou seja, caso existam pedidos da plataforma de e-commerce, ignora-os.
						// Importante: é necessário verificar se o pedido está cadastrado no ambiente ao qual o serviço está conectado (DIS ou OLD01). Além disso, é importante
						// =========== lembrar que o nº do pedido enviado p/ a Clearsale pode conter um sufixo, caso tenham sido enviados mais do que uma requisição de análise AF. Esta situação
						// pode ocorrer no caso de pagamento com múltiplos cartões em que uma das transações foi negada pelo Pagador e o cliente não fez outra em substituição dentro do tempo
						// limite que o serviço aguarda p/ integralizar o pagamento antes de enviar a requisição AF p/ a Clearsale.
						if (ClearsaleDAO.isPedidoERPDesteAmbiente(order.ID, entityCode))
						{
							listaReturnAnalysis.Add(order);
						}
					}
				}
				#endregion

				#region [ Grava log sobre os dados da resposta ]
				sbMsgAprovados = new StringBuilder("");
				sbMsgReprovados = new StringBuilder("");
				sbMsgOutrosStatus = new StringBuilder("");
				for (int i = 0; i < listaReturnAnalysis.Count; i++)
				{
					if (isAFStatusAprovado(listaReturnAnalysis[i].Status))
					{
						if (sbMsgAprovados.Length > 0) sbMsgAprovados.Append(", ");
						sbMsgAprovados.Append(listaReturnAnalysis[i].ID + " (" + listaReturnAnalysis[i].Status + ")");
					}
					else if (isAFStatusReprovado(listaReturnAnalysis[i].Status))
					{
						if (sbMsgReprovados.Length > 0) sbMsgReprovados.Append(", ");
						sbMsgReprovados.Append(listaReturnAnalysis[i].ID + " (" + listaReturnAnalysis[i].Status + ")");
					}
					else
					{
						if (sbMsgOutrosStatus.Length > 0) sbMsgOutrosStatus.Append(", ");
						sbMsgOutrosStatus.Append(listaReturnAnalysis[i].ID + " (" + listaReturnAnalysis[i].Status + ")");
					}
				}

				strMsg = "Método " + Global.Cte.Clearsale.Transacao.GetReturnAnalysis.GetMethodName() + " retornou " + listaReturnAnalysis.Count.ToString() + " pedidos:" +
						" Aprovados = " + (sbMsgAprovados.Length == 0 ? "(nenhum)" : sbMsgAprovados.ToString()) + "; " +
						" Reprovados = " + (sbMsgReprovados.Length == 0 ? "(nenhum)" : sbMsgReprovados.ToString()) + "; " +
						" Outros Status = " + (sbMsgOutrosStatus.Length == 0 ? "(nenhum)" : sbMsgOutrosStatus.ToString());

				if (sbMsgInformativa.Length > 0) sbMsgInformativa.Append("; ");
				sbMsgInformativa.Append(strMsg);

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				svcLogMsg = new FinSvcLog();
				svcLogMsg.operacao = NOME_DESTA_ROTINA;
				svcLogMsg.descricao = strMsg;
				GeralDAO.gravaFinSvcLog(svcLogMsg, out msg_erro_aux);
				#endregion

				#endregion

				#region [ Há dados de resposta? ]
				if (listaReturnAnalysis.Count == 0)
				{
					if (sbMsgInformativa.Length > 0) sbMsgInformativa.Append("; ");
					sbMsgInformativa.Append("Não há pedidos para processar");
					return true;
				}
				#endregion

				#region [ Laço para cada pedido retornado pela Clearsale ]
				for (int iCS = 0; iCS < listaReturnAnalysis.Count; iCS++)
				{
					#region [ Serviço deve parar? ]
					if (FinanceiroService.isOnShutdownAcionado)
					{
						if (sbMsgInformativa.Length > 0) sbMsgInformativa.Append("; ");
						sbMsgInformativa.Append("Rotina interrompida devido a shutdown do serviço");
						return true;
					}
					if (FinanceiroService.isOnStopAcionado)
					{
						if (sbMsgInformativa.Length > 0) sbMsgInformativa.Append("; ");
						sbMsgInformativa.Append("Rotina interrompida devido a parada do serviço");
						return true;
					}
					#endregion

					#region [ Status finalizado? ]
					if ((!isAFStatusAprovado(listaReturnAnalysis[iCS].Status)) && (!isAFStatusReprovado(listaReturnAnalysis[iCS].Status)))
					{
						continue;
					}
					#endregion

					qtdePedidosResultadoProcessado++;
					blnErroFinalizacaoBraspag = false;
					sbErroFinalizacaoBraspag = new StringBuilder("");

					#region [ Atualiza o status no BD ]
					if (!ClearsaleDAO.updateRegistroAFGetReturnAnalysisResponse(listaReturnAnalysis[iCS], entityCode, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar atualizar os dados do pedido " + listaReturnAnalysis[iCS].ID + " devido à resposta obtida pela requisição GetReturnAnalysis()!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Clearsale: Falha ao tentar atualizar os dados do pedido " + listaReturnAnalysis[iCS].ID + " devido à resposta obtida pela requisição GetReturnAnalysis() [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar os dados do pedido " + listaReturnAnalysis[iCS].ID + " devido à resposta obtida pela requisição GetReturnAnalysis()\r\n" + msg_erro_aux;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}
					}
					#endregion

					#region [ Status de erro? ]
					if (listaReturnAnalysis[iCS].Status.Equals(Global.Cte.Clearsale.StatusAF.ERRO.GetValue()))
					{
						// Monta relação de pedidos com status de erro p/ envio de email de alerta
						if (sbAlertaStatusErro.Length > 0) sbAlertaStatusErro.Append(", ");
						sbAlertaStatusErro.Append(listaReturnAnalysis[iCS].ID);
						// Segue para o próximo pedido
						continue;
					}
					#endregion

					#region [ Obtém as transações do Pagador vinculadas a esta análise antifraude ]
					// Importante: processar somente as transações de pagamento que ainda estejam como não processadas no nosso sistema.
					// =========== Isso é importante para o caso do processamento ter sido feito anteriormente e haver ocorrido erro
					// durante a requisição SetOrderAsReturned() (ou SetOrderListAsReturned()). Nessa situação, o processamento de
					// finalização teria sido corretamente realizado na Braspag e no BD, mas não na Clearsale. Com isso, o pedido
					// continuaria sendo informado em consultas posteriores do GetReturnAnalysis().
					clearsaleAF = ClearsaleDAO.getClearsaleAFByOrderID(listaReturnAnalysis[iCS].ID, entityCode, out msg_erro_aux);
					if (clearsaleAF == null)
					{
						Global.gravaLogAtividade("Falha ao tentar recuperar os dados do pedido Clearsale " + listaReturnAnalysis[iCS].ID + "\n" + msg_erro_aux);
						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Clearsale: Falha ao tentar recuperar os dados do pedido Clearsale " + listaReturnAnalysis[iCS].ID + "\n" + msg_erro_aux + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\nFalha ao tentar recuperar os dados do pedido Clearsale " + listaReturnAnalysis[iCS].ID + "\n" + msg_erro_aux;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}
						// PROSSEGUE PARA O PRÓXIMO PEDIDO
						continue;
					}

					#region [ Verifica se este pedido já foi tratado, mas ocorreu erro ao tentar atualizar na Clearsale através do método 'SetOrderAsReturned' ]
					if ((clearsaleAF.SetOrderAsReturned_pendente_status == 1) && (clearsaleAF.SetOrderAsReturned_sucesso_status == 0))
					{
						#region [ Assinala o pedido para não retornar mais no retorno de 'GetReturnAnalysis' ]
						if (!executaSetOrderAsReturned(clearsaleAF, out msg_erro_aux))
						{
							msg_erro_aux = "Falha na execução da rotina Clearsale.executaSetOrderAsReturned()\r\n" + msg_erro_aux;
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = msg_erro_aux;
							svcLog.complemento_1 = Global.serializaObjectToXml(clearsaleAF);
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						}
						#endregion

						#region [ Segue para o próximo pedido ]
						continue;
						#endregion
					}
					#endregion

					#region [ Verifica se este pedido já havia sido finalizado anteriormente ]
					// Prevenção para o possível caso em que um pedido finalizado anteriormente volte a aparecer no resultado do GetReturnAnalysis
					// Ex: Pedido foi aprovado e depois houve um chargeback. O status desse pedido é alterado para um dos seguintes status:
					//		FDL = Fraude Deliberada
					//		ATF = Auto Fraude
					//		FAM = Fraude Amigável
					//		DCC = Desacordo Comercial
					// A informação da Clearsale é de que o pedido não deve voltar a aparecer no resultado do GetReturnAnalysis, sendo necessário
					// consultar o status atualizado através do GetOrderStatus
					// Portanto, esta verificação é apenas preventiva.
					if (clearsaleAF.SetOrderAsReturned_sucesso_status == 1)
					{
						#region [ Assinala o pedido novamente para não retornar mais no retorno de 'GetReturnAnalysis' ]
						if (!executaSetOrderAsReturned(clearsaleAF, out msg_erro_aux))
						{
							msg_erro_aux = "Falha na execução da rotina Clearsale.executaSetOrderAsReturned()\r\n" + msg_erro_aux;
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = msg_erro_aux;
							svcLog.complemento_1 = Global.serializaObjectToXml(clearsaleAF);
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						}
						#endregion

						#region [ Grava log e envia mensagem de alerta ]
						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						msg_erro_aux = "Pedido " + clearsaleAF.pedido + " (" + clearsaleAF.pedido_com_sufixo_nsu + ") voltou a ser informado com o status '" + listaReturnAnalysis[iCS].Status + "' no resultado da requisição GetReturnAnalysis após já ter sido finalizado anteriormente";
						svcLogMsg = new FinSvcLog();
						svcLogMsg.operacao = NOME_DESTA_ROTINA;
						svcLogMsg.descricao = msg_erro_aux;
						svcLogMsg.complemento_1 = Global.serializaObjectToXml(clearsaleAF);
						GeralDAO.gravaFinSvcLog(svcLogMsg, out msg_erro_aux);
						#endregion

						#region [ Envia email de alerta sobre situação não prevista ]
						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Pedido " + clearsaleAF.pedido + " (" + clearsaleAF.pedido_com_sufixo_nsu + ") voltou a ser informado com o status '" + listaReturnAnalysis[iCS].Status + "' no resultado da requisição GetReturnAnalysis após já ter sido finalizado anteriormente [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\r\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Pedido " + clearsaleAF.pedido + " (" + clearsaleAF.pedido_com_sufixo_nsu + ") voltou a ser informado com o status '" + listaReturnAnalysis[iCS].Status + "' no resultado da requisição GetReturnAnalysis após já ter sido finalizado anteriormente";
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}
						#endregion
						#endregion

						#region [ Segue para o próximo pedido ]
						continue;
						#endregion
					}
					#endregion

					sbLogNothingToDo = new StringBuilder("");
					listaBraspagPaymentFinalizacao = new List<BraspagPagPaymentFinalizacao>();
					for (int i = 0; i < clearsaleAF.Payments.Count; i++)
					{
						blnAssinalarPaymentAFFinalizado = false;
						braspagPayment = BraspagDAO.getBraspagPagPaymentById(clearsaleAF.Payments[i].id_pagto_gw_pag_payment, out msg_erro_aux);
						if (braspagPayment != null)
						{
							// Verifica se devido a alguma situação inesperada o processamento de antifraude para este pagamento já foi realizado anteriormente,
							// ou seja, o tratamento já foi feito antes integral ou parcialmente e o pedido continua sendo retornado na lista da Clearsale.
							// IMPORTANTE: caso a transação já tenha registrado pagamento no pedido devido a algum acionamento manual ou mesmo automático,
							// o processamento deve prosseguir para que o devido tratamento de finalização seja realizado. Por exemplo, caso a transação tenha
							// sido capturada manualmente e o resultado da análise antifraude foi de reprovação, o sistema deve fazer o cancelamento/estorno da
							// transação de modo automático. Essa situação pode acontecer c/ certa frequência se estiver ativada a rotina que realiza automaticamente
							// a captura de transações pendentes devido ao prazo final de cancelamento automático.
							if (braspagPayment.st_processamento_AF_finalizado == 1)
							{
								sbLogNothingToDo.AppendLine("t_PAGTO_GW_PAG_PAYMENT.id=" + braspagPayment.id.ToString() + " se encontra com st_processamento_AF_finalizado = 1");
								// Transação já foi processada anteriormente, pula e segue para a próxima
								continue;
							}

							if (isAFStatusAprovado(listaReturnAnalysis[iCS].Status))
							{
								// Caso o status seja de aprovação, é necessário fazer a captura da transação, caso ainda não esteja capturada
								if (braspagPayment.ult_GlobalStatus.Equals(Global.Cte.Braspag.Pagador.GlobalStatus.AUTORIZADA.GetValue()))
								{
									listaBraspagPaymentFinalizacao.Add(new BraspagPagPaymentFinalizacao(Global.Cte.Braspag.Pagador.OperacaoFinalizacao.CAPTURA, braspagPayment));
								}
								else
								{
									blnAssinalarPaymentAFFinalizado = true;
								}
							}
							else if (isAFStatusReprovado(listaReturnAnalysis[iCS].Status))
							{
								if ((braspagPayment.ult_GlobalStatus.Equals(Global.Cte.Braspag.Pagador.GlobalStatus.AUTORIZADA.GetValue())) ||
									(braspagPayment.ult_GlobalStatus.Equals(Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue())))
								{
									listaBraspagPaymentFinalizacao.Add(new BraspagPagPaymentFinalizacao(Global.Cte.Braspag.Pagador.OperacaoFinalizacao.CANCELAMENTO, braspagPayment));
								}
								else
								{
									blnAssinalarPaymentAFFinalizado = true;
								}
							}
							else
							{
								blnAssinalarPaymentAFFinalizado = true;
								sbLogNothingToDo.AppendLine("t_PAGTO_GW_PAG_PAYMENT.id=" + braspagPayment.id.ToString() + " se encontra com ult_GlobalStatus = '" + braspagPayment.ult_GlobalStatus + "'");
							}

							if (blnAssinalarPaymentAFFinalizado)
							{
								// A transação está em uma situação que não necessita de finalização, pois pode ter sido tratada manualmente (captura ou cancelamento/estorno manual)
								// Assinala o campo t_PAGTO_GW_PAG_PAYMENT.st_processamento_AF_finalizado = 1 para indicar que a transação foi analisada, mesmo que nenhum processamento
								// de finalização tenha sido necessário.
								#region [ Atualiza campo t_PAGTO_GW_PAG_PAYMENT.st_processamento_AF_finalizado ]
								if (!BraspagDAO.updatePagPaymentAFFinalizado(braspagPayment.id, out msg_erro_aux))
								{
									msg_erro_temp = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									msg_erro_aux = "Falha ao tentar atualizar o campo de status de finalização do processamento AF em t_PAGTO_GW_PAG_PAYMENT.id=" + braspagPayment.id.ToString() +
													" (pedido=" + clearsaleAF.pedido + ", pedido_com_sufixo_nsu=" + clearsaleAF.pedido_com_sufixo_nsu + ")" +
													" em transação que não necessita de finalização (Status AF=" + listaReturnAnalysis[iCS].Status + ", Status PAG=" + braspagPayment.ult_GlobalStatus + " - " + Global.Cte.Braspag.Pagador.GlobalStatus.GetDescription(braspagPayment.ult_GlobalStatus) + ")" +
													"\r\n" + msg_erro_temp;
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = msg_erro_aux;
									svcLog.complemento_1 = Global.serializaObjectToXml(clearsaleAF);
									svcLog.complemento_2 = Global.serializaObjectToXml(braspagPayment);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									#region [ Envia email de alerta sobre situação não prevista ]
									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar atualizar o campo de status de finalização do processamento AF em t_PAGTO_GW_PAG_PAYMENT.id=" + braspagPayment.id.ToString() + " (pedido=" + clearsaleAF.pedido + ", pedido_com_sufixo_nsu=" + clearsaleAF.pedido_com_sufixo_nsu + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar atualizar o campo de status de finalização do processamento AF em t_PAGTO_GW_PAG_PAYMENT.id=" + braspagPayment.id.ToString() +
												" (pedido=" + clearsaleAF.pedido + ", pedido_com_sufixo_nsu=" + clearsaleAF.pedido_com_sufixo_nsu + ")" +
												" em transação que não necessita de finalização (Status AF=" + listaReturnAnalysis[iCS].Status + ", Status PAG=" + braspagPayment.ult_GlobalStatus + " - " + Global.Cte.Braspag.Pagador.GlobalStatus.GetDescription(braspagPayment.ult_GlobalStatus) + ")" +
												"\r\n" + msg_erro_temp;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
									#endregion
								}
								#endregion
							}
						}
					}
					#endregion

					#region [ Se não houver registros de pagamento para tratar ]
					if (listaBraspagPaymentFinalizacao.Count == 0)
					{
						// Apenas registra mensagem informativa no log
						// IMPORTANTE: o processamento deve prosseguir, pois, ao final, o pedido deve ser marcado como 'já tratado' na Clearsale, consultar e gravar no bloco de notas os comentários do analista
						// e registrar o resultado do AF no bloco de notas do pedido.
						strMsg = "O pedido " + listaReturnAnalysis[iCS].ID + " durante o processamento da chamada ao " + Global.Cte.Clearsale.Transacao.GetReturnAnalysis.GetMethodName() + " não possuia nenhuma transação de pagamento com status que demandasse tratamento de finalização com a Braspag." +
								"\r\nResultado do AF = " + listaReturnAnalysis[iCS].Status + ", Score = " + listaReturnAnalysis[iCS].Score +
								"\r\nTransações de pagamento em situação que não demanda finalização com a Braspag:" +
								"\r\n" + sbLogNothingToDo.ToString();

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						svcLogMsg = new FinSvcLog();
						svcLogMsg.operacao = NOME_DESTA_ROTINA;
						svcLogMsg.descricao = strMsg;
						GeralDAO.gravaFinSvcLog(svcLogMsg, out msg_erro_aux);
						#endregion
					}
					#endregion

					#region [ Laço para cada transação do Pagador vinculada a esta análise antifraude ]
					for (int iBPF = 0; iBPF < listaBraspagPaymentFinalizacao.Count; iBPF++)
					{
						rowFinalizacao = listaBraspagPaymentFinalizacao[iBPF];

						#region [ Obtém status atualizado da transação na Braspag ]
						if (!Braspag.processaConsultaGetTransactionData(rowFinalizacao.payment, out ult_GlobalStatus_original, out ult_GlobalStatus_novo, out msg_erro_aux))
						{
							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = msg_erro_aux;
							svcLog.complemento_1 = Global.serializaObjectToXml(rowFinalizacao);
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion
						}
						#endregion

						#region [ Obtém dados atualizados do registro de 'payment' ]
						braspagPaymentAtualizado = BraspagDAO.getBraspagPagPaymentById(rowFinalizacao.payment.id, out msg_erro_aux);
						braspagPag = BraspagDAO.getBraspagPagById(rowFinalizacao.payment.id_pagto_gw_pag, out msg_erro_aux);
						#endregion

						blnAssinalarPaymentAFFinalizado = false;

						if (rowFinalizacao.operacao.GetValue().Equals(Global.Cte.Braspag.Pagador.OperacaoFinalizacao.CAPTURA.GetValue()))
						{
							if (braspagPaymentAtualizado.ult_GlobalStatus.Equals(Global.Cte.Braspag.Pagador.GlobalStatus.AUTORIZADA.GetValue()))
							{
								#region [ Captura o pagamento ]
								if (Braspag.processaRequisicaoCaptureCreditCardTransaction(braspagPaymentAtualizado, out msg_erro_aux))
								{
									blnAssinalarPaymentAFFinalizado = true;
								}
								else
								{
									msg_erro_temp = msg_erro_aux;

									#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
									msg_erro_aux = "Falha ao tentar confirmar a captura da transação do pedido " + braspagPag.pedido + " (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + braspagPaymentAtualizado.id.ToString() + ")\r\n" + msg_erro_temp;
									FinSvcLog svcLog = new FinSvcLog();
									svcLog.operacao = NOME_DESTA_ROTINA;
									svcLog.descricao = msg_erro_aux;
									svcLog.complemento_1 = Global.serializaObjectToXml(braspagPaymentAtualizado);
									GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
									#endregion

									#region [ Envia email de alerta sobre situação não prevista ]
									strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar confirmar a captura da transação do pedido " + braspagPag.pedido + " (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + braspagPaymentAtualizado.id.ToString() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
									strBody = "Mensagem de Financeiro Service\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar confirmar a captura da transação do pedido " + braspagPag.pedido + " (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + braspagPaymentAtualizado.id.ToString() + ")\r\n" + msg_erro_temp;
									if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
									{
										strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
										Global.gravaLogAtividade(strMsg);
									}
									#endregion

									#region [ Ativa flag de erro na finalização na Braspag p/ não marcar o pedido como tratado na Clearsale ]
									blnErroFinalizacaoBraspag = true;
									#endregion

									#region [ Mensagem de erro para mensagem no final ]
									if (sbErroFinalizacaoBraspag.Length > 0) sbErroFinalizacaoBraspag.AppendLine("\r\n\r\n");
									sbErroFinalizacaoBraspag.AppendLine("Erro na finalização com a Braspag (captura): pedido=" + braspagPag.pedido + " (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + braspagPaymentAtualizado.id.ToString() + ")");
									sbErroFinalizacaoBraspag.AppendLine(msg_erro_temp);
									#endregion
								}
								#endregion
							}
							else if (braspagPaymentAtualizado.ult_GlobalStatus.Equals(Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue()))
							{
								#region [ Tratamento para caso excepcional: a transação já estava capturada ]
								// A captura já pode ter sido realizada devido a:
								//	1) Captura manual realizada pelo usuário
								//	2) Captura automática realizada pelo sistema no último dia antes de terminar o prazo do cancelamento automático, caso esta funcionalidade esteja ativa
								#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
								msg_erro_aux = "Processamento de finalização: não foi possível confirmar a captura porque a transação encontra-se no status " + braspagPaymentAtualizado.ult_GlobalStatus + " - " + Global.Cte.Braspag.Pagador.GlobalStatus.GetDescription(braspagPaymentAtualizado.ult_GlobalStatus);
								FinSvcLog svcLog = new FinSvcLog();
								svcLog.operacao = NOME_DESTA_ROTINA;
								svcLog.descricao = msg_erro_aux;
								svcLog.complemento_1 = Global.serializaObjectToXml(rowFinalizacao);
								GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
								#endregion
								#endregion
							}
							else if (braspagPaymentAtualizado.ult_GlobalStatus.Equals(Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURA_CANCELADA.GetValue())
									||
									braspagPaymentAtualizado.ult_GlobalStatus.Equals(Global.Cte.Braspag.Pagador.GlobalStatus.ESTORNADA.GetValue()))
							{
								#region [ Tratamento para caso excepcional: a transação está cancelada ]
								// A transação já pode ter sido cancelada/estornada devido a:
								//	1) Usuário realizou a operação de Void/Refund manualmente devido a alguma situação excepcional
								#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
								msg_erro_aux = "Processamento de finalização: não foi possível confirmar a captura porque a transação encontra-se no status " + braspagPaymentAtualizado.ult_GlobalStatus + " - " + Global.Cte.Braspag.Pagador.GlobalStatus.GetDescription(braspagPaymentAtualizado.ult_GlobalStatus);
								FinSvcLog svcLog = new FinSvcLog();
								svcLog.operacao = NOME_DESTA_ROTINA;
								svcLog.descricao = msg_erro_aux;
								svcLog.complemento_1 = Global.serializaObjectToXml(rowFinalizacao);
								GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
								#endregion
								#endregion
							}
							else
							{
								#region [ Tratamento para situação não prevista ]

								#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
								msg_erro_aux = "Transação encontra-se em situação não prevista (" + braspagPaymentAtualizado.ult_GlobalStatus + ") durante o processamento da finalização (" + rowFinalizacao.operacao.GetValue() + ")";
								FinSvcLog svcLog = new FinSvcLog();
								svcLog.operacao = NOME_DESTA_ROTINA;
								svcLog.descricao = msg_erro_aux;
								svcLog.complemento_1 = Global.serializaObjectToXml(rowFinalizacao);
								GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
								#endregion

								#region [ Envia email de alerta sobre situação não prevista ]
								strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Transação encontra-se em situação não prevista (" + braspagPaymentAtualizado.ult_GlobalStatus + ") durante o processamento da finalização (" + rowFinalizacao.operacao.GetValue() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
								strBody = "Mensagem de Financeiro Service\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Transação encontra-se em situação não prevista (" + braspagPaymentAtualizado.ult_GlobalStatus + ") durante o processamento da finalização (" + rowFinalizacao.operacao.GetValue() + ")";
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
								{
									strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
									Global.gravaLogAtividade(strMsg);
								}
								#endregion

								#endregion
							}
						}
						else if (rowFinalizacao.operacao.GetValue().Equals(Global.Cte.Braspag.Pagador.OperacaoFinalizacao.CANCELAMENTO.GetValue()))
						{
							#region [ Cancela o pagamento ]
							trxFinalizacaoBraspag = null;
							if (braspagPaymentAtualizado.ult_GlobalStatus.Equals(Global.Cte.Braspag.Pagador.GlobalStatus.AUTORIZADA.GetValue()))
							{
								trxFinalizacaoBraspag = Global.Cte.Braspag.Pagador.Transacao.VoidCreditCardTransaction;
							}
							else if (braspagPaymentAtualizado.ult_GlobalStatus.Equals(Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue()))
							{
								if (braspagPaymentAtualizado.resp_CapturedDate.Date == DateTime.Now.Date)
								{
									trxFinalizacaoBraspag = Global.Cte.Braspag.Pagador.Transacao.VoidCreditCardTransaction;
								}
								else
								{
									trxFinalizacaoBraspag = Global.Cte.Braspag.Pagador.Transacao.RefundCreditCardTransaction;
								}
							}
							else
							{
								#region [ Tratamento para situação não prevista ]

								#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
								msg_erro_aux = "Transação encontra-se em situação não prevista (" + braspagPaymentAtualizado.ult_GlobalStatus + ") durante o processamento da finalização com cancelamento (" + rowFinalizacao.operacao.GetValue() + ")";
								FinSvcLog svcLog = new FinSvcLog();
								svcLog.operacao = NOME_DESTA_ROTINA;
								svcLog.descricao = msg_erro_aux;
								svcLog.complemento_1 = Global.serializaObjectToXml(rowFinalizacao);
								svcLog.complemento_2 = Global.serializaObjectToXml(braspagPaymentAtualizado);
								GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
								#endregion

								#region [ Envia email de alerta sobre situação não prevista ]
								strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Transação encontra-se em situação não prevista (" + braspagPaymentAtualizado.ult_GlobalStatus + ") durante o processamento da finalização com cancelamento (" + rowFinalizacao.operacao.GetValue() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
								strBody = "Mensagem de Financeiro Service\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Transação encontra-se em situação não prevista (" + braspagPaymentAtualizado.ult_GlobalStatus + ") durante o processamento da finalização com cancelamento (" + rowFinalizacao.operacao.GetValue() + ")";
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
								{
									strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
									Global.gravaLogAtividade(strMsg);
								}
								#endregion

								#endregion
							}

							if (trxFinalizacaoBraspag != null)
							{
								if (trxFinalizacaoBraspag.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.VoidCreditCardTransaction.GetCodOpLog()))
								{
									#region [ Cancela a transação ]
									if (Braspag.processaRequisicaoVoidCreditCardTransaction(braspagPaymentAtualizado, out msg_erro_aux))
									{
										blnAssinalarPaymentAFFinalizado = true;
									}
									else
									{
										msg_erro_temp = msg_erro_aux;

										#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
										msg_erro_aux = "Falha ao tentar cancelar (void) a transação do pedido " + braspagPag.pedido + " (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + braspagPaymentAtualizado.id.ToString() + ")\r\n" + msg_erro_temp;
										FinSvcLog svcLog = new FinSvcLog();
										svcLog.operacao = NOME_DESTA_ROTINA;
										svcLog.descricao = msg_erro_aux;
										svcLog.complemento_1 = Global.serializaObjectToXml(braspagPaymentAtualizado);
										GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
										#endregion

										#region [ Envia email de alerta sobre situação não prevista ]
										strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar cancelar (void) a transação do pedido " + braspagPag.pedido + " (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + braspagPaymentAtualizado.id.ToString() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
										strBody = "Mensagem de Financeiro Service\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar cancelar (void) a transação do pedido " + braspagPag.pedido + " (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + braspagPaymentAtualizado.id.ToString() + ")\r\n" + msg_erro_temp;
										if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
										{
											strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
											Global.gravaLogAtividade(strMsg);
										}
										#endregion

										#region [ Ativa flag de erro na finalização na Braspag p/ não marcar o pedido como tratado na Clearsale ]
										blnErroFinalizacaoBraspag = true;
										#endregion

										#region [ Mensagem de erro para mensagem no final ]
										if (sbErroFinalizacaoBraspag.Length > 0) sbErroFinalizacaoBraspag.AppendLine("\r\n\r\n");
										sbErroFinalizacaoBraspag.AppendLine("Erro na finalização com a Braspag (void): pedido=" + braspagPag.pedido + " (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + braspagPaymentAtualizado.id.ToString() + ")");
										sbErroFinalizacaoBraspag.AppendLine(msg_erro_temp);
										#endregion
									}
									#endregion
								}
								else if (trxFinalizacaoBraspag.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.RefundCreditCardTransaction.GetCodOpLog()))
								{
									#region [ Estorna a transação ]
									if (Braspag.processaRequisicaoRefundCreditCardTransaction(braspagPaymentAtualizado, out msg_erro_aux))
									{
										blnAssinalarPaymentAFFinalizado = true;
									}
									else
									{
										msg_erro_temp = msg_erro_aux;

										#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
										msg_erro_aux = "Falha ao tentar estornar (refund) a transação do pedido " + braspagPag.pedido + " (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + braspagPaymentAtualizado.id.ToString() + ")\r\n" + msg_erro_temp;
										FinSvcLog svcLog = new FinSvcLog();
										svcLog.operacao = NOME_DESTA_ROTINA;
										svcLog.descricao = msg_erro_aux;
										svcLog.complemento_1 = Global.serializaObjectToXml(braspagPaymentAtualizado);
										GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
										#endregion

										#region [ Envia email de alerta sobre situação não prevista ]
										strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar estornar (refund) a transação do pedido " + braspagPag.pedido + " (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + braspagPaymentAtualizado.id.ToString() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
										strBody = "Mensagem de Financeiro Service\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar estornar (refund) a transação do pedido " + braspagPag.pedido + " (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + braspagPaymentAtualizado.id.ToString() + ")\r\n" + msg_erro_temp;
										if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
										{
											strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
											Global.gravaLogAtividade(strMsg);
										}
										#endregion

										#region [ Ativa flag de erro na finalização na Braspag p/ não marcar o pedido como tratado na Clearsale ]
										blnErroFinalizacaoBraspag = true;
										#endregion

										#region [ Mensagem de erro para mensagem no final ]
										if (sbErroFinalizacaoBraspag.Length > 0) sbErroFinalizacaoBraspag.AppendLine("\r\n\r\n");
										sbErroFinalizacaoBraspag.AppendLine("Erro na finalização com a Braspag (refund): pedido=" + braspagPag.pedido + " (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + braspagPaymentAtualizado.id.ToString() + ")");
										sbErroFinalizacaoBraspag.AppendLine(msg_erro_temp);
										#endregion
									}
									#endregion
								}
							}
							#endregion
						}
						else
						{
							#region [ Tratamento para situação não prevista ]

							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							msg_erro_aux = "Operação de finalização desconhecida: " + rowFinalizacao.operacao.GetValue();
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = msg_erro_aux;
							svcLog.complemento_1 = Global.serializaObjectToXml(rowFinalizacao);
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							#region [ Envia email de alerta sobre situação não prevista ]
							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Operação de finalização desconhecida (" + rowFinalizacao.operacao.GetValue() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Operação de finalização desconhecida (" + rowFinalizacao.operacao.GetValue() + ")";
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}
							#endregion

							#endregion
						}

						if (blnAssinalarPaymentAFFinalizado)
						{
							#region [ Atualiza campo t_PAGTO_GW_PAG_PAYMENT.st_processamento_AF_finalizado ]
							if (!BraspagDAO.updatePagPaymentAFFinalizado(braspagPaymentAtualizado.id, out msg_erro_aux))
							{
								msg_erro_temp = msg_erro_aux;

								#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
								msg_erro_aux = "Falha ao tentar atualizar o campo de status de finalização do processamento AF em t_PAGTO_GW_PAG_PAYMENT.id=" + braspagPaymentAtualizado.id.ToString() + " (pedido=" + clearsaleAF.pedido + ", pedido_com_sufixo_nsu=" + clearsaleAF.pedido_com_sufixo_nsu + ")\r\n" + msg_erro_temp;
								FinSvcLog svcLog = new FinSvcLog();
								svcLog.operacao = NOME_DESTA_ROTINA;
								svcLog.descricao = msg_erro_aux;
								svcLog.complemento_1 = Global.serializaObjectToXml(clearsaleAF);
								svcLog.complemento_2 = Global.serializaObjectToXml(braspagPaymentAtualizado);
								GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
								#endregion

								#region [ Envia email de alerta sobre situação não prevista ]
								strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar atualizar o campo de status de finalização do processamento AF em t_PAGTO_GW_PAG_PAYMENT.id=" + braspagPaymentAtualizado.id.ToString() + " (pedido=" + clearsaleAF.pedido + ", pedido_com_sufixo_nsu=" + clearsaleAF.pedido_com_sufixo_nsu + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
								strBody = "Mensagem de Financeiro Service\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar atualizar o campo de status de finalização do processamento AF em t_PAGTO_GW_PAG_PAYMENT.id=" + braspagPaymentAtualizado.id.ToString() + " (pedido=" + clearsaleAF.pedido + ", pedido_com_sufixo_nsu=" + clearsaleAF.pedido_com_sufixo_nsu + ")\r\n" + msg_erro_temp;
								if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
								{
									strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
									Global.gravaLogAtividade(strMsg);
								}
								#endregion
							}
							#endregion
						}
					}  //for (int iBPF = 0; iBPF < listaBraspagPaymentFinalizacao.Count; iBPF++)
					#endregion

					#region [ Grava no bloco de notas do pedido os comentários do analista ]
					entityCode = Global.Cte.Clearsale.CS_ENTITY_CODE;
					xmlReqSoap = montaRequisicaoGetAnalystComments(entityCode, listaReturnAnalysis[iCS].ID);
					if (!enviaRequisicaoComRetry(xmlReqSoap, Global.Cte.Clearsale.Transacao.GetAnalystComments, out xmlRespSoap, out msg_erro_aux))
					{
						msg_erro_temp = msg_erro_aux;
						msg_erro_aux = "Falha ao tentar enviar transação para a Clearsale: " + Global.Cte.Clearsale.Transacao.GetAnalystComments.GetMethodName() + "!!\n" + msg_erro_temp;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						svcLogMsg = new FinSvcLog();
						svcLogMsg.operacao = NOME_DESTA_ROTINA;
						svcLogMsg.descricao = msg_erro_aux;
						svcLogMsg.complemento_1 = xmlReqSoap;
						GeralDAO.gravaFinSvcLog(svcLogMsg, out msg_erro_aux);
						#endregion

						#region [ Envia email de alerta sobre situação não prevista ]
						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha na requisição Clearsale " + Global.Cte.Clearsale.Transacao.GetAnalystComments.GetMethodName() + " para o pedido " + listaReturnAnalysis[iCS].ID + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha na requisição Clearsale " + Global.Cte.Clearsale.Transacao.GetAnalystComments.GetMethodName() + " para o pedido " + listaReturnAnalysis[iCS].ID + "\r\n" + msg_erro_temp;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}
						#endregion
					}
					else
					{
						listaComments = obtemAnalystCommentsFromGetAnalystCommentsResponse(xmlRespSoap, out msg_erro_aux);
						if (listaComments != null)
						{
							listaCommentsOrderByCreateDate = listaComments.OrderBy(o => o.DataHoraCreateDate).ToList();
							foreach (var comentario in listaCommentsOrderByCreateDate)
							{
								strBlocoNotas = comentario.Comments;
								if ((strBlocoNotas ?? "").Trim().Length > 0)
								{
									if (strBlocoNotas.Contains("\n") && (!strBlocoNotas.Contains("\r\n"))) strBlocoNotas = strBlocoNotas.Replace("\n", "\r\n");
									if (!PedidoDAO.gravaPedidoBlocoNotasComDataHora(comentario.DataHoraCreateDate, clearsaleAF.pedido, Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA, Global.Cte.BlocoNotasPedidoNivelAcesso.COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, strBlocoNotas, out msg_erro_aux))
									{
										msg_erro_temp = msg_erro_aux;

										#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
										msg_erro_aux = "Falha ao tentar gravar bloco de notas no pedido " + clearsaleAF.pedido + " com os comentários do analista da Clearsale\r\n" + msg_erro_temp;
										FinSvcLog svcLog = new FinSvcLog();
										svcLog.operacao = NOME_DESTA_ROTINA;
										svcLog.descricao = msg_erro_aux;
										svcLog.complemento_1 = strBlocoNotas;
										GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
										#endregion

										#region [ Envia email de alerta sobre situação não prevista ]
										strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar gravar bloco de notas no pedido " + clearsaleAF.pedido + " com os comentários do analista da Clearsale [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
										strBody = "Mensagem de Financeiro Service\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar gravar bloco de notas no pedido " + clearsaleAF.pedido + " com os comentários do analista da Clearsale\r\n" + msg_erro_temp;
										if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
										{
											strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
											Global.gravaLogAtividade(strMsg);
										}
										#endregion
									}
								}
							}
						}
					}
					#endregion

					#region [ Grava um bloco de notas no pedido informando resultado do AF ]
					strBlocoNotas = "Resultado Clearsale: " + listaReturnAnalysis[iCS].Status + " - " + Global.Cte.Clearsale.StatusAF.GetDescription(listaReturnAnalysis[iCS].Status) + " (Score: " + listaReturnAnalysis[iCS].Score + ")";
					if (!PedidoDAO.gravaPedidoBlocoNotas(clearsaleAF.pedido, Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA, Global.Cte.BlocoNotasPedidoNivelAcesso.COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, strBlocoNotas, out msg_erro_aux))
					{
						msg_erro_temp = msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						msg_erro_aux = "Falha ao tentar gravar bloco de notas no pedido " + clearsaleAF.pedido + " com informações do resultado da análise antifraude Clearsale\r\n" + msg_erro_temp;
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro_aux;
						svcLog.complemento_1 = strBlocoNotas;
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						#region [ Envia email de alerta sobre situação não prevista ]
						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar gravar bloco de notas no pedido " + clearsaleAF.pedido + " com informações do resultado da análise antifraude Clearsale [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar gravar bloco de notas no pedido " + clearsaleAF.pedido + " com informações do resultado da análise antifraude Clearsale\r\n" + msg_erro_temp;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}
						#endregion
					}
					#endregion

					#region [ Marca o pedido como tratado na Clearsale ]
					// Importante: mesmo que tenha ocorrido algum erro na finalização com a Braspag em alguma transação, marca o pedido como já tratado na Clearsale
					// para que o pedido não retorne mais na resposta da requisição GetReturnAnalysis para evitar um possível 'loop' infinito.
					blnResultadoSetOrderAsReturned = executaSetOrderAsReturned(clearsaleAF, out msg_erro_aux);
					if (!blnResultadoSetOrderAsReturned)
					{
						msg_erro_temp = msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						msg_erro_aux = "Falha na requisição Clearsale (" + Global.Cte.Clearsale.Transacao.SetOrderAsReturned.GetMethodName() + ") para o pedido " + clearsaleAF.pedido + " (pedido_com_sufixo_nsu = " + clearsaleAF.pedido_com_sufixo_nsu + ", status = " + listaReturnAnalysis[iCS].Status + ")\r\n" + msg_erro_temp;
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro_aux;
						svcLog.complemento_1 = Global.serializaObjectToXml(clearsaleAF);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						#region [ Envia email de alerta sobre situação não prevista ]
						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha na requisição Clearsale (" + Global.Cte.Clearsale.Transacao.SetOrderAsReturned.GetMethodName() + ") para o pedido " + clearsaleAF.pedido + " (pedido_com_sufixo_nsu = " + clearsaleAF.pedido_com_sufixo_nsu + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha na requisição Clearsale (" + Global.Cte.Clearsale.Transacao.SetOrderAsReturned.GetMethodName() + ") para o pedido " + clearsaleAF.pedido + " (pedido_com_sufixo_nsu = " + clearsaleAF.pedido_com_sufixo_nsu + ", status = " + listaReturnAnalysis[iCS].Status + ")\r\n" + msg_erro_temp;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}
						#endregion
					}
					#endregion

					#region [ Grava log e envia mensagem de alerta ]
					if (blnErroFinalizacaoBraspag)
					{
						if (blnResultadoSetOrderAsReturned)
						{
							msg_erro_temp = "O pedido " + clearsaleAF.pedido + " (" + clearsaleAF.pedido_com_sufixo_nsu + ") foi marcado como já tratado na Clearsale com sucesso (status=" + listaReturnAnalysis[iCS].Status + "), mas houve erro na finalização da transação com a Braspag!" +
											"\r\n\r\n" +
											sbErroFinalizacaoBraspag.ToString();

							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = msg_erro_temp;
							svcLog.complemento_1 = Global.serializaObjectToXml(clearsaleAF);
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							#region [ Envia email de alerta ]
							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": o pedido " + clearsaleAF.pedido + " (" + clearsaleAF.pedido_com_sufixo_nsu + ") foi marcado como já tratado na Clearsale com sucesso, mas houve erro na finalização da transação com a Braspag [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + "\r\n" + msg_erro_temp;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}
							#endregion
						}
					}
					#endregion
				} // for (int iCS = 0; iCS < listaReturnAnalysis.Count; iCS++)

				#region [ Se algum pedido retornou com status de erro, envia email de alerta ]
				if (sbAlertaStatusErro.Length > 0)
				{
					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Clearsale: Pedidos que retornaram com status '" + Global.Cte.Clearsale.StatusAF.ERRO.GetValue() + "' [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\nPedidos que retornaram da Clearsale com status '" + Global.Cte.Clearsale.StatusAF.ERRO.GetValue() + "': " + sbAlertaStatusErro.ToString();
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
					}

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = "Pedidos que retornaram da Clearsale com status '" + Global.Cte.Clearsale.StatusAF.ERRO.GetValue() + "': " + sbAlertaStatusErro.ToString();
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion
				}
				#endregion

				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
			finally
			{
				strMsgInformativa = sbMsgInformativa.ToString();
			}
		}
		#endregion

		#region [ obtemAnalystCommentsFromGetAnalystCommentsResponse ]
		private static List<ClearsaleAnalystComments> obtemAnalystCommentsFromGetAnalystCommentsResponse(string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			string strValue;
			XmlDocument xmlDoc;
			XmlDocument xmlDocGetAnalystCommentsResult;
			XmlNamespaceManager nsmgr;
			XmlNode xmlNode;
			XmlNodeList xmlNodeList;
			ClearsaleAnalystComments comments;
			List<ClearsaleAnalystComments> listaComments = new List<ClearsaleAnalystComments>();
			#endregion

			msg_erro = "";

			try
			{
				if ((xmlRespSoap ?? "").Trim().Length == 0) return null;

				xmlDoc = new XmlDocument();
				xmlDoc.LoadXml(xmlRespSoap);
				nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
				nsmgr.AddNamespace("cs", "http://www.clearsale.com.br/integration");
				xmlNode = xmlDoc.SelectSingleNode("//cs:GetAnalystCommentsResult", nsmgr);
				if (xmlNode.ChildNodes.Count > 0)
				{
					strValue = HttpUtility.HtmlDecode(xmlNode.FirstChild.Value);
					xmlDocGetAnalystCommentsResult = new XmlDocument();
					xmlDocGetAnalystCommentsResult.LoadXml(strValue);
					xmlNodeList = xmlDocGetAnalystCommentsResult.SelectNodes("//AnalystComments/AnalystComments");
					foreach (XmlNode node in xmlNodeList)
					{
						comments = new ClearsaleAnalystComments(Global.obtemXmlChildNodeValue(node, "CreateDate"), Global.obtemXmlChildNodeValue(node, "Comments"), Global.obtemXmlChildNodeValue(node, "UserName"), Global.obtemXmlChildNodeValue(node, "Status"), Global.obtemXmlChildNodeValue(node, "LineName"));
						listaComments.Add(comments);
					}
				}

				return listaComments;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				return null;
			}
		}
		#endregion

		#region [ executaSetOrderAsReturned ]
		private static bool executaSetOrderAsReturned(ClearsaleAF clearsaleAF, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Clearsale.executaSetOrderAsReturned()";
			Global.Cte.Clearsale.Transacao trxSelecionada = Global.Cte.Clearsale.Transacao.SetOrderAsReturned;
			int id_emailsndsvc_mensagem;
			bool blnEnviouOk;
			string msg_erro_requisicao;
			string msg_erro_aux;
			string msg_erro_temp;
			string strSubject;
			string strBody;
			string strMsg;
			string xmlReqSoap;
			string xmlRespSoap;
			string strValue;
			string strTransactionStatus_StatusCode = "";
			XmlDocument xmlDoc;
			XmlDocument xmlDocSetOrderAsReturnedResult;
			XmlNode xmlNode;
			XmlNamespaceManager nsmgr;
			ClearsaleAFOpComplementar opComplIns;
			ClearsaleAFOpComplementar opComplUpd;
			ClearsaleAFOpComplementarXml opComplXmlTx;
			ClearsaleAFOpComplementarXml opComplXmlRx;
			FinSvcLog svcLog;
			#endregion

			msg_erro = "";
			try
			{
				if (clearsaleAF == null)
				{
					msg_erro = "Parâmetro com dados da análise antifraude não foi fornecido!";
					return false;
				}

				xmlReqSoap = montaRequisicaoSoapSetOrderAsReturned(clearsaleAF.req_entityCode, clearsaleAF.req_Order_ID);

				#region [ Grava requisição em t_PAGTO_GW_AF_OP_COMPLEMENTAR ]
				opComplIns = new ClearsaleAFOpComplementar();
				opComplIns.id_pagto_gw_af = clearsaleAF.id;
				opComplIns.usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				opComplIns.operacao = trxSelecionada.GetCodOpLog();
				if (!ClearsaleDAO.insereAFOpComplementar(opComplIns, out msg_erro_aux))
				{
					msg_erro_aux = NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_AF_OP_COMPLEMENTAR + "\n" + msg_erro_aux;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro_aux;
					svcLog.complemento_1 = Global.serializaObjectToXml(opComplIns);
					svcLog.complemento_2 = Global.serializaObjectToXml(clearsaleAF);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion
				}
				#endregion

				#region [ Grava requisição em T_PAGTO_GW_AF_OP_COMPLEMENTAR_XML (TX) ]
				opComplXmlTx = new ClearsaleAFOpComplementarXml();
				opComplXmlTx.id_pagto_gw_af_op_complementar = opComplIns.id;
				opComplXmlTx.tipo_transacao = trxSelecionada.GetCodOpLog();
				opComplXmlTx.fluxo_xml = Global.Cte.FluxoXml.TX.GetValue();
				opComplXmlTx.xml = xmlReqSoap;
				if (!ClearsaleDAO.insereAFOpComplementarXml(opComplXmlTx, out msg_erro_aux))
				{
					msg_erro_aux = NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_AF_OP_COMPLEMENTAR_XML + "\n" + msg_erro_aux;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro_aux;
					svcLog.complemento_1 = Global.serializaObjectToXml(opComplXmlTx);
					svcLog.complemento_2 = Global.serializaObjectToXml(clearsaleAF);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion
				}
				#endregion

				#region [ Envia requisição para a Clearsale ]
				blnEnviouOk = enviaRequisicaoComRetry(xmlReqSoap, trxSelecionada, out xmlRespSoap, out msg_erro_requisicao);
				#endregion

				#region [ Grava resposta da requisição em T_PAGTO_GW_AF_OP_COMPLEMENTAR_XML (RX) ]
				opComplXmlRx = new ClearsaleAFOpComplementarXml();
				opComplXmlRx.id_pagto_gw_af_op_complementar = opComplIns.id;
				opComplXmlRx.tipo_transacao = trxSelecionada.GetCodOpLog();
				opComplXmlRx.fluxo_xml = Global.Cte.FluxoXml.RX.GetValue();
				opComplXmlRx.xml = (xmlRespSoap == null ? "" : xmlRespSoap);
				if (!ClearsaleDAO.insereAFOpComplementarXml(opComplXmlRx, out msg_erro_aux))
				{
					msg_erro_aux = NOME_DESTA_ROTINA + ": Falha ao tentar inserir registro na tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_AF_OP_COMPLEMENTAR_XML + "\n" + msg_erro_aux;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro_aux;
					svcLog.complemento_1 = Global.serializaObjectToXml(opComplXmlRx);
					svcLog.complemento_2 = Global.serializaObjectToXml(clearsaleAF);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion
				}
				#endregion

				#region [ Falha no envio? ]
				if (!blnEnviouOk)
				{
					msg_erro_aux = "Falha ao tentar enviar transação para a Clearsale: " + trxSelecionada.GetMethodName() + "!!\n" + msg_erro_requisicao;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro_aux;
					svcLog.complemento_1 = xmlReqSoap;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					#region [ Grava na lista de pedidos pendentes para a transação 'SetOrderAsReturned' ]
					if (!ClearsaleDAO.updateRegistroAFSetOrderAsReturnedPendente(clearsaleAF.id, out msg_erro_aux))
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						msg_erro_aux = "Falha ao tentar registrar a falha na requisição SetOrderAsReturned no registro de AF do pedido " + clearsaleAF.pedido + " (pedido_com_sufixo_nsu=" + clearsaleAF.pedido_com_sufixo_nsu + ")\r\n" + msg_erro;
						svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro_aux;
						svcLog.complemento_1 = Global.serializaObjectToXml(clearsaleAF);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						#region [ Envia email de alerta sobre situação não prevista ]
						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar registrar a falha na requisição SetOrderAsReturned no registro de AF do pedido " + clearsaleAF.pedido + " (pedido_com_sufixo_nsu=" + clearsaleAF.pedido_com_sufixo_nsu + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar a falha na requisição SetOrderAsReturned no registro de AF do pedido " + clearsaleAF.pedido + " (pedido_com_sufixo_nsu=" + clearsaleAF.pedido_com_sufixo_nsu + ")\r\n" + msg_erro;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}
						#endregion

						return false;
					}
					#endregion
				}
				#endregion

				#region [ Resposta nula? ]
				if (xmlRespSoap == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Requisição ao método Clearsale " + trxSelecionada.GetMethodName() + " retornou resposta nula";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = xmlReqSoap;
					svcLog.complemento_2 = Global.serializaObjectToXml(clearsaleAF);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				#region [ Decodifica XML de resposta ]
				opComplUpd = new ClearsaleAFOpComplementar();
				opComplUpd.id = opComplIns.id;
				opComplUpd.id_pagto_gw_af = opComplIns.id_pagto_gw_af;
				opComplUpd.trx_RX_vazio_status = 1;
				if (xmlRespSoap.Trim().Length > 0)
				{
					opComplUpd.trx_RX_vazio_status = 0;
					opComplUpd.trx_RX_status = 1;
				}

				if (xmlRespSoap.Trim().Length > 0)
				{
					xmlDoc = new XmlDocument();
					xmlDoc.LoadXml(xmlRespSoap);
					nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
					nsmgr.AddNamespace("cs", "http://www.clearsale.com.br/integration");
					xmlNode = xmlDoc.SelectSingleNode("//cs:SetOrderAsReturnedResult", nsmgr);
					if (xmlNode.ChildNodes.Count > 0)
					{
						strValue = HttpUtility.HtmlDecode(xmlNode.FirstChild.Value);
						xmlDocSetOrderAsReturnedResult = new XmlDocument();
						xmlDocSetOrderAsReturnedResult.LoadXml(strValue);
						xmlNode = xmlDocSetOrderAsReturnedResult.SelectSingleNode("//TransactionStatus/StatusCode");
						strTransactionStatus_StatusCode = (xmlNode != null ? xmlNode.FirstChild.Value : "");
						if (strTransactionStatus_StatusCode.Equals(Global.Cte.Clearsale.TransactionStatus_StatusCode.OK.GetValue()))
						{
							opComplUpd.st_sucesso = 1;
						}
					}
				}
				#endregion

				#region [ Atualiza dados em t_PAGTO_GW_AF_OP_COMPLEMENTAR ]
				if (!ClearsaleDAO.updateAFOpComplementar(opComplUpd, out msg_erro_aux))
				{
					msg_erro_aux = NOME_DESTA_ROTINA + ": Falha ao tentar atualizar o registro na tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_AF_OP_COMPLEMENTAR + "\n" + msg_erro_aux;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro_aux;
					svcLog.complemento_1 = Global.serializaObjectToXml(opComplUpd);
					svcLog.complemento_2 = Global.serializaObjectToXml(clearsaleAF);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion
				}
				#endregion

				if (strTransactionStatus_StatusCode.Equals(Global.Cte.Clearsale.TransactionStatus_StatusCode.OK.GetValue()))
				{
					if (!ClearsaleDAO.updateRegistroAFSetOrderAsReturnedSucesso(clearsaleAF.id, out msg_erro_aux))
					{
						msg_erro_temp = msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						msg_erro_aux = "Falha ao tentar registrar o sucesso da requisição SetOrderAsReturned no registro de AF do pedido " + clearsaleAF.pedido + " (pedido_com_sufixo_nsu=" + clearsaleAF.pedido_com_sufixo_nsu + ")\r\n" + msg_erro_temp;
						svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro_aux;
						svcLog.complemento_1 = Global.serializaObjectToXml(clearsaleAF);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						#region [ Envia email de alerta sobre situação não prevista ]
						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar registrar o sucesso da requisição SetOrderAsReturned no registro de AF do pedido " + clearsaleAF.pedido + " (pedido_com_sufixo_nsu=" + clearsaleAF.pedido_com_sufixo_nsu + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\n" + Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": Falha ao tentar registrar o sucesso da requisição SetOrderAsReturned no registro de AF do pedido " + clearsaleAF.pedido + " (pedido_com_sufixo_nsu=" + clearsaleAF.pedido_com_sufixo_nsu + ")\r\n" + msg_erro_temp;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}
						#endregion
					}
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ calculaValorPagador ]
		private static bool calculaValorPagador(string numeroPedido, out decimal vl_pagador, out string msg_erro)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			vl_pagador = 0m;
			msg_erro = "";

			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				strSql = "SELECT" +
							" Sum(valor_transacao) AS valor_total_transacao" +
						" FROM t_PAGTO_GW_PAG t_PAG" +
							" INNER JOIN t_PAGTO_GW_PAG_PAYMENT t_PAYMENT ON (t_PAG.id = t_PAYMENT.id_pagto_gw_pag)" +
						" WHERE" +
							" (pedido = '" + numeroPedido + "')" +
							" AND (ult_GlobalStatus IN (" +
										"'" + Global.Cte.Braspag.Pagador.GlobalStatus.AUTORIZADA.GetValue() + "'," +
										"'" + Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue() + "'))";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count > 0)
				{
					vl_pagador = BD.readToDecimal(dtbResultado.Rows[0][0]);
				}

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return false;
			}
		}
		#endregion

		#region [ obtemDataHoraTrxPagadorMaisAntiga ]
		private static bool obtemDataHoraTrxPagadorMaisAntiga(string numeroPedido, out DateTime dtHrTrxPagador, out string msg_erro)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			dtHrTrxPagador = DateTime.MinValue;
			msg_erro = "";

			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				strSql = "SELECT" +
							" Min(data_hora) AS data_hora" +
						" FROM t_PAGTO_GW_PAG t_PAG" +
							" INNER JOIN t_PAGTO_GW_PAG_PAYMENT t_PAYMENT ON (t_PAG.id = t_PAYMENT.id_pagto_gw_pag)" +
						" WHERE" +
							" (pedido = '" + numeroPedido + "')" +
							" AND (st_enviado_analise_AF = 0)" +
							" AND (st_cancelado_envio_analise_AF = 0)" +
							" AND (ult_GlobalStatus IN (" +
										"'" + Global.Cte.Braspag.Pagador.GlobalStatus.AUTORIZADA.GetValue() + "'," +
										"'" + Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue() + "'))";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count > 0)
				{
					dtHrTrxPagador = BD.readToDateTime(dtbResultado.Rows[0][0]);
				}

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.Message;
				return false;
			}
		}
		#endregion

		#region [ xmlMontaCampo ]
		private static string xmlMontaCampo(string conteudo, string tagName)
		{
			return xmlMontaCampo(conteudo, tagName, 0);
		}

		private static string xmlMontaCampo(string conteudo, string tagName, int qtdeTabs)
		{
			string strResp;
			strResp = new String('\t', qtdeTabs) + "<" + tagName + ">" + conteudo + "</" + tagName + ">";
			return strResp;
		}
		#endregion

		#region [ montaXmlPaymentAF ]
		private static string montaXmlPaymentAF(int sequencia, DataRow rowPagPayment, out ClearsaleAFPayment afPayment)
		{
			#region [ Declarações ]
			StringBuilder sbXml = new StringBuilder();
			StringBuilder sbXmlAddress = new StringBuilder();
			string strValue;
			#endregion

			afPayment = new ClearsaleAFPayment();

			afPayment.id_pagto_gw_pag_payment = BD.readToInt(rowPagPayment["id"]);
			afPayment.bandeira = BD.readToString(rowPagPayment["bandeira"]);
			afPayment.valor_transacao = BD.readToDecimal(rowPagPayment["valor_transacao"]);

			strValue = sequencia.ToString();
			sbXml.Append(xmlMontaCampo(strValue, "Sequential"));
			afPayment.af_Sequential = strValue;

			strValue = Global.formataDataHoraYyyyMmDdTHhMmSs(BD.readToDateTime(rowPagPayment["data_hora"]));
			sbXml.Append(xmlMontaCampo(strValue, "Date"));
			afPayment.af_Date = strValue;

			strValue = Global.formataMoedaClearsale(BD.readToDecimal(rowPagPayment["valor_transacao"]));
			sbXml.Append(xmlMontaCampo(strValue, "Amount"));
			afPayment.af_Amount = strValue;

			strValue = "1";
			sbXml.Append(xmlMontaCampo(strValue, "PaymentTypeID"));  // 1=Cartão de Crédito
			afPayment.af_PaymentTypeID = strValue;

			strValue = BD.readToString(rowPagPayment["req_PaymentDataRequest_NumberOfPayments"]);
			sbXml.Append(xmlMontaCampo(strValue, "QtyInstallments"));
			afPayment.af_QtyInstallments = strValue;

			// Envia o número do cartão somente se ele não estiver mascarado p/ proteger os dados
			strValue = BD.readToString(rowPagPayment["req_PaymentDataRequest_CardNumber"]);
			if (!strValue.Contains("*")) sbXml.Append(xmlMontaCampo(strValue, "CardNumber"));
			afPayment.af_CardNumber = strValue;

			strValue = Texto.leftStr(BD.readToString(rowPagPayment["req_PaymentDataRequest_CardNumber"]), 6);
			sbXml.Append(xmlMontaCampo(strValue, "CardBin"));
			afPayment.af_CardBin = strValue;

			strValue = Texto.rightStr(BD.readToString(rowPagPayment["req_PaymentDataRequest_CardNumber"]), 4);
			sbXml.Append(xmlMontaCampo(strValue, "CardEndNumber"));
			afPayment.af_CardEndNumber = strValue;

			strValue = bandeiraCodificaParaPadraoClearsale(BD.readToString(rowPagPayment["bandeira"]));
			sbXml.Append(xmlMontaCampo(strValue, "CardType"));
			afPayment.af_CardType = strValue;

			strValue = BD.readToString(rowPagPayment["checkout_cartao_validade_mes"]) + "/" + BD.readToString(rowPagPayment["checkout_cartao_validade_ano"]);
			sbXml.Append(xmlMontaCampo(strValue, "CardExpirationDate"));
			afPayment.af_CardExpirationDate = strValue;

			strValue = Global.filtraAmpersand(BD.readToString(rowPagPayment["checkout_titular_nome"]));
			sbXml.Append(xmlMontaCampo(strValue, "Name"));
			afPayment.af_Name = strValue;

			strValue = BD.readToString(rowPagPayment["checkout_titular_cpf_cnpj"]);
			sbXml.Append(xmlMontaCampo(strValue, "LegalDocument"));
			afPayment.af_LegalDocument = strValue;

			// Address
			strValue = Global.filtraAmpersand(BD.readToString(rowPagPayment["checkout_fatura_end_logradouro"]));
			sbXmlAddress.Append(xmlMontaCampo(strValue, "Street"));
			afPayment.af_Address_Street = strValue;

			strValue = Global.filtraAmpersand(BD.readToString(rowPagPayment["checkout_fatura_end_numero"]));
			sbXmlAddress.Append(xmlMontaCampo(strValue, "Number"));
			afPayment.af_Address_Number = strValue;

			strValue = Global.filtraAmpersand(BD.readToString(rowPagPayment["checkout_fatura_end_complemento"]));
			if (strValue.Length > 0)
			{
				sbXmlAddress.Append(xmlMontaCampo(strValue, "Comp"));
				afPayment.af_Address_Comp = strValue;
			}

			strValue = Global.filtraAmpersand(BD.readToString(rowPagPayment["checkout_fatura_end_bairro"]));
			sbXmlAddress.Append(xmlMontaCampo(strValue, "County"));
			afPayment.af_Address_County = strValue;

			strValue = Global.filtraAmpersand(BD.readToString(rowPagPayment["checkout_fatura_end_cidade"]));
			sbXmlAddress.Append(xmlMontaCampo(strValue, "City"));
			afPayment.af_Address_City = strValue;

			strValue = BD.readToString(rowPagPayment["checkout_fatura_end_uf"]);
			sbXmlAddress.Append(xmlMontaCampo(strValue, "State"));
			afPayment.af_Address_State = strValue;

			strValue = BD.readToString(rowPagPayment["checkout_fatura_end_cep"]);
			sbXmlAddress.Append(xmlMontaCampo(strValue, "ZipCode"));
			afPayment.af_Address_ZipCode = strValue;

			sbXml.Append("<Address>" + sbXmlAddress.ToString() + "</Address>");

			strValue = "986";
			sbXml.Append(xmlMontaCampo(strValue, "Currency"));  // Brazilian Real: BRL / 986
			afPayment.af_Currency = strValue;

			return "<Payment>" + sbXml.ToString() + "</Payment>";
		}
		#endregion

		#region [ montaXmlItem ]
		private static string montaXmlItem(PedidoItem item, out ClearsaleAFItem afItem)
		{
			#region [ Declarações ]
			StringBuilder sbXml = new StringBuilder();
			string strValue;
			#endregion

			afItem = new ClearsaleAFItem();

			strValue = item.produto;
			sbXml.Append(xmlMontaCampo(strValue, "ID"));
			afItem.af_ID = strValue;

			strValue = Global.filtraAmpersand(item.descricao);
			sbXml.Append(xmlMontaCampo(strValue, "Name"));
			afItem.af_Name = strValue;

			strValue = Global.formataMoedaClearsale(item.preco_NF);
			sbXml.Append(xmlMontaCampo(strValue, "ItemValue"));
			afItem.af_ItemValue = strValue;

			strValue = item.qtde.ToString();
			sbXml.Append(xmlMontaCampo(strValue, "Qty"));
			afItem.af_Qty = strValue;

			return "<Item>" + sbXml.ToString() + "</Item>";
		}
		#endregion

		#region [ montaRequisicaoSoapSendOrders ]
		private static string montaRequisicaoSoapSendOrders(string entityCode, string xmlRequisicao)
		{
			string xmlRequisicaoSoap;

			xmlRequisicaoSoap = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
								"<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:int=\"http://www.clearsale.com.br/integration\">" +
								"<soapenv:Header/>" +
								"<soapenv:Body>" +
								"<int:SendOrders>" +
								"<int:entityCode>" + entityCode + "</int:entityCode>" +
								"<int:xml><![CDATA[" + Global.filtraAcentuacao(xmlRequisicao) + "]]></int:xml>" +
								"</int:SendOrders>" +
								"</soapenv:Body>" +
								"</soapenv:Envelope>";
			return xmlRequisicaoSoap;
		}
		#endregion

		#region [ montaRequisicaoSoapGetReturnAnalysis ]
		private static string montaRequisicaoSoapGetReturnAnalysis(string entityCode)
		{
			string xmlSoap;

			xmlSoap = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
					"<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:int=\"http://www.clearsale.com.br/integration\">" +
					"<soapenv:Header/>" +
					"<soapenv:Body>" +
					"<int:GetReturnAnalysis>" +
					"<int:entityCode>" + entityCode + "</int:entityCode>" +
					"</int:GetReturnAnalysis>" +
					"</soapenv:Body>" +
					"</soapenv:Envelope>";

			return xmlSoap;
		}
		#endregion

		#region [ montaRequisicaoSoapSetOrderAsReturned ]
		private static string montaRequisicaoSoapSetOrderAsReturned(string entityCode, string orderID)
		{
			string xmlSoap;

			xmlSoap = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
					"<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:int=\"http://www.clearsale.com.br/integration\">" +
					"<soapenv:Header/>" +
					"<soapenv:Body>" +
					"<int:SetOrderAsReturned>" +
					"<int:entityCode>" + entityCode + "</int:entityCode>" +
					"<int:orderID>" + orderID + "</int:orderID>" +
					"</int:SetOrderAsReturned>" +
					"</soapenv:Body>" +
					"</soapenv:Envelope>";

			return xmlSoap;
		}
		#endregion

		#region [ montaRequisicaoGetAnalystComments ]
		private static string montaRequisicaoGetAnalystComments(string entityCode, string orderID)
		{
			string xmlRequisicaoSoap;

			xmlRequisicaoSoap = "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
								"<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:int=\"http://www.clearsale.com.br/integration\">" +
								"<soapenv:Header/>" +
								"<soapenv:Body>" +
								"<int:GetAnalystComments>" +
								"<int:entityCode>" + entityCode + "</int:entityCode>" +
								"<int:orderID>" + orderID + "</int:orderID>" +
								"<int:getAll>true</int:getAll>" +
								"</int:GetAnalystComments>" +
								"</soapenv:Body>" +
								"</soapenv:Envelope>";

			return xmlRequisicaoSoap;
		}
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
		/// <param name="xmlRespSoap"></param>
		/// <param name="msg_erro"></param>
		/// <returns></returns>
		private static bool enviaRequisicaoComRetry(string xmlReqSoap, Global.Cte.Clearsale.Transacao trxParam, out string xmlRespSoap, out string msg_erro)
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

				Thread.Sleep(5*1000);
			} while (qtdeTentativasRealizadas < MAX_TENTATIVAS);

			return blnResposta;
		}
		#endregion

		#region [ enviaRequisicao ]
		private static bool enviaRequisicao(string xmlReqSoap, Global.Cte.Clearsale.Transacao trxParam, out string xmlRespSoap, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Clearsale.enviaRequisicao()";
			HttpWebRequest req;
			HttpWebResponse resp;
			#endregion

			xmlRespSoap = "";
			msg_erro = "";

			try
			{
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + trxParam.GetMethodName() + " - XML (TX)\n" + xmlReqSoap);

				req = (HttpWebRequest)WebRequest.Create(trxParam.GetEnderecoWebService());
				// The Timeout applies to the entire request and response, not individually to the GetRequestStream and GetResponse method calls
				req.Timeout = Global.Cte.Clearsale.REQUEST_TIMEOUT_EM_MS;
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
	}

	#region [ ClearsaleAF ]
	public class ClearsaleAF
	{
		#region [ Construtor ]
		public ClearsaleAF()
		{
			Order_BillingData_Phones = new List<ClearsaleAFPhone>();
			Order_ShippingData_Phones = new List<ClearsaleAFPhone>();
			Payments = new List<ClearsaleAFPayment>();
			Items = new List<ClearsaleAFItem>();
		}
		#endregion

		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private DateTime _data;
		public DateTime data
		{
			get { return _data; }
			set { _data = value; }
		}

		private DateTime _data_hora;
		public DateTime data_hora
		{
			get { return _data_hora; }
			set { _data_hora = value; }
		}

		private string _usuario;
		public string usuario
		{
			get { return _usuario; }
			set { _usuario = value; }
		}

		private int _owner;
		public int owner
		{
			get { return _owner; }
			set { _owner = value; }
		}

		private string _loja;
		public string loja
		{
			get { return _loja; }
			set { _loja = value; }
		}

		private string _id_cliente;
		public string id_cliente
		{
			get { return _id_cliente; }
			set { _id_cliente = value; }
		}

		private string _pedido;
		public string pedido
		{
			get { return _pedido; }
			set { _pedido = value; }
		}

		private string _pedido_com_sufixo_nsu;
		public string pedido_com_sufixo_nsu
		{
			get { return _pedido_com_sufixo_nsu; }
			set { _pedido_com_sufixo_nsu = value; }
		}

		private decimal _valor_pedido;
		public decimal valor_pedido
		{
			get { return _valor_pedido; }
			set { _valor_pedido = value; }
		}

		private DateTime _trx_TX_data;
		public DateTime trx_TX_data
		{
			get { return _trx_TX_data; }
			set { _trx_TX_data = value; }
		}

		private DateTime _trx_TX_data_hora;
		public DateTime trx_TX_data_hora
		{
			get { return _trx_TX_data_hora; }
			set { _trx_TX_data_hora = value; }
		}

		private short _trx_RX_status;
		public short trx_RX_status
		{
			get { return _trx_RX_status; }
			set { _trx_RX_status = value; }
		}

		private DateTime _trx_RX_data;
		public DateTime trx_RX_data
		{
			get { return _trx_RX_data; }
			set { _trx_RX_data = value; }
		}

		private DateTime _trx_RX_data_hora;
		public DateTime trx_RX_data_hora
		{
			get { return _trx_RX_data_hora; }
			set { _trx_RX_data_hora = value; }
		}

		private short _trx_RX_vazio_status;
		public short trx_RX_vazio_status
		{
			get { return _trx_RX_vazio_status; }
			set { _trx_RX_vazio_status = value; }
		}

		private short _trx_erro_status;
		public short trx_erro_status
		{
			get { return _trx_erro_status; }
			set { _trx_erro_status = value; }
		}

		private string _trx_erro_codigo;
		public string trx_erro_codigo
		{
			get { return _trx_erro_codigo; }
			set { _trx_erro_codigo = value; }
		}

		private string _trx_erro_mensagem;
		public string trx_erro_mensagem
		{
			get { return _trx_erro_mensagem; }
			set { _trx_erro_mensagem = value; }
		}

		private int _trx_TX_id_pagto_gw_af_xml;
		public int trx_TX_id_pagto_gw_af_xml
		{
			get { return _trx_TX_id_pagto_gw_af_xml; }
			set { _trx_TX_id_pagto_gw_af_xml = value; }
		}

		private int _trx_RX_id_pagto_gw_af_xml;
		public int trx_RX_id_pagto_gw_af_xml
		{
			get { return _trx_RX_id_pagto_gw_af_xml; }
			set { _trx_RX_id_pagto_gw_af_xml = value; }
		}

		private string _req_entityCode = "";
		public string req_entityCode
		{
			get { return _req_entityCode; }
			set { _req_entityCode = value; }
		}

		private string _req_Order_ID = "";
		public string req_Order_ID
		{
			get { return _req_Order_ID; }
			set { _req_Order_ID = value; }
		}

		private string _req_Order_FingerPrint_SessionID = "";
		public string req_Order_FingerPrint_SessionID
		{
			get { return _req_Order_FingerPrint_SessionID; }
			set { _req_Order_FingerPrint_SessionID = value; }
		}

		private string _req_Order_Date = "";
		public string req_Order_Date
		{
			get { return _req_Order_Date; }
			set { _req_Order_Date = value; }
		}

		private string _req_Order_Email = "";
		public string req_Order_Email
		{
			get { return _req_Order_Email; }
			set { _req_Order_Email = value; }
		}

		private string _req_Order_B2B_B2C = "";
		public string req_Order_B2B_B2C
		{
			get { return _req_Order_B2B_B2C; }
			set { _req_Order_B2B_B2C = value; }
		}

		private string _req_Order_ShippingPrice = "";
		public string req_Order_ShippingPrice
		{
			get { return _req_Order_ShippingPrice; }
			set { _req_Order_ShippingPrice = value; }
		}

		private string _req_Order_TotalItems = "";
		public string req_Order_TotalItems
		{
			get { return _req_Order_TotalItems; }
			set { _req_Order_TotalItems = value; }
		}

		private string _req_Order_TotalOrder = "";
		public string req_Order_TotalOrder
		{
			get { return _req_Order_TotalOrder; }
			set { _req_Order_TotalOrder = value; }
		}

		private string _req_Order_QtyInstallments = "";
		public string req_Order_QtyInstallments
		{
			get { return _req_Order_QtyInstallments; }
			set { _req_Order_QtyInstallments = value; }
		}

		private string _req_Order_DeliveryTimeCD = "";
		public string req_Order_DeliveryTimeCD
		{
			get { return _req_Order_DeliveryTimeCD; }
			set { _req_Order_DeliveryTimeCD = value; }
		}

		private string _req_Order_QtyItems = "";
		public string req_Order_QtyItems
		{
			get { return _req_Order_QtyItems; }
			set { _req_Order_QtyItems = value; }
		}

		private string _req_Order_QtyPaymentTypes = "";
		public string req_Order_QtyPaymentTypes
		{
			get { return _req_Order_QtyPaymentTypes; }
			set { _req_Order_QtyPaymentTypes = value; }
		}

		private string _req_Order_IP = "";
		public string req_Order_IP
		{
			get { return _req_Order_IP; }
			set { _req_Order_IP = value; }
		}

		private string _req_Order_Status = "";
		public string req_Order_Status
		{
			get { return _req_Order_Status; }
			set { _req_Order_Status = value; }
		}

		private string _req_Order_Reanalise = "";
		public string req_Order_Reanalise
		{
			get { return _req_Order_Reanalise; }
			set { _req_Order_Reanalise = value; }
		}

		private string _req_Order_Origin = "";
		public string req_Order_Origin
		{
			get { return _req_Order_Origin; }
			set { _req_Order_Origin = value; }
		}

		private string _req_Order_BillingData_ID = "";
		public string req_Order_BillingData_ID
		{
			get { return _req_Order_BillingData_ID; }
			set { _req_Order_BillingData_ID = value; }
		}

		private string _req_Order_BillingData_Type = "";
		public string req_Order_BillingData_Type
		{
			get { return _req_Order_BillingData_Type; }
			set { _req_Order_BillingData_Type = value; }
		}

		private string _req_Order_BillingData_LegalDocument1 = "";
		public string req_Order_BillingData_LegalDocument1
		{
			get { return _req_Order_BillingData_LegalDocument1; }
			set { _req_Order_BillingData_LegalDocument1 = value; }
		}

		private string _req_Order_BillingData_LegalDocument2 = "";
		public string req_Order_BillingData_LegalDocument2
		{
			get { return _req_Order_BillingData_LegalDocument2; }
			set { _req_Order_BillingData_LegalDocument2 = value; }
		}

		private string _req_Order_BillingData_Name = "";
		public string req_Order_BillingData_Name
		{
			get { return _req_Order_BillingData_Name; }
			set { _req_Order_BillingData_Name = value; }
		}

		private string _req_Order_BillingData_BirthDate = "";
		public string req_Order_BillingData_BirthDate
		{
			get { return _req_Order_BillingData_BirthDate; }
			set { _req_Order_BillingData_BirthDate = value; }
		}

		private string _req_Order_BillingData_Email = "";
		public string req_Order_BillingData_Email
		{
			get { return _req_Order_BillingData_Email; }
			set { _req_Order_BillingData_Email = value; }
		}

		private string _req_Order_BillingData_Gender = "";
		public string req_Order_BillingData_Gender
		{
			get { return _req_Order_BillingData_Gender; }
			set { _req_Order_BillingData_Gender = value; }
		}

		private string _req_Order_BillingData_Address_Street = "";
		public string req_Order_BillingData_Address_Street
		{
			get { return _req_Order_BillingData_Address_Street; }
			set { _req_Order_BillingData_Address_Street = value; }
		}

		private string _req_Order_BillingData_Address_Number = "";
		public string req_Order_BillingData_Address_Number
		{
			get { return _req_Order_BillingData_Address_Number; }
			set { _req_Order_BillingData_Address_Number = value; }
		}

		private string _req_Order_BillingData_Address_Comp = "";
		public string req_Order_BillingData_Address_Comp
		{
			get { return _req_Order_BillingData_Address_Comp; }
			set { _req_Order_BillingData_Address_Comp = value; }
		}

		private string _req_Order_BillingData_Address_County = "";
		public string req_Order_BillingData_Address_County
		{
			get { return _req_Order_BillingData_Address_County; }
			set { _req_Order_BillingData_Address_County = value; }
		}

		private string _req_Order_BillingData_Address_City = "";
		public string req_Order_BillingData_Address_City
		{
			get { return _req_Order_BillingData_Address_City; }
			set { _req_Order_BillingData_Address_City = value; }
		}

		private string _req_Order_BillingData_Address_State = "";
		public string req_Order_BillingData_Address_State
		{
			get { return _req_Order_BillingData_Address_State; }
			set { _req_Order_BillingData_Address_State = value; }
		}

		private string _req_Order_BillingData_Address_Country = "";
		public string req_Order_BillingData_Address_Country
		{
			get { return _req_Order_BillingData_Address_Country; }
			set { _req_Order_BillingData_Address_Country = value; }
		}

		private string _req_Order_BillingData_Address_ZipCode = "";
		public string req_Order_BillingData_Address_ZipCode
		{
			get { return _req_Order_BillingData_Address_ZipCode; }
			set { _req_Order_BillingData_Address_ZipCode = value; }
		}

		private string _req_Order_BillingData_Address_Reference = "";
		public string req_Order_BillingData_Address_Reference
		{
			get { return _req_Order_BillingData_Address_Reference; }
			set { _req_Order_BillingData_Address_Reference = value; }
		}

		private string _req_Order_ShippingData_ID = "";
		public string req_Order_ShippingData_ID
		{
			get { return _req_Order_ShippingData_ID; }
			set { _req_Order_ShippingData_ID = value; }
		}

		private string _req_Order_ShippingData_Type = "";
		public string req_Order_ShippingData_Type
		{
			get { return _req_Order_ShippingData_Type; }
			set { _req_Order_ShippingData_Type = value; }
		}

		private string _req_Order_ShippingData_LegalDocument1 = "";
		public string req_Order_ShippingData_LegalDocument1
		{
			get { return _req_Order_ShippingData_LegalDocument1; }
			set { _req_Order_ShippingData_LegalDocument1 = value; }
		}

		private string _req_Order_ShippingData_LegalDocument2 = "";
		public string req_Order_ShippingData_LegalDocument2
		{
			get { return _req_Order_ShippingData_LegalDocument2; }
			set { _req_Order_ShippingData_LegalDocument2 = value; }
		}

		private string _req_Order_ShippingData_Name = "";
		public string req_Order_ShippingData_Name
		{
			get { return _req_Order_ShippingData_Name; }
			set { _req_Order_ShippingData_Name = value; }
		}

		private string _req_Order_ShippingData_BirthDate = "";
		public string req_Order_ShippingData_BirthDate
		{
			get { return _req_Order_ShippingData_BirthDate; }
			set { _req_Order_ShippingData_BirthDate = value; }
		}

		private string _req_Order_ShippingData_Email = "";
		public string req_Order_ShippingData_Email
		{
			get { return _req_Order_ShippingData_Email; }
			set { _req_Order_ShippingData_Email = value; }
		}

		private string _req_Order_ShippingData_Gender = "";
		public string req_Order_ShippingData_Gender
		{
			get { return _req_Order_ShippingData_Gender; }
			set { _req_Order_ShippingData_Gender = value; }
		}

		private string _req_Order_ShippingData_Address_Street = "";
		public string req_Order_ShippingData_Address_Street
		{
			get { return _req_Order_ShippingData_Address_Street; }
			set { _req_Order_ShippingData_Address_Street = value; }
		}

		private string _req_Order_ShippingData_Address_Number = "";
		public string req_Order_ShippingData_Address_Number
		{
			get { return _req_Order_ShippingData_Address_Number; }
			set { _req_Order_ShippingData_Address_Number = value; }
		}

		private string _req_Order_ShippingData_Address_Comp = "";
		public string req_Order_ShippingData_Address_Comp
		{
			get { return _req_Order_ShippingData_Address_Comp; }
			set { _req_Order_ShippingData_Address_Comp = value; }
		}

		private string _req_Order_ShippingData_Address_County = "";
		public string req_Order_ShippingData_Address_County
		{
			get { return _req_Order_ShippingData_Address_County; }
			set { _req_Order_ShippingData_Address_County = value; }
		}

		private string _req_Order_ShippingData_Address_City = "";
		public string req_Order_ShippingData_Address_City
		{
			get { return _req_Order_ShippingData_Address_City; }
			set { _req_Order_ShippingData_Address_City = value; }
		}

		private string _req_Order_ShippingData_Address_State = "";
		public string req_Order_ShippingData_Address_State
		{
			get { return _req_Order_ShippingData_Address_State; }
			set { _req_Order_ShippingData_Address_State = value; }
		}

		private string _req_Order_ShippingData_Address_Country = "";
		public string req_Order_ShippingData_Address_Country
		{
			get { return _req_Order_ShippingData_Address_Country; }
			set { _req_Order_ShippingData_Address_Country = value; }
		}

		private string _req_Order_ShippingData_Address_ZipCode = "";
		public string req_Order_ShippingData_Address_ZipCode
		{
			get { return _req_Order_ShippingData_Address_ZipCode; }
			set { _req_Order_ShippingData_Address_ZipCode = value; }
		}

		private string _req_Order_ShippingData_Address_Reference = "";
		public string req_Order_ShippingData_Address_Reference
		{
			get { return _req_Order_ShippingData_Address_Reference; }
			set { _req_Order_ShippingData_Address_Reference = value; }
		}

		private string _resp_ID;
		public string resp_ID
		{
			get { return _resp_ID; }
			set { _resp_ID = value; }
		}

		private string _resp_Status;
		public string resp_Status
		{
			get { return _resp_Status; }
			set { _resp_Status = value; }
		}

		private string _resp_Score;
		public string resp_Score
		{
			get { return _resp_Score; }
			set { _resp_Score = value; }
		}

		private string _prim_Status = "";
		public string prim_Status
		{
			get { return _prim_Status; }
			set { _prim_Status = value; }
		}

		private DateTime _prim_atualizacao_data_hora;
		public DateTime prim_atualizacao_data_hora
		{
			get { return _prim_atualizacao_data_hora; }
			set { _prim_atualizacao_data_hora = value; }
		}

		private string _ult_Status = "";
		public string ult_Status
		{
			get { return _ult_Status; }
			set { _ult_Status = value; }
		}

		private DateTime _ult_atualizacao_data_hora;
		public DateTime ult_atualizacao_data_hora
		{
			get { return _ult_atualizacao_data_hora; }
			set { _ult_atualizacao_data_hora = value; }
		}

		private short _anulado_status;
		public short anulado_status
		{
			get { return _anulado_status; }
			set { _anulado_status = value; }
		}

		private DateTime _anulado_data;
		public DateTime anulado_data
		{
			get { return _anulado_data; }
			set { _anulado_data = value; }
		}

		private DateTime _anulado_data_hora;
		public DateTime anulado_data_hora
		{
			get { return _anulado_data_hora; }
			set { _anulado_data_hora = value; }
		}

		private int _anulado_por_id_pagto_gw_af;
		public int anulado_por_id_pagto_gw_af
		{
			get { return _anulado_por_id_pagto_gw_af; }
			set { _anulado_por_id_pagto_gw_af = value; }
		}

		private byte _SetOrderAsReturned_pendente_status;
		public byte SetOrderAsReturned_pendente_status
		{
			get { return _SetOrderAsReturned_pendente_status; }
			set { _SetOrderAsReturned_pendente_status = value; }
		}

		private DateTime _SetOrderAsReturned_pendente_data_hora;
		public DateTime SetOrderAsReturned_pendente_data_hora
		{
			get { return _SetOrderAsReturned_pendente_data_hora; }
			set { _SetOrderAsReturned_pendente_data_hora = value; }
		}

		private byte _SetOrderAsReturned_sucesso_status;
		public byte SetOrderAsReturned_sucesso_status
		{
			get { return _SetOrderAsReturned_sucesso_status; }
			set { _SetOrderAsReturned_sucesso_status = value; }
		}

		private DateTime _SetOrderAsReturned_sucesso_data_hora;
		public DateTime SetOrderAsReturned_sucesso_data_hora
		{
			get { return _SetOrderAsReturned_sucesso_data_hora; }
			set { _SetOrderAsReturned_sucesso_data_hora = value; }
		}

		public List<ClearsaleAFPhone> Order_BillingData_Phones;
		public List<ClearsaleAFPhone> Order_ShippingData_Phones;
		public List<ClearsaleAFPayment> Payments;
		public List<ClearsaleAFItem> Items;
	}
	#endregion

	#region [ ClearsaleAFPhone ]
	public class ClearsaleAFPhone
	{
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _id_pagto_gw_af;
		public int id_pagto_gw_af
		{
			get { return _id_pagto_gw_af; }
			set { _id_pagto_gw_af = value; }
		}

		private string _IdBlocoXml;
		public string idBlocoXml
		{
			get { return _IdBlocoXml; }
			set { _IdBlocoXml = value; }
		}

		private string _af_Type = "";
		public string af_Type
		{
			get { return _af_Type; }
			set { _af_Type = value; }
		}

		private string _af_DDI = "";
		public string af_DDI
		{
			get { return _af_DDI; }
			set { _af_DDI = value; }
		}

		private string _af_DDD = "";
		public string af_DDD
		{
			get { return _af_DDD; }
			set { _af_DDD = value; }
		}

		private string _af_Number = "";
		public string af_Number
		{
			get { return _af_Number; }
			set { _af_Number = value; }
		}

		private string _af_Extension = "";
		public string af_Extension
		{
			get { return _af_Extension; }
			set { _af_Extension = value; }
		}
	}
	#endregion

	#region [ ClearsaleAFPayment ]
	public class ClearsaleAFPayment
	{
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _id_pagto_gw_af;
		public int id_pagto_gw_af
		{
			get { return _id_pagto_gw_af; }
			set { _id_pagto_gw_af = value; }
		}

		private int _id_pagto_gw_pag_payment;
		public int id_pagto_gw_pag_payment
		{
			get { return _id_pagto_gw_pag_payment; }
			set { _id_pagto_gw_pag_payment = value; }
		}

		private int _ordem;
		public int ordem
		{
			get { return _ordem; }
			set { _ordem = value; }
		}

		private string _bandeira;
		public string bandeira
		{
			get { return _bandeira; }
			set { _bandeira = value; }
		}

		private decimal _valor_transacao;
		public decimal valor_transacao
		{
			get { return _valor_transacao; }
			set { _valor_transacao = value; }
		}

		private string _af_Sequential = "";
		public string af_Sequential
		{
			get { return _af_Sequential; }
			set { _af_Sequential = value; }
		}

		private string _af_Date = "";
		public string af_Date
		{
			get { return _af_Date; }
			set { _af_Date = value; }
		}

		private string _af_Amount = "";
		public string af_Amount
		{
			get { return _af_Amount; }
			set { _af_Amount = value; }
		}

		private string _af_PaymentTypeID = "";
		public string af_PaymentTypeID
		{
			get { return _af_PaymentTypeID; }
			set { _af_PaymentTypeID = value; }
		}

		private string _af_QtyInstallments = "";
		public string af_QtyInstallments
		{
			get { return _af_QtyInstallments; }
			set { _af_QtyInstallments = value; }
		}

		private string _af_Interest = "";
		public string af_Interest
		{
			get { return _af_Interest; }
			set { _af_Interest = value; }
		}

		private string _af_InterestValue = "";
		public string af_InterestValue
		{
			get { return _af_InterestValue; }
			set { _af_InterestValue = value; }
		}

		private string _af_CardNumber = "";
		public string af_CardNumber
		{
			get { return _af_CardNumber; }
			set { _af_CardNumber = value; }
		}

		private string _af_CardBin = "";
		public string af_CardBin
		{
			get { return _af_CardBin; }
			set { _af_CardBin = value; }
		}

		private string _af_CardEndNumber = "";
		public string af_CardEndNumber
		{
			get { return _af_CardEndNumber; }
			set { _af_CardEndNumber = value; }
		}

		private string _af_CardType = "";
		public string af_CardType
		{
			get { return _af_CardType; }
			set { _af_CardType = value; }
		}

		private string _af_CardExpirationDate = "";
		public string af_CardExpirationDate
		{
			get { return _af_CardExpirationDate; }
			set { _af_CardExpirationDate = value; }
		}

		private string _af_Name = "";
		public string af_Name
		{
			get { return _af_Name; }
			set { _af_Name = value; }
		}

		private string _af_LegalDocument = "";
		public string af_LegalDocument
		{
			get { return _af_LegalDocument; }
			set { _af_LegalDocument = value; }
		}

		private string _af_Address_Street = "";
		public string af_Address_Street
		{
			get { return _af_Address_Street; }
			set { _af_Address_Street = value; }
		}

		private string _af_Address_Number = "";
		public string af_Address_Number
		{
			get { return _af_Address_Number; }
			set { _af_Address_Number = value; }
		}

		private string _af_Address_Comp = "";
		public string af_Address_Comp
		{
			get { return _af_Address_Comp; }
			set { _af_Address_Comp = value; }
		}

		private string _af_Address_County = "";
		public string af_Address_County
		{
			get { return _af_Address_County; }
			set { _af_Address_County = value; }
		}

		private string _af_Address_City = "";
		public string af_Address_City
		{
			get { return _af_Address_City; }
			set { _af_Address_City = value; }
		}

		private string _af_Address_State = "";
		public string af_Address_State
		{
			get { return _af_Address_State; }
			set { _af_Address_State = value; }
		}

		private string _af_Address_Country = "";
		public string af_Address_Country
		{
			get { return _af_Address_Country; }
			set { _af_Address_Country = value; }
		}

		private string _af_Address_ZipCode = "";
		public string af_Address_ZipCode
		{
			get { return _af_Address_ZipCode; }
			set { _af_Address_ZipCode = value; }
		}

		private string _af_Address_Reference = "";
		public string af_Address_Reference
		{
			get { return _af_Address_Reference; }
			set { _af_Address_Reference = value; }
		}

		private string _af_Nsu = "";
		public string af_Nsu
		{
			get { return _af_Nsu; }
			set { _af_Nsu = value; }
		}

		private string _af_Currency = "";
		public string af_Currency
		{
			get { return _af_Currency; }
			set { _af_Currency = value; }
		}
	}
	#endregion

	#region [ ClearsaleAFItem ]
	public class ClearsaleAFItem
	{
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _id_pagto_gw_af;
		public int id_pagto_gw_af
		{
			get { return _id_pagto_gw_af; }
			set { _id_pagto_gw_af = value; }
		}

		private string _af_ID = "";
		public string af_ID
		{
			get { return _af_ID; }
			set { _af_ID = value; }
		}

		private string _af_Name = "";
		public string af_Name
		{
			get { return _af_Name; }
			set { _af_Name = value; }
		}

		private string _af_ItemValue = "";
		public string af_ItemValue
		{
			get { return _af_ItemValue; }
			set { _af_ItemValue = value; }
		}

		private string _af_Qty = "";
		public string af_Qty
		{
			get { return _af_Qty; }
			set { _af_Qty = value; }
		}

		private string _af_CategoryID = "";
		public string af_CategoryID
		{
			get { return _af_CategoryID; }
			set { _af_CategoryID = value; }
		}

		private string _af_CategoryName = "";
		public string af_CategoryName
		{
			get { return _af_CategoryName; }
			set { _af_CategoryName = value; }
		}
	}
	#endregion

	#region [ ClearsaleAFXml ]
	public class ClearsaleAFXml
	{
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _id_pagto_gw_af;
		public int id_pagto_gw_af
		{
			get { return _id_pagto_gw_af; }
			set { _id_pagto_gw_af = value; }
		}

		private DateTime _data;
		public DateTime data
		{
			get { return _data; }
			set { _data = value; }
		}

		private DateTime _data_hora;
		public DateTime data_hora
		{
			get { return _data_hora; }
			set { _data_hora = value; }
		}

		private string _tipo_transacao;
		public string tipo_transacao
		{
			get { return _tipo_transacao; }
			set { _tipo_transacao = value; }
		}

		private string _fluxo_xml;
		public string fluxo_xml
		{
			get { return _fluxo_xml; }
			set { _fluxo_xml = value; }
		}

		private string _xml;
		public string xml
		{
			get { return _xml; }
			set { _xml = value; }
		}
	}
	#endregion

	#region [ ClearsaleSendOrdersResponseOrder ]
	public class ClearsaleSendOrdersResponseOrder
	{
		private string _ID;
		public string ID
		{
			get { return _ID; }
			set { _ID = value; }
		}

		private string _Status;
		public string Status
		{
			get { return _Status; }
			set { _Status = value; }
		}

		private string _Score;
		public string Score
		{
			get { return _Score; }
			set { _Score = value; }
		}

		public ClearsaleSendOrdersResponseOrder(string ID, string Status, string Score)
		{
			this._ID = ID;
			this._Status = Status;
			this._Score = Score;
		}
	}
	#endregion

	#region [ ClearsaleSendOrdersResponse ]
	public class ClearsaleSendOrdersResponse
	{
		private string _TransactionID;
		public string TransactionID
		{
			get { return _TransactionID; }
			set { _TransactionID = value; }
		}

		private string _StatusCode;
		public string StatusCode
		{
			get { return _StatusCode; }
			set { _StatusCode = value; }
		}

		private string _Message;
		public string Message
		{
			get { return _Message; }
			set { _Message = value; }
		}

		public List<ClearsaleSendOrdersResponseOrder> Orders;

		public ClearsaleSendOrdersResponse()
		{
			Orders = new List<ClearsaleSendOrdersResponseOrder>();
		}
	}
	#endregion

	#region [ ClearsaleGetReturnAnalysisResponse ]
	public class ClearsaleGetReturnAnalysisResponse
	{
		private string _ID;
		public string ID
		{
			get { return _ID; }
			set { _ID = value; }
		}

		private string _Status;
		public string Status
		{
			get { return _Status; }
			set { _Status = value; }
		}

		private string _Score;
		public string Score
		{
			get { return _Score; }
			set { _Score = value; }
		}

		public ClearsaleGetReturnAnalysisResponse(string ID, string Status, string Score)
		{
			this._ID = ID;
			this._Status = Status;
			this._Score = Score;
		}
	}
	#endregion

	#region [ ClearsaleSetOrderAsReturnedResponse ]
	public class ClearsaleSetOrderAsReturnedResponse
	{
		private string _StatusCode;
		public string StatusCode
		{
			get { return _StatusCode; }
			set { _StatusCode = value; }
		}

		private string _Message;
		public string Message
		{
			get { return _Message; }
			set { _Message = value; }
		}

		public ClearsaleSetOrderAsReturnedResponse(string StatusCode, string Message)
		{
			this._StatusCode = StatusCode;
			this._Message = Message;
		}
	}
	#endregion

	#region [ ClearsaleAFOpComplementar ]
	public class ClearsaleAFOpComplementar
	{
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _id_pagto_gw_af;
		public int id_pagto_gw_af
		{
			get { return _id_pagto_gw_af; }
			set { _id_pagto_gw_af = value; }
		}

		private DateTime _data;
		public DateTime data
		{
			get { return _data; }
			set { _data = value; }
		}

		private DateTime _data_hora;
		public DateTime data_hora
		{
			get { return _data_hora; }
			set { _data_hora = value; }
		}

		private string _usuario;
		public string usuario
		{
			get { return _usuario; }
			set { _usuario = value; }
		}

		private string _operacao;
		public string operacao
		{
			get { return _operacao; }
			set { _operacao = value; }
		}

		private DateTime _trx_TX_data;
		public DateTime trx_TX_data
		{
			get { return _trx_TX_data; }
			set { _trx_TX_data = value; }
		}

		private DateTime _trx_TX_data_hora;
		public DateTime trx_TX_data_hora
		{
			get { return _trx_TX_data_hora; }
			set { _trx_TX_data_hora = value; }
		}

		private DateTime _trx_RX_data;
		public DateTime trx_RX_data
		{
			get { return _trx_RX_data; }
			set { _trx_RX_data = value; }
		}

		private DateTime _trx_RX_data_hora;
		public DateTime trx_RX_data_hora
		{
			get { return _trx_RX_data_hora; }
			set { _trx_RX_data_hora = value; }
		}

		private byte _trx_RX_status;
		public byte trx_RX_status
		{
			get { return _trx_RX_status; }
			set { _trx_RX_status = value; }
		}

		private byte _trx_RX_vazio_status;
		public byte trx_RX_vazio_status
		{
			get { return _trx_RX_vazio_status; }
			set { _trx_RX_vazio_status = value; }
		}

		private byte _st_sucesso;
		public byte st_sucesso
		{
			get { return _st_sucesso; }
			set { _st_sucesso = value; }
		}
	}
	#endregion

	#region [ ClearsaleAFOpComplementarXml ]
	public class ClearsaleAFOpComplementarXml
	{
		private int _id;
		public int id
		{
			get { return _id; }
			set { _id = value; }
		}

		private int _id_pagto_gw_af_op_complementar;
		public int id_pagto_gw_af_op_complementar
		{
			get { return _id_pagto_gw_af_op_complementar; }
			set { _id_pagto_gw_af_op_complementar = value; }
		}

		private DateTime _data;
		public DateTime data
		{
			get { return _data; }
			set { _data = value; }
		}

		private DateTime _data_hora;
		public DateTime data_hora
		{
			get { return _data_hora; }
			set { _data_hora = value; }
		}

		private string _tipo_transacao;
		public string tipo_transacao
		{
			get { return _tipo_transacao; }
			set { _tipo_transacao = value; }
		}

		private string _fluxo_xml;
		public string fluxo_xml
		{
			get { return _fluxo_xml; }
			set { _fluxo_xml = value; }
		}

		private string _xml;
		public string xml
		{
			get { return _xml; }
			set { _xml = value; }
		}
	}
	#endregion

	#region [ ClearsaleAnalystComments ]
	class ClearsaleAnalystComments
	{
		public string CreateDate { get; set; }
		public DateTime DataHoraCreateDate { get; set; }
		public string Comments { get; set; }
		public string UserName { get; set; }
		public string Status { get; set; }
		public string LineName { get; set; }

		public ClearsaleAnalystComments(string createDate, string comments, string userName, string status, string lineName)
		{
			CreateDate = createDate;
			DataHoraCreateDate = Global.converteDateTimeFromISO8601(createDate);
			Comments = comments;
			UserName = userName;
			Status = status;
			LineName = lineName;
		}
	}
	#endregion
}
