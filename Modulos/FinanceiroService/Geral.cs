using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace FinanceiroService
{
	#region [ Geral ]
	public static class Geral
	{
		#region [ Métodos ]

		#region [ executaProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito ]
		public static bool executaProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito(out bool blnEmailAlertaEnviado, out string strMsgInformativa, out string msg_erro)
		{
			#region [ Declarações ]
			const String NOME_DESTA_ROTINA = "Geral.executaProcessamentoEnviarEmailAlertaPedidoNovoAnaliseCredito()";
			int qtdePedidos = 0;
			int id_emailsndsvc_mensagem;
			string msg_erro_aux;
			string strSql;
			string strMsg;
			string strSubject;
			string strBody;
			StringBuilder sbNumeroPedido = new StringBuilder("");
			StringBuilder sbDetalhePedido = new StringBuilder("");
			StringBuilder sbMsgInformativa = new StringBuilder("");
			StringBuilder sbBody = new StringBuilder("");
			List<string> listaPedido = new List<string>();
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			#endregion

			blnEmailAlertaEnviado = false;
			strMsgInformativa = "";
			msg_erro = "";
			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				#region [ Monta SQL ]
				// A consulta seleciona os pedidos aprovados automaticamente pela Clearsale, pois nesse caso o sistema coloca o status da
				// análise de crédito como "Pendente Vendas" para que o analista possa verificar se o nome do titular do cartão diverge do
				// cliente que está realizando a compra. O email de alerta informa que há pedidos para serem tratados na fila de "Pendente Vendas".
				strSql = "SELECT " +
							"*" +
						" FROM " +
							"(" +
							"SELECT" +
								" pedido," +
								" data_hora," +
								" vendedor," +
								" indicador," +
								" analise_credito," +
								" vl_total_NF," +
								" vl_previsto_cartao," +
								" tipo_parcelamento," +
								" av_forma_pagto," +
								" pu_forma_pagto, pu_valor, pu_vencto_apos," +
								" pc_qtde_parcelas, pc_valor_parcela," +
								" pc_maquineta_qtde_parcelas, pc_maquineta_valor_parcela," +
								" pce_forma_pagto_entrada, pce_forma_pagto_prestacao, pce_entrada_valor, pce_prestacao_qtde, pce_prestacao_valor, pce_prestacao_periodo," +
								" pse_forma_pagto_prim_prest, pse_forma_pagto_demais_prest, pse_prim_prest_valor, pse_prim_prest_apos, pse_demais_prest_qtde, pse_demais_prest_valor, pse_demais_prest_periodo," +
								" (" +
									"SELECT" +
										" SUM(payment.valor_transacao)" +
									" FROM t_PAGTO_GW_PAG pag INNER JOIN t_PAGTO_GW_PAG_PAYMENT payment ON (pag.id = payment.id_pagto_gw_pag)" +
									" WHERE" +
										" (pag.pedido = t_PEDIDO.pedido)" +
										" AND" +
										" (ult_GlobalStatus = '" + Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue() + "')" +
								") AS vl_pago_cartao" +
							" FROM t_PEDIDO" +
							" WHERE" +
								" (analise_credito = " + Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS.ToString() + ")" +
								" AND (st_pedido_novo_analise_credito_msg_alerta = 0)" +
								" AND (st_forma_pagto_possui_parcela_cartao = 1)" +
								" AND (st_entrega <> '" + Global.Cte.StEntregaPedido.ST_ENTREGA_ENTREGUE + "')" +
								" AND (st_entrega <> '" + Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO + "')" +
							") t" +
						" WHERE" +
							" (vl_pago_cartao >= vl_previsto_cartao)" +
						" ORDER BY" +
							" data_hora," +
							" pedido";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				#endregion

				#region [ Processa o resultado ]
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					row = dtbResultado.Rows[i];
					qtdePedidos++;

					if (sbNumeroPedido.Length > 0) sbNumeroPedido.Append(", ");
					sbNumeroPedido.Append(BD.readToString(row["pedido"]));

					listaPedido.Add(BD.readToString(row["pedido"]));

					#region [ Prepara mensagem de alerta ]
					strMsg = "Pedido " + BD.readToString(row["pedido"]) + " cadastrado em " + Global.formataDataDdMmYyyyHhMmComSeparador(BD.readToDateTime(row["data_hora"])) + " por " + BD.readToString(row["vendedor"]) + " (" + Global.obtemDescricaoAnaliseCredito(BD.readToShort(row["analise_credito"])) + ")";
					sbDetalhePedido.AppendLine(strMsg);
					#endregion
				}
				#endregion

				#region [ Há pedidos? ]
				if (qtdePedidos == 0)
				{
					if (sbMsgInformativa.Length > 0) sbMsgInformativa.Append("; ");
					sbMsgInformativa.Append("Nenhum pedido novo aguardando tratamento da análise de crédito");
					return true;
				}
				#endregion

				#region [ Monta mensagem informativa desta rotina ]
				if (sbMsgInformativa.Length > 0) sbMsgInformativa.Append("; ");
				strMsg = qtdePedidos.ToString() +
						(qtdePedidos == 1 ?
							" pedido novo aguardando tratamento da análise de crédito foi informado no email de alerta: "
							:
							" pedidos novos aguardando tratamento da análise de crédito foram informados no email de alerta: "
						) + sbNumeroPedido.ToString();
				sbMsgInformativa.Append(strMsg);
				#endregion

				#region [ Monta mensagem do email de alerta ]
				strMsg = qtdePedidos.ToString() +
						(qtdePedidos == 1 ?
							" pedido novo aguardando tratamento da análise de crédito"
							:
							" pedidos novos aguardando tratamento da análise de crédito");
				strSubject = Global.montaIdInstanciaServicoEmailSubject() + "  " + strMsg + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
				strBody = "Ambiente: " + (Global.Cte.Aplicativo.IDENTIFICADOR_AMBIENTE_OWNER.Length > 0 ? Global.Cte.Aplicativo.IDENTIFICADOR_AMBIENTE_OWNER : Global.Cte.Aplicativo.ID_SISTEMA_EVENTLOG) +
						"\r\n\r\n" +
						qtdePedidos.ToString() +
						(qtdePedidos == 1 ?
							" pedido novo aguardando tratamento da análise de crédito"
							:
							" pedidos novos aguardando tratamento da análise de crédito") +
						"\r\n\r\n" +
						sbDetalhePedido.ToString();
				#endregion

				#region [ Envia mensagem de alerta ]
				if (EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Parametros.Geral.DESTINATARIO_MSG_ALERTA_PEDIDO_NOVO_ANALISE_CREDITO, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
				{
					#region [ Registra no pedido que já foi enviado um email de alerta para que ele não seja incluído novamente no próximo alerta ]
					for (int i = 0; i < listaPedido.Count; i++)
					{
						if (listaPedido[i].Trim().Length > 0)
						{
							if (!PedidoDAO.updatePedidoStPedidoNovoAnaliseCreditoMsgAlertaFlagAtivo(listaPedido[i], out msg_erro_aux))
							{
								if (msg_erro.Length > 0) msg_erro += "\r\n";
								msg_erro += "Falha ao tentar atualizar o campo t_PEDIDO.st_pedido_novo_analise_credito_msg_alerta = 1 no pedido " + listaPedido[i] + ": " + msg_erro_aux;
							}
						}
					}
					#endregion
				}
				else
				{
					msg_erro = "Falha ao tentar inserir email de alerta sobre pedido novo aguardando tratamento da análise de crédito na fila de mensagens!!\n" + msg_erro_aux;

					strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta sobre pedido novo aguardando tratamento da análise de crédito na fila de mensagens!!\n" + msg_erro_aux;
					Global.gravaLogAtividade(strMsg);

					return false;
				}
				#endregion
				
				blnEmailAlertaEnviado = true;

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

		#region [ executaEnvioEmailAlertaVendedoresPrazoPrevisaoEntrega ]
		public static bool executaEnvioEmailAlertaVendedoresPrazoPrevisaoEntrega(out string strMsgInformativa, out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "Geral.executaEnvioEmailAlertaVendedoresPrazoPrevisaoEntrega()";
			const string TAB_x1 = "    ";
			const string TAB_x2 = TAB_x1 + TAB_x1;
			string strLinhaSeparadoraInicio = new string('-', 30) + "( Início )" + new string('-', 30);
			string strLinhaSeparadoraFim = new string('-', 31) + "( Fim )" + new string('-', 32);
			int loja;
			int periodoCorteEmDias;
			int qtdeEmailsEnviados = 0;
			int qtdeEmailsNaoEnviados = 0;
			int qtdeEmailsFalhaEnvio = 0;
			int qtdeTotalPedidos = 0;
			bool bIncluirVendedor;
			bool bIncluirData;
			string strMsg;
			string strSql;
			string strLoja;
			string strLojasIgnoradas;
			string strVendedor;
			string strWhereLojasIgnoradas = "";
			string[] vLojasIgnoradas;
			string sRemetente;
			StringBuilder sbPedido;
			StringBuilder sbMsgVendedor;
			StringBuilder sbMsgEnviado = new StringBuilder("");
			StringBuilder sbMsgNaoEnviado = new StringBuilder("");
			StringBuilder sbMsgFalhaEnvio = new StringBuilder("");
			List<EmailAlertaVendedoresPrazoPrevisaoEntrega> listaEmailAlerta = new List<EmailAlertaVendedoresPrazoPrevisaoEntrega>();
			EmailAlertaVendedoresPrazoPrevisaoEntrega emailAlertaVendedor;
			EmailAlertaVendedoresPrazoPrevisaoEntregaPedido emailAlertaPorPedido;
			EmailAlertaVendedoresPrazoPrevisaoEntregaAgrupaPorData emailAlertaPorData;
			DateTime dtDataCorte;
			DateTime dtPrevisaoEntregaData;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
			int id_emailsndsvc_mensagem;
			string msg_erro_aux;
			string strSubject;
			string strBody;
			#endregion

			strMsgInformativa = "";
			strMsgErro = "";

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

				#region [ Obtém remetente a ser usado no envio ]
				sRemetente = GeralDAO.getCampoTextoTabelaParametro(Global.Cte.FIN.ID_T_PARAMETRO.ID_PARAMETRO_EMAILSNDSVC_REMETENTE__MENSAGEM_SISTEMA);
				if (sRemetente.Trim().Length == 0) sRemetente = Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA;
				#endregion

				#region [ Prepara restrição para lojas ignoradas ]
				strLojasIgnoradas = (Global.Parametros.Geral.EnvioEmailAlertaVendedoresPrazoPrevisaoEntregaLojasIgnoradas ?? "").Trim();
				strLojasIgnoradas = strLojasIgnoradas.Replace(';', ',');
				strLojasIgnoradas = strLojasIgnoradas.Replace('|', ',');
				vLojasIgnoradas = strLojasIgnoradas.Split(',');
				for (int i = 0; i < vLojasIgnoradas.Length; i++)
				{
					strLoja = (vLojasIgnoradas[i] ?? "");
					if (strLoja.Length > 0)
					{
						loja = (int)Global.converteInteiro(strLoja);
						if (loja > 0)
						{
							if (strWhereLojasIgnoradas.Length > 0) strWhereLojasIgnoradas += ",";
							strWhereLojasIgnoradas += loja.ToString();
						}
					}
				}
				if (strWhereLojasIgnoradas.Length > 0)
				{
					strWhereLojasIgnoradas = " AND (tPedBase.numero_loja NOT IN (" + strWhereLojasIgnoradas + "))";
				}
				#endregion

				#region [ Prepara restrição por período de corte ]
				periodoCorteEmDias = Global.Parametros.Geral.EnvioEmailAlertaVendedoresPrazoPrevisaoEntregaPeriodoCorte;
				dtDataCorte = DateTime.Today.AddDays(Math.Abs(periodoCorteEmDias));
				#endregion

				#region [ Monta o SQL da consulta ]
				strSql = "SELECT" +
							" tPed.vendedor" +
							", tPed.pedido" +
							", tPed.PrevisaoEntregaData" +
							", tPedBase.loja" +
							", tPed.endereco_nome_iniciais_em_maiusculas AS cliente_nome" +
							", tPed.endereco_cnpj_cpf AS cliente_cnpj_cpf" +
							", t_USUARIO.email AS vendedor_email" +
						" FROM t_PEDIDO tPed" +
							" INNER JOIN t_PEDIDO AS tPedBase ON (tPed.pedido_base=tPedBase.pedido)" +
							" INNER JOIN t_USUARIO ON (tPedBase.vendedor = t_USUARIO.usuario)" +
						" WHERE" +
							" (tPed.st_etg_imediata = " + Global.Cte.T_PEDIDO__ENTREGA_IMEDIATA_STATUS.ETG_IMEDIATA_NAO.ToString() + ")" +
							" AND ( (tPed.PrevisaoEntregaData IS NULL) OR (tPed.PrevisaoEntregaData <= " + Global.sqlMontaDateTimeParaSqlDateTimeSomenteData(dtDataCorte) + ") )" +
							" AND (tPedBase.analise_credito = " + Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK.ToString() + ")" +
							" AND (tPed.st_entrega NOT IN ('" + Global.Cte.StEntregaPedido.ST_ENTREGA_ENTREGUE + "','" + Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO + "'))" +
							strWhereLojasIgnoradas +
						" ORDER BY" +
							" tPedBase.vendedor" +
							", tPed.PrevisaoEntregaData" +
							", tPed.pedido";
				#endregion

				#region [ Log informativo da consulta realizada ]
				strMsg = NOME_DESTA_ROTINA + ":\r\n" + strSql;
				Global.gravaLogAtividade(strMsg);
				#endregion

				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				#endregion

				#region [ Processa o resultado da consulta ao BD em uma lista ]
				for (int i = 0; i < dtbConsulta.Rows.Count; i++)
				{
					rowConsulta = dtbConsulta.Rows[i];
					strVendedor = BD.readToString(rowConsulta["vendedor"]);
					dtPrevisaoEntregaData = BD.readToDateTime(rowConsulta["PrevisaoEntregaData"]);

					try
					{
						emailAlertaVendedor = listaEmailAlerta.Single(p => p.vendedor.ToUpper().Equals(strVendedor.ToUpper()));
						bIncluirVendedor = false;
					}
					catch (Exception)
					{
						emailAlertaVendedor = null;
						bIncluirVendedor = true;
					}

					if (bIncluirVendedor)
					{
						emailAlertaVendedor = new EmailAlertaVendedoresPrazoPrevisaoEntrega();
						emailAlertaVendedor.vendedor = strVendedor;
						emailAlertaVendedor.email = BD.readToString(rowConsulta["vendedor_email"]);
					}

					emailAlertaPorPedido = new EmailAlertaVendedoresPrazoPrevisaoEntregaPedido();
					emailAlertaPorPedido.pedido = BD.readToString(rowConsulta["pedido"]);
					emailAlertaPorPedido.loja = BD.readToString(rowConsulta["loja"]);
					emailAlertaPorPedido.PrevisaoEntregaData = BD.readToDateTime(rowConsulta["PrevisaoEntregaData"]);
					emailAlertaPorPedido.cliente_nome = BD.readToString(rowConsulta["cliente_nome"]);
					emailAlertaPorPedido.cliente_cnpj_cpf = BD.readToString(rowConsulta["cliente_cnpj_cpf"]);

					emailAlertaVendedor.listaPorPedido.Add(emailAlertaPorPedido);

					qtdeTotalPedidos++;

					try
					{
						emailAlertaPorData = emailAlertaVendedor.listaPorData.Single(p => p.PrevisaoEntregaData == dtPrevisaoEntregaData);
						bIncluirData = false;
					}
					catch (Exception)
					{
						emailAlertaPorData = null;
						bIncluirData = true;
					}

					if (bIncluirData)
					{
						emailAlertaPorData = new EmailAlertaVendedoresPrazoPrevisaoEntregaAgrupaPorData();
						emailAlertaPorData.PrevisaoEntregaData = dtPrevisaoEntregaData;
					}

					emailAlertaPorData.listaPedido.Add(emailAlertaPorPedido);
					if (bIncluirData) emailAlertaVendedor.listaPorData.Add(emailAlertaPorData);

					if (bIncluirVendedor) listaEmailAlerta.Add(emailAlertaVendedor);
				}
				#endregion

				#region [ Processa o envio dos e-mails a partir da lista ]
				foreach (EmailAlertaVendedoresPrazoPrevisaoEntrega vendedor in listaEmailAlerta)
				{
					#region [ Monta mensagem informativa para ser enviada no e-mail de resumo geral ]
					sbMsgVendedor = new StringBuilder("");
					foreach (EmailAlertaVendedoresPrazoPrevisaoEntregaAgrupaPorData data in vendedor.listaPorData)
					{
						sbPedido = new StringBuilder("");
						foreach (EmailAlertaVendedoresPrazoPrevisaoEntregaPedido pedido in data.listaPedido)
						{
							if (sbPedido.Length > 0) sbPedido.Append(", ");
							sbPedido.Append(pedido.pedido);
						}

						sbMsgVendedor.AppendLine(TAB_x2 + ((data.PrevisaoEntregaData == DateTime.MinValue) ? "(data não preenchida)" : Global.formataDataDdMmYyyyComSeparador(data.PrevisaoEntregaData)) + ": " + sbPedido.ToString());
					}
					#endregion

					if (vendedor.email.Length == 0)
					{
						// O vendedor não possui e-mail cadastrado, então não será possível enviar o e-mail de alerta
						qtdeEmailsNaoEnviados++;

						#region [ Registra os dados para incluir no e-mail informativo de resumo geral ]
						if (sbMsgNaoEnviado.Length > 0) sbMsgNaoEnviado.AppendLine("");
						sbMsgNaoEnviado.AppendLine("Vendedor: " + vendedor.vendedor);
						sbMsgNaoEnviado.AppendLine("E-mail: " + (vendedor.email.Length > 0 ? vendedor.email : "(e-mail não cadastrado)"));
						sbMsgNaoEnviado.AppendLine("Qtde pedidos: " + vendedor.listaPorPedido.Count.ToString());
						sbMsgNaoEnviado.AppendLine(sbMsgVendedor.ToString());
						#endregion
					}
					else
					{
						#region [ Envia e-mail de alerta para o vendedor ]
						strSubject = "Pedidos com entrega imediata 'Não': informativo sobre previsão de entrega próxima de expirar ou já expirada";
						strBody = "Informativo sobre previsão de entrega próxima de expirar ou já expirada para pedidos com entrega imediata 'Não'"
								+ "\r\n"
								+ "\r\n"
								+ sbMsgVendedor.ToString()
								+ "\r\n"
								+ "\r\n"
								+ "Atenção: esta é uma mensagem automática, NÃO responda a este e-mail!";
						if (EmailSndSvcDAO.gravaMensagemParaEnvio(sRemetente, vendedor.email, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							qtdeEmailsEnviados++;

							#region [ Registra os dados para incluir no e-mail informativo de resumo geral ]
							if (sbMsgEnviado.Length > 0) sbMsgEnviado.AppendLine("");
							sbMsgEnviado.AppendLine("Vendedor: " + vendedor.vendedor);
							sbMsgEnviado.AppendLine("E-mail: " + vendedor.email);
							sbMsgEnviado.AppendLine("Qtde pedidos: " + vendedor.listaPorPedido.Count.ToString());
							sbMsgEnviado.AppendLine(sbMsgVendedor.ToString());
							#endregion
						}
						else
						{
							qtdeEmailsFalhaEnvio++;

							#region [ Registra os dados para incluir no e-mail informativo de resumo geral ]
							if (sbMsgFalhaEnvio.Length > 0) sbMsgFalhaEnvio.AppendLine("");
							sbMsgFalhaEnvio.AppendLine("Vendedor: " + vendedor.vendedor);
							sbMsgFalhaEnvio.AppendLine("E-mail: " + vendedor.email);
							sbMsgFalhaEnvio.AppendLine("Qtde pedidos: " + vendedor.listaPorPedido.Count.ToString());
							sbMsgFalhaEnvio.AppendLine(sbMsgVendedor.ToString());
							#endregion

							if (strMsgErro.Length > 0) strMsgErro += "\r\n\r\n";
							strMsgErro += "Falha ao tentar inserir email de alerta na fila de mensagens sobre pedidos com data de previsão de entrega próxima de expirar ou já expirada para o vendedor '" + vendedor.vendedor + "'!!\n" + msg_erro_aux;

							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens sobre pedidos com data de previsão de entrega próxima de expirar ou já expirada!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}
						#endregion
					}
				}
				#endregion

				strMsgInformativa = "Qtde de vendedores: " + listaEmailAlerta.Count.ToString()
						+ "\r\n"
						+ "Qtde de pedidos: " + Global.formataInteiro(qtdeTotalPedidos)
						+ "\r\n"
						+ "E-mails enviados: " + qtdeEmailsEnviados.ToString()
						+ "\r\n"
						+ "E-mails não enviados por ausência de e-mail no cadastro: " + qtdeEmailsNaoEnviados.ToString()
						+ "\r\n"
						+ "E-mails com falha no envio: " + qtdeEmailsFalhaEnvio.ToString()
						+ "\r\n"
						+ "\r\n"
						+ "Informações detalhadas sobre e-mails enviados"
						+ "\r\n"
						+ strLinhaSeparadoraInicio
						+ "\r\n"
						+ sbMsgEnviado.ToString()
						+ "\r\n"
						+ strLinhaSeparadoraFim
						+ "\r\n"
						+ "\r\n"
						+ "\r\n"
						+ "Informações detalhadas sobre e-mails não enviados"
						+ "\r\n"
						+ strLinhaSeparadoraInicio
						+ "\r\n"
						+ sbMsgNaoEnviado.ToString()
						+ "\r\n"
						+ strLinhaSeparadoraFim
						+ "\r\n"
						+ "\r\n"
						+ "\r\n"
						+ "Informações detalhadas sobre e-mails com falha no envio"
						+ "\r\n"
						+ strLinhaSeparadoraInicio
						+ "\r\n"
						+ sbMsgFalhaEnvio.ToString()
						+ "\r\n"
						+ strLinhaSeparadoraFim;

				#region [ Envia e-mail de resumo geral? ]
				if (Global.Parametros.Geral.EnvioEmailAlertaVendedoresPrazoPrevisaoEntregaDestinatarioResumoGeral.Length > 0)
				{
					strSubject = "Pedidos com entrega imediata 'Não': informativo com resumo geral sobre previsão de entrega próxima de expirar ou já expirada";
					strBody = "Informativo com resumo geral sobre previsão de entrega próxima de expirar ou já expirada para pedidos com entrega imediata 'Não'"
								+ "\r\n"
								+ "\r\n"
								+ strMsgInformativa
								+ "\r\n"
								+ "\r\n"
								+ "Atenção: esta é uma mensagem automática, NÃO responda a este e-mail!";
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(sRemetente, Global.Parametros.Geral.EnvioEmailAlertaVendedoresPrazoPrevisaoEntregaDestinatarioResumoGeral, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta com o resumo geral na fila de mensagens sobre pedidos com data de previsão de entrega próxima de expirar ou já expirada!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
					}
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#endregion
	}
	#endregion

	#region [ FinSvcLog ]
	class FinSvcLog
	{
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

		private string _operacao;
		public string operacao
		{
			get { return _operacao; }
			set { _operacao = value; }
		}

		private string _tabela;
		public string tabela
		{
			get { return _tabela; }
			set { _tabela = value; }
		}

		private string _descricao;
		public string descricao
		{
			get { return _descricao; }
			set { _descricao = value; }
		}

		private string _complemento_1;
		public string complemento_1
		{
			get { return _complemento_1; }
			set { _complemento_1 = value; }
		}

		private string _complemento_2;
		public string complemento_2
		{
			get { return _complemento_2; }
			set { _complemento_2 = value; }
		}

		private string _complemento_3;
		public string complemento_3
		{
			get { return _complemento_3; }
			set { _complemento_3 = value; }
		}

		private string _complemento_4;
		public string complemento_4
		{
			get { return _complemento_4; }
			set { _complemento_4 = value; }
		}

		private string _complemento_5;
		public string complemento_5
		{
			get { return _complemento_5; }
			set { _complemento_5 = value; }
		}

		private string _complemento_6;
		public string complemento_6
		{
			get { return _complemento_6; }
			set { _complemento_6 = value; }
		}
	}
	#endregion

	#region [ EmailAlertaVendedoresPrazoPrevisaoEntrega ]
	class EmailAlertaVendedoresPrazoPrevisaoEntrega
	{
		public string vendedor { get; set; } = "";
		public string email { get; set; } = "";
		public List<EmailAlertaVendedoresPrazoPrevisaoEntregaPedido> listaPorPedido { get; set; } = new List<EmailAlertaVendedoresPrazoPrevisaoEntregaPedido>();
		public List<EmailAlertaVendedoresPrazoPrevisaoEntregaAgrupaPorData> listaPorData { get; set; } = new List<EmailAlertaVendedoresPrazoPrevisaoEntregaAgrupaPorData>();
	}
	#endregion

	#region [ EmailAlertaVendedoresPrazoPrevisaoEntregaPedido ]
	class EmailAlertaVendedoresPrazoPrevisaoEntregaPedido
	{
		public string pedido { get; set; } = "";
		public string loja { get; set; } = "";
		public DateTime PrevisaoEntregaData { get; set; } = DateTime.MinValue;
		public string cliente_nome { get; set; } = "";
		public string cliente_cnpj_cpf { get; set; } = "";
	}
	#endregion

	#region [ EmailAlertaVendedoresPrazoPrevisaoEntregaAgrupaPorData ]
	class EmailAlertaVendedoresPrazoPrevisaoEntregaAgrupaPorData
	{
		public DateTime PrevisaoEntregaData { get; set; } = DateTime.MinValue;
		public List<EmailAlertaVendedoresPrazoPrevisaoEntregaPedido> listaPedido { get; set; } = new List<EmailAlertaVendedoresPrazoPrevisaoEntregaPedido>();
	}
	#endregion
}
