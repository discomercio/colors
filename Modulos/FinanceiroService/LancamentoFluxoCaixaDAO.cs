#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
#endregion

namespace FinanceiroService
{
	static class LancamentoFluxoCaixaDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmLancamentoInsertDevidoBoletoEC;
		#endregion

		#region [ inicializaConstrutorEstatico ]
		public static void inicializaConstrutorEstatico()
		{
			// NOP
			// 1) The static constructor for a class executes before any instance of the class is created.
			// 2) The static constructor for a class executes before any of the static members for the class are referenced.
			// 3) The static constructor for a class executes after the static field initializers (if any) for the class.
			// 4) The static constructor for a class executes at most one time during a single program instantiation
			// 5) A static constructor does not take access modifiers or have parameters.
			// 6) A static constructor is called automatically to initialize the class before the first instance is created or any static members are referenced.
			// 7) A static constructor cannot be called directly.
			// 8) The user has no control on when the static constructor is executed in the program.
			// 9) A typical use of static constructors is when the class is using a log file and the constructor is used to write entries to this file.
		}
		#endregion

		#region [ Construtor estático ]
		static LancamentoFluxoCaixaDAO()
		{
			inicializaObjetosEstaticos();
		}
		#endregion

		#region [ inicializaObjetosEstaticos ]
		public static void inicializaObjetosEstaticos()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmLancamentoInsertDevidoBoletoEC ]
			strSql = "INSERT INTO t_FIN_FLUXO_CAIXA (" +
						"id, " +
						"id_conta_corrente, " +
						"id_plano_contas_empresa, " +
						"id_plano_contas_grupo, " +
						"id_plano_contas_conta, " +
						"natureza, " +
						"dt_competencia, " +
						"valor, " +
						"descricao, " +
						"ctrl_pagto_id_parcela, " +
						"ctrl_pagto_modulo, " +
						"id_cliente, " +
						"cnpj_cpf, " +
						"tipo_cadastro, " +
						"editado_manual, " +
						"dt_cadastro, " +
						"dt_hr_cadastro, " +
						"usuario_cadastro, " +
						"dt_ult_atualizacao, " +
						"dt_hr_ult_atualizacao, " +
						"usuario_ult_atualizacao " +
					") VALUES (" +
						"@id, " +
						"@id_conta_corrente, " +
						"@id_plano_contas_empresa, " +
						"@id_plano_contas_grupo, " +
						"@id_plano_contas_conta, " +
						"@natureza, " +
						"@dt_competencia, " +
						"@valor, " +
						"@descricao, " +
						"@ctrl_pagto_id_parcela, " +
						"@ctrl_pagto_modulo, " +
						"@id_cliente, " +
						"@cnpj_cpf, " +
						"@tipo_cadastro, " +
						"@editado_manual, " +
						Global.sqlMontaGetdateSomenteData() + ", " +
						"getdate(), " +
						"@usuario_cadastro, " +
						Global.sqlMontaGetdateSomenteData() + ", " +
						"getdate(), " +
						"@usuario_ult_atualizacao " +
					")";
			cmLancamentoInsertDevidoBoletoEC = BD.criaSqlCommand();
			cmLancamentoInsertDevidoBoletoEC.CommandText = strSql;
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@id", SqlDbType.Int);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@id_conta_corrente", SqlDbType.TinyInt);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@id_plano_contas_empresa", SqlDbType.TinyInt);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@id_plano_contas_grupo", SqlDbType.SmallInt);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@id_plano_contas_conta", SqlDbType.Int);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@natureza", SqlDbType.Char, 1);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@dt_competencia", SqlDbType.VarChar, 10);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@valor", SqlDbType.Money);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@descricao", SqlDbType.VarChar, Global.Cte.FIN.TamanhoCampo.FLUXO_CAIXA_DESCRICAO);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@id_cliente", SqlDbType.VarChar, 12);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@cnpj_cpf", SqlDbType.VarChar, 14);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@tipo_cadastro", SqlDbType.Char, 1);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@editado_manual", SqlDbType.Char, 1);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@usuario_cadastro", SqlDbType.VarChar, 10);
			cmLancamentoInsertDevidoBoletoEC.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmLancamentoInsertDevidoBoletoEC.Prepare();
			#endregion
		}
		#endregion

		#region [ insereLancamentoDevidoBoletoEC ]
		public static bool insereLancamentoDevidoBoletoEC(LancamentoFluxoCaixaInsertDevidoBoletoEC lancamento, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "LancamentoFluxoCaixaDAO.insereLancamentoDevidoBoletoEC()";
			int id_emailsndsvc_mensagem;
			int intRetorno;
			int intNsuLancamento;
			bool blnGerouNsu;
			string msg_erro_aux;
			string strMsg;
			string strSubject;
			string strBody;
			string strDescricaoLog;
			StringBuilder sbLog = new StringBuilder("");
			FinLog finLog = new FinLog();
			#endregion

			msg_erro = "";
			try
			{
				#region [ Gera NSU ]
				blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.FIN.NSU.T_FIN_FLUXO_CAIXA, out intNsuLancamento, out msg_erro_aux);
				if (!blnGerouNsu)
				{
					msg_erro = "Falha ao gerar o NSU para o lançamento no fluxo de caixa!!\n" + msg_erro_aux;
					return false;
				}
				#endregion

				#region [ Atualiza o ID gerado p/ que seja retornado p/ a rotina chamadora ]
				lancamento.id = intNsuLancamento;
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmLancamentoInsertDevidoBoletoEC.Parameters["@id"].Value = intNsuLancamento;
				cmLancamentoInsertDevidoBoletoEC.Parameters["@id_conta_corrente"].Value = lancamento.id_conta_corrente;
				cmLancamentoInsertDevidoBoletoEC.Parameters["@id_plano_contas_empresa"].Value = lancamento.id_plano_contas_empresa;
				cmLancamentoInsertDevidoBoletoEC.Parameters["@id_plano_contas_grupo"].Value = lancamento.id_plano_contas_grupo;
				cmLancamentoInsertDevidoBoletoEC.Parameters["@id_plano_contas_conta"].Value = lancamento.id_plano_contas_conta;
				cmLancamentoInsertDevidoBoletoEC.Parameters["@natureza"].Value = Global.Cte.FIN.Natureza.CREDITO;
				cmLancamentoInsertDevidoBoletoEC.Parameters["@dt_competencia"].Value = Global.formataDataYyyyMmDdComSeparador(lancamento.dt_competencia);
				cmLancamentoInsertDevidoBoletoEC.Parameters["@valor"].Value = lancamento.valor;
				cmLancamentoInsertDevidoBoletoEC.Parameters["@descricao"].Value = lancamento.descricao;
				cmLancamentoInsertDevidoBoletoEC.Parameters["@ctrl_pagto_id_parcela"].Value = lancamento.ctrl_pagto_id_parcela;
				cmLancamentoInsertDevidoBoletoEC.Parameters["@ctrl_pagto_modulo"].Value = lancamento.ctrl_pagto_modulo;
				cmLancamentoInsertDevidoBoletoEC.Parameters["@id_cliente"].Value = (lancamento.id_cliente ?? "");
				cmLancamentoInsertDevidoBoletoEC.Parameters["@cnpj_cpf"].Value = (lancamento.cnpj_cpf ?? "");
				cmLancamentoInsertDevidoBoletoEC.Parameters["@tipo_cadastro"].Value = Global.Cte.FIN.TipoCadastro.SISTEMA;
				cmLancamentoInsertDevidoBoletoEC.Parameters["@editado_manual"].Value = Global.Cte.FIN.EditadoManual.NAO;
				cmLancamentoInsertDevidoBoletoEC.Parameters["@usuario_cadastro"].Value = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				cmLancamentoInsertDevidoBoletoEC.Parameters["@usuario_ult_atualizacao"].Value = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Monta texto p/ o log ]
				foreach (SqlParameter item in cmLancamentoInsertDevidoBoletoEC.Parameters)
				{
					if (sbLog.Length > 0) sbLog.Append("; ");
					sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
				}
				#endregion

				#region [ Tenta inserir o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmLancamentoInsertDevidoBoletoEC);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.Message;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(lancamento);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": falha ao tentar inserir lançamento no fluxo de caixa devido a boleto de e-commerce [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\r\n" + NOME_DESTA_ROTINA + ": falha ao tentar inserir lançamento no fluxo de caixa devido a boleto de e-commerce (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + lancamento.ctrl_pagto_id_parcela.ToString() + ")\r\n\r\nDetalhes dos dados a serem gravados:\r\n" + Global.serializaObjectToXml(lancamento) + "\r\n\r\nInformações sobre o exception:\r\n" + ex.ToString();
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
					}
				}
				#endregion

				#region [ Gravou o registro? ]
				if (intRetorno == 0)
				{
					msg_erro = "Falha ao tentar inserir o lançamento no fluxo de caixa devido a boleto de e-commerce!!\n" + msg_erro;
					Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": " + msg_erro);
					return false;
				}
				#endregion

				#region [ Registra log com a inclusão dos dados ]
				strDescricaoLog = "Inserção do registro em t_FIN_FLUXO_CAIXA.id=" + intNsuLancamento.ToString() + " devido a boleto de e-commerce (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + lancamento.ctrl_pagto_id_parcela.ToString() + "): " + sbLog.ToString();
				Global.gravaLogAtividade(strDescricaoLog);
				finLog.usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				finLog.operacao = Global.Cte.FIN.LogOperacao.FLUXO_CAIXA_INSERE_DEVIDO_BOLETO_ECOMMERCE;
				finLog.natureza = Global.Cte.FIN.Natureza.CREDITO;
				finLog.tipo_cadastro = Global.Cte.FIN.TipoCadastro.SISTEMA;
				finLog.fin_modulo = Global.Cte.FIN.Modulo.FINANCEIRO_SERVICE;
				finLog.cod_tabela_origem = Global.Cte.FIN.TabelaOrigem.T_FIN_FLUXO_CAIXA;
				finLog.id_registro_origem = intNsuLancamento;
				finLog.id_conta_corrente = lancamento.id_conta_corrente;
				finLog.id_plano_contas_empresa = lancamento.id_plano_contas_empresa;
				finLog.id_plano_contas_grupo = lancamento.id_plano_contas_grupo;
				finLog.id_plano_contas_conta = lancamento.id_plano_contas_conta;
				finLog.id_cliente = lancamento.id_cliente;
				finLog.cnpj_cpf = Global.digitos(lancamento.cnpj_cpf);
				finLog.descricao = strDescricaoLog;
				FinLogDAO.insere(Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA, finLog, ref msg_erro_aux);
				#endregion

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Sucesso na inserção do lançamento no fluxo de caixa devido a boleto de e-commerce (" + Global.Cte.FIN.NSU.T_FIN_FLUXO_CAIXA + ".id=" + lancamento.id.ToString() + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + lancamento.ctrl_pagto_id_parcela.ToString() + ")");

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(lancamento);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": falha ao tentar inserir lançamento no fluxo de caixa devido a boleto de e-commerce [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
				strBody = "Mensagem de Financeiro Service\r\n" + NOME_DESTA_ROTINA + ": falha ao tentar inserir lançamento no fluxo de caixa devido a boleto de e-commerce (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + lancamento.ctrl_pagto_id_parcela.ToString() + ")\r\n\r\nDetalhes dos dados a serem gravados:\r\n" + Global.serializaObjectToXml(lancamento) + "\r\n\r\nInformações sobre o exception:\r\n" + ex.ToString();
				if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
				{
					strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
					Global.gravaLogAtividade(strMsg);
				}

				return false;
			}
		}
		#endregion
	}
}
