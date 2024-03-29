﻿#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
#endregion

namespace Financeiro
{
	class BoletoDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmBoletoInsert;
		private static SqlCommand cmBoletoItemInsert;
		private static SqlCommand cmBoletoItemRateioInsert;
		private static SqlCommand cmBoletoMarcaEnviadoRemessaBanco;
		private static SqlCommand cmBoletoItemMarcaEnviadoRemessaBanco;
		private static SqlCommand cmBoletoMarcaCanceladoManual;
		private static SqlCommand cmBoletoItemMarcaCanceladoManual;
		private static SqlCommand cmBoletoItemMarcaCanceladoManualByIdBoleto;
		private static SqlCommand cmBoletoArqRemessaInsert;
		private static SqlCommand cmBoletoArqRetornoInsert;
		private static SqlCommand cmBoletoArqRetornoUpdate;
		private static SqlCommand cmBoletoItemAtualizaOcorrencia02;
		private static SqlCommand cmBoletoItemAtualizaOcorrencia06;
		private static SqlCommand cmBoletoItemAtualizaOcorrencia09;
		private static SqlCommand cmBoletoItemAtualizaOcorrencia10;
		private static SqlCommand cmBoletoItemAtualizaOcorrencia12;
		private static SqlCommand cmBoletoItemAtualizaOcorrencia13;
		private static SqlCommand cmBoletoItemAtualizaOcorrencia14;
		private static SqlCommand cmBoletoItemAtualizaOcorrencia15;
		private static SqlCommand cmBoletoItemAtualizaOcorrencia16;
		private static SqlCommand cmBoletoItemAtualizaOcorrencia17;
		private static SqlCommand cmBoletoItemAtualizaOcorrencia19;
		private static SqlCommand cmBoletoItemAtualizaOcorrencia22;
		private static SqlCommand cmBoletoItemAtualizaOcorrencia23;
		private static SqlCommand cmBoletoItemAtualizaOcorrenciaValaComum;
		private static SqlCommand cmBoletoItemAtualizaOcorrencia24;
		private static SqlCommand cmBoletoItemAtualizaOcorrencia34;
		private static SqlCommand cmBoletoMovimentoInsert;
		private static SqlCommand cmBoletoOcorrenciaInsert;
		private static SqlCommand cmBoletoCorrigeOcorrencia24CepIrregular;
		private static SqlCommand cmBoletoOcorrenciaMarcaComoJaTratada;
		private static SqlCommand cmBoletoOcorrenciasMarcaComoJaTratadasByIdBoleto;
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
		static BoletoDAO()
		{
			inicializaObjetosEstaticos();
		}
		#endregion

		#region [ Métodos ]

		#region [ inicializaObjetosEstaticos ]
		public static void inicializaObjetosEstaticos()
		{
			#region [ Declarações ]
			String strSql;
			#endregion

			#region [ cmBoletoInsert ]
			strSql = "INSERT INTO t_FIN_BOLETO (" +
						"id, " +
						"id_cliente, " +
						"id_nf_parcela_pagto, " +
						"tipo_vinculo, " +
						"status, " +
						"numero_NF, " +
						"num_documento_boleto_avulso, " +
						"qtde_parcelas, " +
						"id_boleto_cedente, " +
						"codigo_empresa, " +
						"nome_empresa, " +
						"num_banco, " +
						"nome_banco, " +
						"agencia, " +
						"digito_agencia, " +
						"conta, " +
						"digito_conta, " +
						"carteira, " +
						"juros_mora, " +
						"perc_multa, " +
						"primeira_instrucao, " +
						"segunda_instrucao, " +
						"qtde_dias_protesto, " +
						"qtde_dias_decurso_prazo, " +
						"tipo_sacado, " +
						"num_inscricao_sacado, " +
						"nome_sacado, " +
						"endereco_sacado, " +
						"cep_sacado, " +
						"bairro_sacado, " +
						"cidade_sacado, " +
						"uf_sacado, " +
						"email_sacado, " +
						"segunda_mensagem, " +
						"mensagem_1, " +
						"mensagem_2, " +
						"mensagem_3, " +
						"mensagem_4, " +
						"usuario_cadastro, " +
						"usuario_ult_atualizacao" +
					") VALUES (" +
						"@id, " +
						"@id_cliente, " +
						"@id_nf_parcela_pagto, " +
						"@tipo_vinculo, " +
						"@status, " +
						"@numero_NF, " +
						"@num_documento_boleto_avulso, " +
						"@qtde_parcelas, " +
						"@id_boleto_cedente, " +
						"@codigo_empresa, " +
						"@nome_empresa, " +
						"@num_banco, " +
						"@nome_banco, " +
						"@agencia, " +
						"@digito_agencia, " +
						"@conta, " +
						"@digito_conta, " +
						"@carteira, " +
						"@juros_mora, " +
						"@perc_multa, " +
						"@primeira_instrucao, " +
						"@segunda_instrucao, " +
						"@qtde_dias_protesto, " +
						"@qtde_dias_decurso_prazo, " +
						"@tipo_sacado, " +
						"@num_inscricao_sacado, " +
						"@nome_sacado, " +
						"@endereco_sacado, " +
						"@cep_sacado, " +
						"@bairro_sacado, " +
						"@cidade_sacado, " +
						"@uf_sacado, " +
						"@email_sacado, " +
						"@segunda_mensagem, " +
						"@mensagem_1, " +
						"@mensagem_2, " +
						"@mensagem_3, " +
						"@mensagem_4, " +
						"@usuario_cadastro, " +
						"@usuario_ult_atualizacao" +
					")";
			cmBoletoInsert = BD.criaSqlCommand();
			cmBoletoInsert.CommandText = strSql;
			cmBoletoInsert.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoInsert.Parameters.Add("@id_cliente", SqlDbType.VarChar, 12);
			cmBoletoInsert.Parameters.Add("@id_nf_parcela_pagto", SqlDbType.Int);
			cmBoletoInsert.Parameters.Add("@tipo_vinculo", SqlDbType.TinyInt);
			cmBoletoInsert.Parameters.Add("@status", SqlDbType.SmallInt);
			cmBoletoInsert.Parameters.Add("@numero_NF", SqlDbType.Int);
			cmBoletoInsert.Parameters.Add("@num_documento_boleto_avulso", SqlDbType.Int);
			cmBoletoInsert.Parameters.Add("@qtde_parcelas", SqlDbType.TinyInt);
			cmBoletoInsert.Parameters.Add("@id_boleto_cedente", SqlDbType.SmallInt);
			cmBoletoInsert.Parameters.Add("@codigo_empresa", SqlDbType.VarChar, 20);
			cmBoletoInsert.Parameters.Add("@nome_empresa", SqlDbType.VarChar, 30);
			cmBoletoInsert.Parameters.Add("@num_banco", SqlDbType.VarChar, 3);
			cmBoletoInsert.Parameters.Add("@nome_banco", SqlDbType.VarChar, 15);
			cmBoletoInsert.Parameters.Add("@agencia", SqlDbType.VarChar, 5);
			cmBoletoInsert.Parameters.Add("@digito_agencia", SqlDbType.VarChar, 1);
			cmBoletoInsert.Parameters.Add("@conta", SqlDbType.VarChar, 7);
			cmBoletoInsert.Parameters.Add("@digito_conta", SqlDbType.VarChar, 1);
			cmBoletoInsert.Parameters.Add("@carteira", SqlDbType.VarChar, 3);
			cmBoletoInsert.Parameters.Add("@juros_mora", SqlDbType.Real);
			cmBoletoInsert.Parameters.Add("@perc_multa", SqlDbType.Real);
			cmBoletoInsert.Parameters.Add("@primeira_instrucao", SqlDbType.VarChar, 2);
			cmBoletoInsert.Parameters.Add("@segunda_instrucao", SqlDbType.VarChar, 2);
			cmBoletoInsert.Parameters.Add("@qtde_dias_protesto", SqlDbType.SmallInt);
			cmBoletoInsert.Parameters.Add("@qtde_dias_decurso_prazo", SqlDbType.SmallInt);
			cmBoletoInsert.Parameters.Add("@tipo_sacado", SqlDbType.VarChar, 2);
			cmBoletoInsert.Parameters.Add("@num_inscricao_sacado", SqlDbType.VarChar, 14);
			cmBoletoInsert.Parameters.Add("@nome_sacado", SqlDbType.VarChar, 40);
			cmBoletoInsert.Parameters.Add("@endereco_sacado", SqlDbType.VarChar, 40);
			cmBoletoInsert.Parameters.Add("@cep_sacado", SqlDbType.VarChar, 8);
			cmBoletoInsert.Parameters.Add("@bairro_sacado", SqlDbType.VarChar, 72);
			cmBoletoInsert.Parameters.Add("@cidade_sacado", SqlDbType.VarChar, 60);
			cmBoletoInsert.Parameters.Add("@uf_sacado", SqlDbType.VarChar, 2);
			cmBoletoInsert.Parameters.Add("@email_sacado", SqlDbType.VarChar, 512);
			cmBoletoInsert.Parameters.Add("@segunda_mensagem", SqlDbType.VarChar, 60);
			cmBoletoInsert.Parameters.Add("@mensagem_1", SqlDbType.VarChar, 80);
			cmBoletoInsert.Parameters.Add("@mensagem_2", SqlDbType.VarChar, 80);
			cmBoletoInsert.Parameters.Add("@mensagem_3", SqlDbType.VarChar, 80);
			cmBoletoInsert.Parameters.Add("@mensagem_4", SqlDbType.VarChar, 80);
			cmBoletoInsert.Parameters.Add("@usuario_cadastro", SqlDbType.VarChar, 10);
			cmBoletoInsert.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoInsert.Prepare();
			#endregion

			#region [ cmBoletoItemInsert ]
			strSql = "INSERT INTO t_FIN_BOLETO_ITEM (" +
						"id, " +
						"id_boleto, " +
						"num_parcela, " +
						"status, " +
						"tipo_vencimento, " +
						"dt_vencto, " +
						"valor, " +
						"bonificacao_por_dia, " +
						"valor_por_dia_atraso, " +
						"dt_limite_desconto, " +
						"valor_desconto, " +
						"numero_documento, " +
						"primeira_mensagem, " +
						"num_controle_participante, " +
						"usuario_ult_atualizacao, " +
						"st_instrucao_protesto" +
					") VALUES (" +
						"@id, " +
						"@id_boleto, " +
						"@num_parcela, " +
						"@status, " +
						"@tipo_vencimento, " +
						"@dt_vencto, " +
						"@valor, " +
						"@bonificacao_por_dia, " +
						"@valor_por_dia_atraso, " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_limite_desconto") + ", " +
						"@valor_desconto, " +
						"@numero_documento, " +
						"@primeira_mensagem, " +
						"@num_controle_participante, " +
						"@usuario_ult_atualizacao, " +
						"@st_instrucao_protesto" +
					")";
			cmBoletoItemInsert = BD.criaSqlCommand();
			cmBoletoItemInsert.CommandText = strSql;
			cmBoletoItemInsert.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemInsert.Parameters.Add("@id_boleto", SqlDbType.Int);
			cmBoletoItemInsert.Parameters.Add("@num_parcela", SqlDbType.TinyInt);
			cmBoletoItemInsert.Parameters.Add("@status", SqlDbType.SmallInt);
			cmBoletoItemInsert.Parameters.Add("@tipo_vencimento", SqlDbType.TinyInt);
			cmBoletoItemInsert.Parameters.Add("@dt_vencto", SqlDbType.VarChar, 10);
			cmBoletoItemInsert.Parameters.Add("@valor", SqlDbType.Money);
			cmBoletoItemInsert.Parameters.Add("@bonificacao_por_dia", SqlDbType.Money);
			cmBoletoItemInsert.Parameters.Add("@valor_por_dia_atraso", SqlDbType.Money);
			cmBoletoItemInsert.Parameters.Add("@dt_limite_desconto", SqlDbType.VarChar, 10);
			cmBoletoItemInsert.Parameters.Add("@valor_desconto", SqlDbType.Money);
			cmBoletoItemInsert.Parameters.Add("@numero_documento", SqlDbType.VarChar, 10);
			cmBoletoItemInsert.Parameters.Add("@primeira_mensagem", SqlDbType.VarChar, 12);
			cmBoletoItemInsert.Parameters.Add("@num_controle_participante", SqlDbType.VarChar, 25);
			cmBoletoItemInsert.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemInsert.Parameters.Add("@st_instrucao_protesto", SqlDbType.TinyInt);
			cmBoletoItemInsert.Prepare();
			#endregion

			#region [ cmBoletoItemRateioInsert ]
			strSql = "INSERT INTO t_FIN_BOLETO_ITEM_RATEIO (" +
						"id_boleto_item, " +
						"pedido, " +
						"id_boleto, " +
						"valor" +
					") VALUES (" +
						"@id_boleto_item, " +
						"@pedido, " +
						"@id_boleto, " +
						"@valor" +
					")";
			cmBoletoItemRateioInsert = BD.criaSqlCommand();
			cmBoletoItemRateioInsert.CommandText = strSql;
			cmBoletoItemRateioInsert.Parameters.Add("@id_boleto_item", SqlDbType.Int);
			cmBoletoItemRateioInsert.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmBoletoItemRateioInsert.Parameters.Add("@id_boleto", SqlDbType.Int);
			cmBoletoItemRateioInsert.Parameters.Add("@valor", SqlDbType.Money);
			cmBoletoItemRateioInsert.Prepare();
			#endregion

			#region [ cmBoletoMarcaEnviadoRemessaBanco ]
			strSql = "UPDATE t_FIN_BOLETO SET " +
						"status = " + Global.Cte.FIN.CodBoletoStatus.ENVIADO_REMESSA_BANCO.ToString() + ", " +
						"dt_remessa = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"id_boleto_arq_remessa = @id_boleto_arq_remessa, " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id) " +
						"AND (status = " + Global.Cte.FIN.CodBoletoStatus.INICIAL.ToString() + ")";
			cmBoletoMarcaEnviadoRemessaBanco = BD.criaSqlCommand();
			cmBoletoMarcaEnviadoRemessaBanco.CommandText = strSql;
			cmBoletoMarcaEnviadoRemessaBanco.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoMarcaEnviadoRemessaBanco.Parameters.Add("@id_boleto_arq_remessa", SqlDbType.Int);
			cmBoletoMarcaEnviadoRemessaBanco.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoMarcaEnviadoRemessaBanco.Prepare();
			#endregion

			#region [ cmBoletoItemMarcaEnviadoRemessaBanco ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"status = " + Global.Cte.FIN.CodBoletoItemStatus.ENVIADO_REMESSA_BANCO.ToString() + ", " +
						"dt_emissao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id) " +
						"AND (status = " + Global.Cte.FIN.CodBoletoItemStatus.INICIAL.ToString() + ")";
			cmBoletoItemMarcaEnviadoRemessaBanco = BD.criaSqlCommand();
			cmBoletoItemMarcaEnviadoRemessaBanco.CommandText = strSql;
			cmBoletoItemMarcaEnviadoRemessaBanco.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemMarcaEnviadoRemessaBanco.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemMarcaEnviadoRemessaBanco.Prepare();
			#endregion

			#region [ cmBoletoMarcaCanceladoManual ]
			strSql = "UPDATE t_FIN_BOLETO SET " +
						"status = " + Global.Cte.FIN.CodBoletoStatus.CANCELADO_MANUAL.ToString() + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id) " +
						"AND (status = " + Global.Cte.FIN.CodBoletoStatus.INICIAL.ToString() + ")";
			cmBoletoMarcaCanceladoManual = BD.criaSqlCommand();
			cmBoletoMarcaCanceladoManual.CommandText = strSql;
			cmBoletoMarcaCanceladoManual.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoMarcaCanceladoManual.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoMarcaCanceladoManual.Prepare();
			#endregion

			#region [ cmBoletoItemMarcaCanceladoManual ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"status = " + Global.Cte.FIN.CodBoletoItemStatus.CANCELADO_MANUAL.ToString() + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id) " +
						"AND (status = " + Global.Cte.FIN.CodBoletoItemStatus.INICIAL.ToString() + ")";
			cmBoletoItemMarcaCanceladoManual = BD.criaSqlCommand();
			cmBoletoItemMarcaCanceladoManual.CommandText = strSql;
			cmBoletoItemMarcaCanceladoManual.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemMarcaCanceladoManual.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemMarcaCanceladoManual.Prepare();
			#endregion

			#region [ cmBoletoItemMarcaCanceladoManualByIdBoleto ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"status = " + Global.Cte.FIN.CodBoletoItemStatus.CANCELADO_MANUAL.ToString() + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id_boleto = @id_boleto) " +
						"AND (status = " + Global.Cte.FIN.CodBoletoItemStatus.INICIAL.ToString() + ")";
			cmBoletoItemMarcaCanceladoManualByIdBoleto = BD.criaSqlCommand();
			cmBoletoItemMarcaCanceladoManualByIdBoleto.CommandText = strSql;
			cmBoletoItemMarcaCanceladoManualByIdBoleto.Parameters.Add("@id_boleto", SqlDbType.Int);
			cmBoletoItemMarcaCanceladoManualByIdBoleto.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemMarcaCanceladoManualByIdBoleto.Prepare();
			#endregion

			#region [ cmBoletoArqRemessaInsert ]
			strSql = "INSERT INTO t_FIN_BOLETO_ARQ_REMESSA (" +
						"id, " +
						"nsu_arq_remessa, " +
						"usuario_geracao, " +
						"qtde_registros, " +
						"qtde_serie_boletos, " +
						"id_boleto_cedente, " +
						"codigo_empresa, " +
						"nome_empresa, " +
						"num_banco, " +
						"nome_banco, " +
						"agencia, " +
						"digito_agencia, " +
						"conta, " +
						"digito_conta, " +
						"carteira, " +
						"vl_total, " +
						"duracao_proc_em_seg, " +
						"nome_arq_remessa, " +
						"caminho_arq_remessa, " +
						"st_geracao, " +
						"msg_erro_geracao" +
					") VALUES (" +
						"@id, " +
						"@nsu_arq_remessa, " +
						"@usuario_geracao, " +
						"@qtde_registros, " +
						"@qtde_serie_boletos, " +
						"@id_boleto_cedente, " +
						"@codigo_empresa, " +
						"@nome_empresa, " +
						"@num_banco, " +
						"@nome_banco, " +
						"@agencia, " +
						"@digito_agencia, " +
						"@conta, " +
						"@digito_conta, " +
						"@carteira, " +
						"@vl_total, " +
						"@duracao_proc_em_seg, " +
						"@nome_arq_remessa, " +
						"@caminho_arq_remessa, " +
						"@st_geracao, " +
						"@msg_erro_geracao" +
					")";
			cmBoletoArqRemessaInsert = BD.criaSqlCommand();
			cmBoletoArqRemessaInsert.CommandText = strSql;
			cmBoletoArqRemessaInsert.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoArqRemessaInsert.Parameters.Add("@nsu_arq_remessa", SqlDbType.Int);
			cmBoletoArqRemessaInsert.Parameters.Add("@usuario_geracao", SqlDbType.VarChar, 10);
			cmBoletoArqRemessaInsert.Parameters.Add("@qtde_registros", SqlDbType.Int);
			cmBoletoArqRemessaInsert.Parameters.Add("@qtde_serie_boletos", SqlDbType.Int);
			cmBoletoArqRemessaInsert.Parameters.Add("@id_boleto_cedente", SqlDbType.SmallInt);
			cmBoletoArqRemessaInsert.Parameters.Add("@codigo_empresa", SqlDbType.VarChar, 20);
			cmBoletoArqRemessaInsert.Parameters.Add("@nome_empresa", SqlDbType.VarChar, 30);
			cmBoletoArqRemessaInsert.Parameters.Add("@num_banco", SqlDbType.VarChar, 3);
			cmBoletoArqRemessaInsert.Parameters.Add("@nome_banco", SqlDbType.VarChar, 15);
			cmBoletoArqRemessaInsert.Parameters.Add("@agencia", SqlDbType.VarChar, 5);
			cmBoletoArqRemessaInsert.Parameters.Add("@digito_agencia", SqlDbType.VarChar, 1);
			cmBoletoArqRemessaInsert.Parameters.Add("@conta", SqlDbType.VarChar, 7);
			cmBoletoArqRemessaInsert.Parameters.Add("@digito_conta", SqlDbType.VarChar, 1);
			cmBoletoArqRemessaInsert.Parameters.Add("@carteira", SqlDbType.VarChar, 3);
			cmBoletoArqRemessaInsert.Parameters.Add("@vl_total", SqlDbType.Money);
			cmBoletoArqRemessaInsert.Parameters.Add("@duracao_proc_em_seg", SqlDbType.Int);
			cmBoletoArqRemessaInsert.Parameters.Add("@nome_arq_remessa", SqlDbType.VarChar, 40);
			cmBoletoArqRemessaInsert.Parameters.Add("@caminho_arq_remessa", SqlDbType.VarChar, 1024);
			cmBoletoArqRemessaInsert.Parameters.Add("@st_geracao", SqlDbType.SmallInt);
			cmBoletoArqRemessaInsert.Parameters.Add("@msg_erro_geracao", SqlDbType.VarChar, 1024);
			cmBoletoArqRemessaInsert.Prepare();
			#endregion

			#region [ cmBoletoArqRetornoInsert ]
			strSql = "INSERT INTO t_FIN_BOLETO_ARQ_RETORNO (" +
						"id, " +
						"id_boleto_cedente, " +
						"usuario_processamento, " +
						"qtde_registros, " +
						"codigo_empresa, " +
						"nome_empresa, " +
						"num_banco, " +
						"nome_banco, " +
						"data_gravacao_arquivo, " +
						"dt_gravacao_arquivo, " +
						"numero_aviso_bancario, " +
						"data_credito, " +
						"dt_credito, " +
						"qtdeTitulosEmCobranca, " +
						"valorTotalEmCobranca, " +
						"qtdeRegsOcorrencia02ConfirmacaoEntradas, " +
						"valorRegsOcorrencia02ConfirmacaoEntradas, " +
						"valorRegsOcorrencia06Liquidacao, " +
						"qtdeRegsOcorrencia06Liquidacao, " +
						"valorRegsOcorrencia06, " +
						"qtdeRegsOcorrencia09e10TitulosBaixados, " +
						"valorRegsOcorrencia09e10TitulosBaixados, " +
						"qtdeRegsOcorrencia13AbatimentoCancelado, " +
						"valorRegsOcorrencia13AbatimentoCancelado, " +
						"qtdeRegsOcorrencia14VenctoAlterado, " +
						"valorRegsOcorrencia14VenctoAlterado, " +
						"qtdeRegsOcorrencia12AbatimentoConcedido, " +
						"valorRegsOcorrencia12AbatimentoConcedido, " +
						"qtdeRegsOcorrencia19ConfirmacaoInstrucaoProtesto, " +
						"valorRegsOcorrencia19ConfirmacaoInstrucaoProtesto, " +
						"valorTotalRateiosEfetuados, " +
						"qtdeTotalRateiosEfetuados, " +
						"duracao_proc_em_seg, " +
						"nome_arq_retorno, " +
						"caminho_arq_retorno, " +
						"st_processamento, " +
						"msg_erro_processamento" +
					") VALUES (" +
						"@id, " +
						"@id_boleto_cedente, " +
						"@usuario_processamento, " +
						"@qtde_registros, " +
						"@codigo_empresa, " +
						"@nome_empresa, " +
						"@num_banco, " +
						"@nome_banco, " +
						"@data_gravacao_arquivo, " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_gravacao_arquivo") + ", " +
						"@numero_aviso_bancario, " +
						"@data_credito, " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_credito") + ", " +
						"@qtdeTitulosEmCobranca, " +
						"@valorTotalEmCobranca, " +
						"@qtdeRegsOcorrencia02ConfirmacaoEntradas, " +
						"@valorRegsOcorrencia02ConfirmacaoEntradas, " +
						"@valorRegsOcorrencia06Liquidacao, " +
						"@qtdeRegsOcorrencia06Liquidacao, " +
						"@valorRegsOcorrencia06, " +
						"@qtdeRegsOcorrencia09e10TitulosBaixados, " +
						"@valorRegsOcorrencia09e10TitulosBaixados, " +
						"@qtdeRegsOcorrencia13AbatimentoCancelado, " +
						"@valorRegsOcorrencia13AbatimentoCancelado, " +
						"@qtdeRegsOcorrencia14VenctoAlterado, " +
						"@valorRegsOcorrencia14VenctoAlterado, " +
						"@qtdeRegsOcorrencia12AbatimentoConcedido, " +
						"@valorRegsOcorrencia12AbatimentoConcedido, " +
						"@qtdeRegsOcorrencia19ConfirmacaoInstrucaoProtesto, " +
						"@valorRegsOcorrencia19ConfirmacaoInstrucaoProtesto, " +
						"@valorTotalRateiosEfetuados, " +
						"@qtdeTotalRateiosEfetuados, " +
						"@duracao_proc_em_seg, " +
						"@nome_arq_retorno, " +
						"@caminho_arq_retorno, " +
						"@st_processamento, " +
						"@msg_erro_processamento" +
					")";
			cmBoletoArqRetornoInsert = BD.criaSqlCommand();
			cmBoletoArqRetornoInsert.CommandText = strSql;
			cmBoletoArqRetornoInsert.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoArqRetornoInsert.Parameters.Add("@id_boleto_cedente", SqlDbType.SmallInt);
			cmBoletoArqRetornoInsert.Parameters.Add("@usuario_processamento", SqlDbType.VarChar, 10);
			cmBoletoArqRetornoInsert.Parameters.Add("@qtde_registros", SqlDbType.Int);
			cmBoletoArqRetornoInsert.Parameters.Add("@codigo_empresa", SqlDbType.VarChar, 20);
			cmBoletoArqRetornoInsert.Parameters.Add("@nome_empresa", SqlDbType.VarChar, 30);
			cmBoletoArqRetornoInsert.Parameters.Add("@num_banco", SqlDbType.VarChar, 3);
			cmBoletoArqRetornoInsert.Parameters.Add("@nome_banco", SqlDbType.VarChar, 15);
			cmBoletoArqRetornoInsert.Parameters.Add("@data_gravacao_arquivo", SqlDbType.VarChar, 6);
			cmBoletoArqRetornoInsert.Parameters.Add("@dt_gravacao_arquivo", SqlDbType.VarChar, 10);
			cmBoletoArqRetornoInsert.Parameters.Add("@numero_aviso_bancario", SqlDbType.VarChar, 5);
			cmBoletoArqRetornoInsert.Parameters.Add("@data_credito", SqlDbType.VarChar, 6);
			cmBoletoArqRetornoInsert.Parameters.Add("@dt_credito", SqlDbType.VarChar, 10);
			cmBoletoArqRetornoInsert.Parameters.Add("@qtdeTitulosEmCobranca", SqlDbType.VarChar, 8);
			cmBoletoArqRetornoInsert.Parameters.Add("@valorTotalEmCobranca", SqlDbType.VarChar, 14);
			cmBoletoArqRetornoInsert.Parameters.Add("@qtdeRegsOcorrencia02ConfirmacaoEntradas", SqlDbType.VarChar, 5);
			cmBoletoArqRetornoInsert.Parameters.Add("@valorRegsOcorrencia02ConfirmacaoEntradas", SqlDbType.VarChar, 12);
			cmBoletoArqRetornoInsert.Parameters.Add("@valorRegsOcorrencia06Liquidacao", SqlDbType.VarChar, 12);
			cmBoletoArqRetornoInsert.Parameters.Add("@qtdeRegsOcorrencia06Liquidacao", SqlDbType.VarChar, 5);
			cmBoletoArqRetornoInsert.Parameters.Add("@valorRegsOcorrencia06", SqlDbType.VarChar, 12);
			cmBoletoArqRetornoInsert.Parameters.Add("@qtdeRegsOcorrencia09e10TitulosBaixados", SqlDbType.VarChar, 5);
			cmBoletoArqRetornoInsert.Parameters.Add("@valorRegsOcorrencia09e10TitulosBaixados", SqlDbType.VarChar, 12);
			cmBoletoArqRetornoInsert.Parameters.Add("@qtdeRegsOcorrencia13AbatimentoCancelado", SqlDbType.VarChar, 5);
			cmBoletoArqRetornoInsert.Parameters.Add("@valorRegsOcorrencia13AbatimentoCancelado", SqlDbType.VarChar, 12);
			cmBoletoArqRetornoInsert.Parameters.Add("@qtdeRegsOcorrencia14VenctoAlterado", SqlDbType.VarChar, 5);
			cmBoletoArqRetornoInsert.Parameters.Add("@valorRegsOcorrencia14VenctoAlterado", SqlDbType.VarChar, 12);
			cmBoletoArqRetornoInsert.Parameters.Add("@qtdeRegsOcorrencia12AbatimentoConcedido", SqlDbType.VarChar, 5);
			cmBoletoArqRetornoInsert.Parameters.Add("@valorRegsOcorrencia12AbatimentoConcedido", SqlDbType.VarChar, 12);
			cmBoletoArqRetornoInsert.Parameters.Add("@qtdeRegsOcorrencia19ConfirmacaoInstrucaoProtesto", SqlDbType.VarChar, 5);
			cmBoletoArqRetornoInsert.Parameters.Add("@valorRegsOcorrencia19ConfirmacaoInstrucaoProtesto", SqlDbType.VarChar, 12);
			cmBoletoArqRetornoInsert.Parameters.Add("@valorTotalRateiosEfetuados", SqlDbType.VarChar, 15);
			cmBoletoArqRetornoInsert.Parameters.Add("@qtdeTotalRateiosEfetuados", SqlDbType.VarChar, 8);
			cmBoletoArqRetornoInsert.Parameters.Add("@duracao_proc_em_seg", SqlDbType.Int);
			cmBoletoArqRetornoInsert.Parameters.Add("@nome_arq_retorno", SqlDbType.VarChar, 40);
			cmBoletoArqRetornoInsert.Parameters.Add("@caminho_arq_retorno", SqlDbType.VarChar, 1024);
			cmBoletoArqRetornoInsert.Parameters.Add("@st_processamento", SqlDbType.SmallInt);
			cmBoletoArqRetornoInsert.Parameters.Add("@msg_erro_processamento", SqlDbType.VarChar, 1024);
			cmBoletoArqRetornoInsert.Prepare();
			#endregion

			#region [ cmBoletoArqRetornoUpdate ]
			strSql = "UPDATE t_FIN_BOLETO_ARQ_RETORNO SET " +
						"st_processamento = @st_processamento, " +
						"duracao_proc_em_seg = @duracao_proc_em_seg, " +
						"msg_erro_processamento = @msg_erro_processamento " +
					"WHERE " +
						"(id = @id)";
			cmBoletoArqRetornoUpdate = BD.criaSqlCommand();
			cmBoletoArqRetornoUpdate.CommandText = strSql;
			cmBoletoArqRetornoUpdate.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoArqRetornoUpdate.Parameters.Add("@st_processamento", SqlDbType.SmallInt);
			cmBoletoArqRetornoUpdate.Parameters.Add("@duracao_proc_em_seg", SqlDbType.Int);
			cmBoletoArqRetornoUpdate.Parameters.Add("@msg_erro_processamento", SqlDbType.VarChar, 1024);
			cmBoletoArqRetornoUpdate.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrencia02 ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"status = " + Global.Cte.FIN.CodBoletoItemStatus.ENTRADA_CONFIRMADA.ToString() + ", " +
						"nosso_numero = @nosso_numero, " +
						"digito_nosso_numero = @digito_nosso_numero, " +
						"codigo_barras = @codigo_barras, " +
						"linha_digitavel = @linha_digitavel, " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_motivo_ocorrencia_19 = @ult_motivo_ocorrencia_19, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"dt_entrada_confirmada = @dt_entrada_confirmada, " +
						"vl_tarifa_registro = @vl_tarifa_registro, " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrencia02 = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrencia02.CommandText = strSql;
			cmBoletoItemAtualizaOcorrencia02.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrencia02.Parameters.Add("@nosso_numero", SqlDbType.VarChar, 11);
			cmBoletoItemAtualizaOcorrencia02.Parameters.Add("@digito_nosso_numero", SqlDbType.VarChar, 1);
			cmBoletoItemAtualizaOcorrencia02.Parameters.Add("@codigo_barras", SqlDbType.VarChar, 44);
			cmBoletoItemAtualizaOcorrencia02.Parameters.Add("@linha_digitavel", SqlDbType.VarChar, 54);
			cmBoletoItemAtualizaOcorrencia02.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrencia02.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia02.Parameters.Add("@ult_motivo_ocorrencia_19", SqlDbType.VarChar, 1);
			cmBoletoItemAtualizaOcorrencia02.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia02.Parameters.Add("@dt_entrada_confirmada", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia02.Parameters.Add("@vl_tarifa_registro", SqlDbType.Money);
			cmBoletoItemAtualizaOcorrencia02.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia02.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrencia06 ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"status = " + Global.Cte.FIN.CodBoletoItemStatus.BOLETO_PAGO.ToString() + ", " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"vl_abatimento_concedido = @vl_abatimento_concedido, " +
						"vl_desconto_concedido = @vl_desconto_concedido, " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"st_boleto_ocorrencia_06 = @st_boleto_ocorrencia_06, " +
						"dt_ocorrencia_banco_boleto_ocorrencia_06 = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_ocorrencia_06") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrencia06 = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrencia06.CommandText = strSql;
			cmBoletoItemAtualizaOcorrencia06.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrencia06.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrencia06.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia06.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia06.Parameters.Add("@vl_abatimento_concedido", SqlDbType.Money);
			cmBoletoItemAtualizaOcorrencia06.Parameters.Add("@vl_desconto_concedido", SqlDbType.Money);
			cmBoletoItemAtualizaOcorrencia06.Parameters.Add("@st_boleto_ocorrencia_06", SqlDbType.TinyInt);
			cmBoletoItemAtualizaOcorrencia06.Parameters.Add("@dt_ocorrencia_banco_boleto_ocorrencia_06", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia06.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia06.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrencia09 ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"status = " + Global.Cte.FIN.CodBoletoItemStatus.BOLETO_BAIXADO.ToString() + ", " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"st_boleto_baixado = @st_boleto_baixado, " +
						"dt_ocorrencia_banco_boleto_baixado = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_baixado") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrencia09 = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrencia09.CommandText = strSql;
			cmBoletoItemAtualizaOcorrencia09.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrencia09.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrencia09.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia09.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia09.Parameters.Add("@st_boleto_baixado", SqlDbType.TinyInt);
			cmBoletoItemAtualizaOcorrencia09.Parameters.Add("@dt_ocorrencia_banco_boleto_baixado", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia09.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia09.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrencia10 ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"status = " + Global.Cte.FIN.CodBoletoItemStatus.BOLETO_BAIXADO.ToString() + ", " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"st_boleto_baixado = @st_boleto_baixado, " +
						"dt_ocorrencia_banco_boleto_baixado = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_baixado") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrencia10 = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrencia10.CommandText = strSql;
			cmBoletoItemAtualizaOcorrencia10.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrencia10.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrencia10.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia10.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia10.Parameters.Add("@st_boleto_baixado", SqlDbType.TinyInt);
			cmBoletoItemAtualizaOcorrencia10.Parameters.Add("@dt_ocorrencia_banco_boleto_baixado", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia10.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia10.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrencia12 ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"vl_abatimento_concedido = @vl_abatimento_concedido, " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrencia12 = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrencia12.CommandText = strSql;
			cmBoletoItemAtualizaOcorrencia12.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrencia12.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrencia12.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia12.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia12.Parameters.Add("@vl_abatimento_concedido", SqlDbType.Money);
			cmBoletoItemAtualizaOcorrencia12.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia12.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrencia13 ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"vl_abatimento_concedido = @vl_abatimento_concedido, " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrencia13 = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrencia13.CommandText = strSql;
			cmBoletoItemAtualizaOcorrencia13.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrencia13.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrencia13.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia13.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia13.Parameters.Add("@vl_abatimento_concedido", SqlDbType.Money);
			cmBoletoItemAtualizaOcorrencia13.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia13.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrencia14 ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"dt_vencto = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_vencto") + ", " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrencia14 = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrencia14.CommandText = strSql;
			cmBoletoItemAtualizaOcorrencia14.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrencia14.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrencia14.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia14.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia14.Parameters.Add("@dt_vencto", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia14.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia14.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrencia15 ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"status = " + Global.Cte.FIN.CodBoletoItemStatus.BOLETO_PAGO_EM_OCORRENCIA_15.ToString() + ", " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"vl_abatimento_concedido = @vl_abatimento_concedido, " +
						"vl_desconto_concedido = @vl_desconto_concedido, " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"st_boleto_ocorrencia_15 = @st_boleto_ocorrencia_15, " +
						"dt_ocorrencia_banco_boleto_ocorrencia_15 = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_ocorrencia_15") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrencia15 = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrencia15.CommandText = strSql;
			cmBoletoItemAtualizaOcorrencia15.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrencia15.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrencia15.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia15.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia15.Parameters.Add("@vl_abatimento_concedido", SqlDbType.Money);
			cmBoletoItemAtualizaOcorrencia15.Parameters.Add("@vl_desconto_concedido", SqlDbType.Money);
			cmBoletoItemAtualizaOcorrencia15.Parameters.Add("@st_boleto_ocorrencia_15", SqlDbType.TinyInt);
			cmBoletoItemAtualizaOcorrencia15.Parameters.Add("@dt_ocorrencia_banco_boleto_ocorrencia_15", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia15.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia15.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrencia16 ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"status = " + Global.Cte.FIN.CodBoletoItemStatus.BOLETO_PAGO_EM_CHEQUE.ToString() + ", " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"st_boleto_pago_cheque = @st_boleto_pago_cheque, " +
						"dt_ocorrencia_banco_boleto_pago_cheque = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_pago_cheque") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrencia16 = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrencia16.CommandText = strSql;
			cmBoletoItemAtualizaOcorrencia16.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrencia16.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrencia16.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia16.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia16.Parameters.Add("@st_boleto_pago_cheque", SqlDbType.TinyInt);
			cmBoletoItemAtualizaOcorrencia16.Parameters.Add("@dt_ocorrencia_banco_boleto_pago_cheque", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia16.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia16.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrencia17 ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"status = " + Global.Cte.FIN.CodBoletoItemStatus.BOLETO_PAGO_EM_OCORRENCIA_17.ToString() + ", " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"st_boleto_ocorrencia_17 = @st_boleto_ocorrencia_17, " +
						"dt_ocorrencia_banco_boleto_ocorrencia_17 = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_ocorrencia_17") + ", " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrencia17 = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrencia17.CommandText = strSql;
			cmBoletoItemAtualizaOcorrencia17.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrencia17.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrencia17.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia17.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia17.Parameters.Add("@st_boleto_ocorrencia_17", SqlDbType.TinyInt);
			cmBoletoItemAtualizaOcorrencia17.Parameters.Add("@dt_ocorrencia_banco_boleto_ocorrencia_17", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia17.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia17.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrencia19 ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_motivo_ocorrencia_19 = @ult_motivo_ocorrencia_19, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrencia19 = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrencia19.CommandText = strSql;
			cmBoletoItemAtualizaOcorrencia19.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrencia19.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrencia19.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia19.Parameters.Add("@ult_motivo_ocorrencia_19", SqlDbType.VarChar, 1);
			cmBoletoItemAtualizaOcorrencia19.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia19.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia19.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrencia22 ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"status = " + Global.Cte.FIN.CodBoletoItemStatus.BOLETO_COM_PAGAMENTO_CANCELADO.ToString() + ", " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"st_boleto_pago_cheque = @st_boleto_pago_cheque, " +
						"dt_ocorrencia_banco_boleto_pago_cheque = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_pago_cheque") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrencia22 = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrencia22.CommandText = strSql;
			cmBoletoItemAtualizaOcorrencia22.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrencia22.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrencia22.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia22.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia22.Parameters.Add("@st_boleto_pago_cheque", SqlDbType.TinyInt);
			cmBoletoItemAtualizaOcorrencia22.Parameters.Add("@dt_ocorrencia_banco_boleto_pago_cheque", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia22.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia22.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrencia23 ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"st_boleto_ocorrencia_23 = @st_boleto_ocorrencia_23, " +
						"dt_ocorrencia_banco_boleto_ocorrencia_23 = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_ocorrencia_23") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrencia23 = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrencia23.CommandText = strSql;
			cmBoletoItemAtualizaOcorrencia23.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrencia23.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrencia23.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia23.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia23.Parameters.Add("@st_boleto_ocorrencia_23", SqlDbType.TinyInt);
			cmBoletoItemAtualizaOcorrencia23.Parameters.Add("@dt_ocorrencia_banco_boleto_ocorrencia_23", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia23.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia23.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrenciaValaComum ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"status = " + Global.Cte.FIN.CodBoletoItemStatus.VALA_COMUM.ToString() + ", " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_motivo_ocorrencia_19 = @ult_motivo_ocorrencia_19, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrenciaValaComum = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrenciaValaComum.CommandText = strSql;
			cmBoletoItemAtualizaOcorrenciaValaComum.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrenciaValaComum.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrenciaValaComum.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrenciaValaComum.Parameters.Add("@ult_motivo_ocorrencia_19", SqlDbType.VarChar, 1);
			cmBoletoItemAtualizaOcorrenciaValaComum.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrenciaValaComum.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrenciaValaComum.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrencia24 ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"status = " + Global.Cte.FIN.CodBoletoItemStatus.BOLETO_REJEITADO_CEP_IRREGULAR.ToString() + ", " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_motivo_ocorrencia_19 = @ult_motivo_ocorrencia_19, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrencia24 = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrencia24.CommandText = strSql;
			cmBoletoItemAtualizaOcorrencia24.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrencia24.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrencia24.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia24.Parameters.Add("@ult_motivo_ocorrencia_19", SqlDbType.VarChar, 1);
			cmBoletoItemAtualizaOcorrencia24.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia24.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia24.Prepare();
			#endregion

			#region [ cmBoletoItemAtualizaOcorrencia34 ]
			strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
						"ult_identificacao_ocorrencia = @ult_identificacao_ocorrencia, " +
						"ult_motivos_rejeicoes = @ult_motivos_rejeicoes, " +
						"ult_data_ocorrencia_banco = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ult_data_ocorrencia_banco") + ", " +
						"ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"st_boleto_ocorrencia_34 = @st_boleto_ocorrencia_34, " +
						"dt_ocorrencia_banco_boleto_ocorrencia_34 = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_ocorrencia_34") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoItemAtualizaOcorrencia34 = BD.criaSqlCommand();
			cmBoletoItemAtualizaOcorrencia34.CommandText = strSql;
			cmBoletoItemAtualizaOcorrencia34.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoItemAtualizaOcorrencia34.Parameters.Add("@ult_identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoItemAtualizaOcorrencia34.Parameters.Add("@ult_motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia34.Parameters.Add("@ult_data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia34.Parameters.Add("@st_boleto_ocorrencia_34", SqlDbType.TinyInt);
			cmBoletoItemAtualizaOcorrencia34.Parameters.Add("@dt_ocorrencia_banco_boleto_ocorrencia_34", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia34.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoItemAtualizaOcorrencia34.Prepare();
			#endregion

			#region [ cmBoletoMovimentoInsert ]
			strSql = "INSERT INTO t_FIN_BOLETO_MOVIMENTO (" +
						"id, " +
						"id_arq_retorno, " +
						"id_boleto, " +
						"id_boleto_item, " +
						"identificacao_ocorrencia, " +
						"motivos_rejeicoes, " +
						"data_ocorrencia_banco, " +
						"numero_documento, " +
						"nosso_numero, " +
						"digito_nosso_numero, " +
						"dt_vencto, " +
						"vl_titulo, " +
						"vl_despesas_cobranca, " +
						"vl_outras_despesas, " +
						"vl_IOF, " +
						"vl_abatimento, " +
						"vl_desconto, " +
						"vl_pago, " +
						"vl_juros_mora, " +
						"dt_credito" +
					") VALUES (" +
						"@id, " +
						"@id_arq_retorno, " +
						"@id_boleto, " +
						"@id_boleto_item, " +
						"@identificacao_ocorrencia, " +
						"@motivos_rejeicoes, " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@data_ocorrencia_banco") + ", " +
						"@numero_documento, " +
						"@nosso_numero, " +
						"@digito_nosso_numero, " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_vencto") + ", " +
						"@vl_titulo, " +
						"@vl_despesas_cobranca, " +
						"@vl_outras_despesas, " +
						"@vl_IOF, " +
						"@vl_abatimento, " +
						"@vl_desconto, " +
						"@vl_pago, " +
						"@vl_juros_mora, " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_credito") +
					")";
			cmBoletoMovimentoInsert = BD.criaSqlCommand();
			cmBoletoMovimentoInsert.CommandText = strSql;
			cmBoletoMovimentoInsert.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoMovimentoInsert.Parameters.Add("@id_arq_retorno", SqlDbType.Int);
			cmBoletoMovimentoInsert.Parameters.Add("@id_boleto", SqlDbType.Int);
			cmBoletoMovimentoInsert.Parameters.Add("@id_boleto_item", SqlDbType.Int);
			cmBoletoMovimentoInsert.Parameters.Add("@identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoMovimentoInsert.Parameters.Add("@motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoMovimentoInsert.Parameters.Add("@data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoMovimentoInsert.Parameters.Add("@numero_documento", SqlDbType.VarChar, 10);
			cmBoletoMovimentoInsert.Parameters.Add("@nosso_numero", SqlDbType.VarChar, 11);
			cmBoletoMovimentoInsert.Parameters.Add("@digito_nosso_numero", SqlDbType.VarChar, 1);
			cmBoletoMovimentoInsert.Parameters.Add("@dt_vencto", SqlDbType.VarChar, 10);
			cmBoletoMovimentoInsert.Parameters.Add("@vl_titulo", SqlDbType.Money);
			cmBoletoMovimentoInsert.Parameters.Add("@vl_despesas_cobranca", SqlDbType.Money);
			cmBoletoMovimentoInsert.Parameters.Add("@vl_outras_despesas", SqlDbType.Money);
			cmBoletoMovimentoInsert.Parameters.Add("@vl_IOF", SqlDbType.Money);
			cmBoletoMovimentoInsert.Parameters.Add("@vl_abatimento", SqlDbType.Money);
			cmBoletoMovimentoInsert.Parameters.Add("@vl_desconto", SqlDbType.Money);
			cmBoletoMovimentoInsert.Parameters.Add("@vl_pago", SqlDbType.Money);
			cmBoletoMovimentoInsert.Parameters.Add("@vl_juros_mora", SqlDbType.Money);
			cmBoletoMovimentoInsert.Parameters.Add("@dt_credito", SqlDbType.VarChar, 10);
			cmBoletoMovimentoInsert.Prepare();
			#endregion

			#region [ cmBoletoOcorrenciaInsert ]
			strSql = "INSERT INTO t_FIN_BOLETO_OCORRENCIA (" +
						"id, " +
						"id_arq_retorno, " +
						"id_boleto_cedente, " +
						"id_boleto, " +
						"id_boleto_item, " +
						"st_divergencia_valor, " +
						"numero_documento, " +
						"nosso_numero, " +
						"digito_nosso_numero, " +
						"dt_vencto, " +
						"vl_titulo, " +
						"identificacao_ocorrencia, " +
						"motivos_rejeicoes, " +
						"motivo_ocorrencia_19, " +
						"data_ocorrencia_banco, " +
						"obs_ocorrencia, " +
						"registro_arq_retorno" +
					") VALUES (" +
						"@id, " +
						"@id_arq_retorno, " +
						"@id_boleto_cedente, " +
						"@id_boleto, " +
						"@id_boleto_item, " +
						"@st_divergencia_valor, " +
						"@numero_documento, " +
						"@nosso_numero, " +
						"@digito_nosso_numero, " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_vencto") + ", " +
						"@vl_titulo, " +
						"@identificacao_ocorrencia, " +
						"@motivos_rejeicoes, " +
						"@motivo_ocorrencia_19, " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@data_ocorrencia_banco") + ", " +
						"@obs_ocorrencia, " +
						"@registro_arq_retorno" +
					")";
			cmBoletoOcorrenciaInsert = BD.criaSqlCommand();
			cmBoletoOcorrenciaInsert.CommandText = strSql;
			cmBoletoOcorrenciaInsert.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoOcorrenciaInsert.Parameters.Add("@id_arq_retorno", SqlDbType.Int);
			cmBoletoOcorrenciaInsert.Parameters.Add("@id_boleto_cedente", SqlDbType.SmallInt);
			cmBoletoOcorrenciaInsert.Parameters.Add("@id_boleto", SqlDbType.Int);
			cmBoletoOcorrenciaInsert.Parameters.Add("@id_boleto_item", SqlDbType.Int);
			cmBoletoOcorrenciaInsert.Parameters.Add("@st_divergencia_valor", SqlDbType.TinyInt);
			cmBoletoOcorrenciaInsert.Parameters.Add("@numero_documento", SqlDbType.VarChar, 10);
			cmBoletoOcorrenciaInsert.Parameters.Add("@nosso_numero", SqlDbType.VarChar, 11);
			cmBoletoOcorrenciaInsert.Parameters.Add("@digito_nosso_numero", SqlDbType.VarChar, 1);
			cmBoletoOcorrenciaInsert.Parameters.Add("@dt_vencto", SqlDbType.VarChar, 10);
			cmBoletoOcorrenciaInsert.Parameters.Add("@vl_titulo", SqlDbType.Money);
			cmBoletoOcorrenciaInsert.Parameters.Add("@identificacao_ocorrencia", SqlDbType.VarChar, 2);
			cmBoletoOcorrenciaInsert.Parameters.Add("@motivos_rejeicoes", SqlDbType.VarChar, 10);
			cmBoletoOcorrenciaInsert.Parameters.Add("@motivo_ocorrencia_19", SqlDbType.VarChar, 1);
			cmBoletoOcorrenciaInsert.Parameters.Add("@data_ocorrencia_banco", SqlDbType.VarChar, 10);
			cmBoletoOcorrenciaInsert.Parameters.Add("@obs_ocorrencia", SqlDbType.VarChar, 240);
			cmBoletoOcorrenciaInsert.Parameters.Add("@registro_arq_retorno", SqlDbType.VarChar, 400);
			cmBoletoOcorrenciaInsert.Prepare();
			#endregion

			#region [ cmBoletoCorrigeOcorrencia24CepIrregular ]
			strSql = "UPDATE t_FIN_BOLETO SET " +
						"status = " + Global.Cte.FIN.CodBoletoStatus.INICIAL.ToString() + ", " +
						"endereco_sacado = @endereco_sacado, " +
						"bairro_sacado = @bairro_sacado, " +
						"cep_sacado = @cep_sacado, " +
						"cidade_sacado = @cidade_sacado, " +
						"uf_sacado = @uf_sacado, " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ult_atualizacao = getdate(), " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					"WHERE " +
						"(id = @id)";
			cmBoletoCorrigeOcorrencia24CepIrregular = BD.criaSqlCommand();
			cmBoletoCorrigeOcorrencia24CepIrregular.CommandText = strSql;
			cmBoletoCorrigeOcorrencia24CepIrregular.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoCorrigeOcorrencia24CepIrregular.Parameters.Add("@endereco_sacado", SqlDbType.VarChar, 40);
			cmBoletoCorrigeOcorrencia24CepIrregular.Parameters.Add("@bairro_sacado", SqlDbType.VarChar, 72);
			cmBoletoCorrigeOcorrencia24CepIrregular.Parameters.Add("@cep_sacado", SqlDbType.VarChar, 8);
			cmBoletoCorrigeOcorrencia24CepIrregular.Parameters.Add("@cidade_sacado", SqlDbType.VarChar, 60);
			cmBoletoCorrigeOcorrencia24CepIrregular.Parameters.Add("@uf_sacado", SqlDbType.VarChar, 2);
			cmBoletoCorrigeOcorrencia24CepIrregular.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmBoletoCorrigeOcorrencia24CepIrregular.Prepare();
			#endregion

			#region [ cmBoletoOcorrenciaMarcaComoJaTratada ]
			strSql = "UPDATE t_FIN_BOLETO_OCORRENCIA SET " +
						"st_ocorrencia_tratada = " + Global.Cte.FIN.CodBoletoOcorrenciaStOcorrenciaTratada.JA_TRATADA.ToString() + ", " +
						"comentario_ocorrencia_tratada = @comentario_ocorrencia_tratada, " +
						"dt_ocorrencia_tratada = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ocorrencia_tratada = getdate(), " +
						"usuario_ocorrencia_tratada = @usuario_ocorrencia_tratada " +
					"WHERE " +
						"(id = @id)";
			cmBoletoOcorrenciaMarcaComoJaTratada = BD.criaSqlCommand();
			cmBoletoOcorrenciaMarcaComoJaTratada.CommandText = strSql;
			cmBoletoOcorrenciaMarcaComoJaTratada.Parameters.Add("@id", SqlDbType.Int);
			cmBoletoOcorrenciaMarcaComoJaTratada.Parameters.Add("@comentario_ocorrencia_tratada", SqlDbType.VarChar, 240);
			cmBoletoOcorrenciaMarcaComoJaTratada.Parameters.Add("@usuario_ocorrencia_tratada", SqlDbType.VarChar, 10);
			cmBoletoOcorrenciaMarcaComoJaTratada.Prepare();
			#endregion

			#region [ cmBoletoOcorrenciasMarcaComoJaTratadasByIdBoleto ]
			strSql = "UPDATE t_FIN_BOLETO_OCORRENCIA SET " +
						"st_ocorrencia_tratada = " + Global.Cte.FIN.CodBoletoOcorrenciaStOcorrenciaTratada.JA_TRATADA.ToString() + ", " +
						"comentario_ocorrencia_tratada = @comentario_ocorrencia_tratada, " +
						"dt_ocorrencia_tratada = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_ocorrencia_tratada = getdate(), " +
						"usuario_ocorrencia_tratada = @usuario_ocorrencia_tratada " +
					"WHERE " +
						"(id_boleto = @id_boleto)" +
						" AND (st_ocorrencia_tratada = " + Global.Cte.FIN.CodBoletoOcorrenciaStOcorrenciaTratada.NAO_TRATADA.ToString() + ")";
			cmBoletoOcorrenciasMarcaComoJaTratadasByIdBoleto = BD.criaSqlCommand();
			cmBoletoOcorrenciasMarcaComoJaTratadasByIdBoleto.CommandText = strSql;
			cmBoletoOcorrenciasMarcaComoJaTratadasByIdBoleto.Parameters.Add("@id_boleto", SqlDbType.Int);
			cmBoletoOcorrenciasMarcaComoJaTratadasByIdBoleto.Parameters.Add("@comentario_ocorrencia_tratada", SqlDbType.VarChar, 240);
			cmBoletoOcorrenciasMarcaComoJaTratadasByIdBoleto.Parameters.Add("@usuario_ocorrencia_tratada", SqlDbType.VarChar, 10);
			cmBoletoOcorrenciasMarcaComoJaTratadasByIdBoleto.Prepare();
			#endregion
		}
		#endregion

		#region [ obtemBoletoCedenteDefinidoParaLoja ]
		/// <summary>
		/// Obtém o cedente pré-definido no sistema para a loja informada no parâmetro. Se a loja não tiver sido explicitamente alocada p/ um determinado cedente, então retorna o cedente padrão.
		/// </summary>
		/// <param name="numeroLoja">Nº da loja p/ a qual se deseja obter o cedente</param>
		/// <returns>Retorna o código de identificação do cedente</returns>
		public static int obtemBoletoCedenteDefinidoParaLoja(int numeroLoja)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Prepara objetos de acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			#endregion

			#region [ Consulta SQL ]
			strSql = "SELECT" +
						" id_boleto_cedente" +
					" FROM t_FIN_BOLETO_CEDENTE_X_LOJA" +
					" WHERE" +
						" (CONVERT(smallint, loja) = " + numeroLoja.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbResultado);
			#endregion

			#region [ Consistência ]
			if (dtbResultado.Rows.Count > 1)
			{
				throw new FinanceiroException("A loja nº " + numeroLoja.ToString().PadLeft(Global.Cte.Etc.TAM_MIN_LOJA, '0') + " possui mais do que um cedente definido!!");
			}
			#endregion

			#region [ Encontrou registro p/ a referida loja ]
			if (dtbResultado.Rows.Count == 1)
			{
				rowResultado = dtbResultado.Rows[0];
				return BD.readToInt(rowResultado["id_boleto_cedente"]);
			}
			#endregion

			#region [ A loja não possui um cedente definido, então localiza o cedente padrão ]
			strSql = "SELECT" +
						" id" +
					" FROM t_FIN_BOLETO_CEDENTE" +
					" WHERE" +
						" (st_boleto_cedente_padrao = 1)";
			cmCommand.CommandText = strSql;
			dtbResultado.Reset();
			daDataAdapter.Fill(dtbResultado);

			#region [ Consistência ]
			if (dtbResultado.Rows.Count > 1)
			{
				throw new FinanceiroException("Há mais do que um cedente padrão definido no sistema!!");
			}
			#endregion

			#region [ Localizou o cedente padrão ]
			if (dtbResultado.Rows.Count == 1)
			{
				rowResultado = dtbResultado.Rows[0];
				return BD.readToInt(rowResultado["id"]);
			}
			#endregion

			#endregion

			#region [ Não há cedente padrão definido no sistema ]
			return 0;
			#endregion
		}
		#endregion

		#region [ obtemBoletoPlanoContasDestinoByIdBoletoItem ]
		/// <summary>
		/// Dada a identificação do registro de um boleto, consulta os pedidos relacionados
		/// na tabela de rateio e localiza as lojas às quais os pedidos pertencem. 
		/// Para cada loja, obtém o plano de contas para o qual deve ser lançado o lançamento 
		/// do fluxo de caixa gerado em decorrência do boleto.
		/// Gera uma exceção no caso de não encontrar nenhum plano de contas ou se houver mais
		/// do que 1 plano de contas associado a um único boleto.
		/// </summary>
		/// <param name="idBoletoItem">Nº identificação do registro do boleto</param>
		/// <returns>Retorna um objeto do tipo BoletoPlanoContasDestino com os dados do plano de contas</returns>
		public static BoletoPlanoContasDestino obtemBoletoPlanoContasDestinoByIdBoletoItem(int idBoletoItem)
		{
			#region [ Declarações ]
			BoletoPlanoContasDestino boletoPlanoContasDestino = new BoletoPlanoContasDestino();
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Prepara objetos de acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			#endregion

			#region [ Consulta SQL ]
			strSql = "SELECT DISTINCT" +
						" id_plano_contas_empresa," +
						" id_plano_contas_grupo," +
						" id_plano_contas_conta," +
						" natureza" +
					" FROM t_FIN_BOLETO_ITEM_RATEIO tBIR" +
						" INNER JOIN t_PEDIDO tP ON (tBIR.pedido=tP.pedido)" +
						" INNER JOIN t_LOJA tL ON (tP.loja=tL.loja)" +
					" WHERE" +
						" (id_boleto_item = " + idBoletoItem.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbResultado);
			#endregion

			#region [ Consistência ]
			if (dtbResultado.Rows.Count == 0)
			{
				throw new FinanceiroException("Não há informações do plano de contas para o boleto id=" + idBoletoItem.ToString());
			}
			else if (dtbResultado.Rows.Count > 1)
			{
				throw new FinanceiroException("Há mais de 1 plano de contas associado ao boleto id=" + idBoletoItem.ToString());
			}
			rowResultado = dtbResultado.Rows[0];
			if (BD.readToInt(rowResultado["id_plano_contas_conta"]) == 0)
			{
				throw new FinanceiroException("A informação do plano de contas não foi preenchida adequadamente no cadastro de lojas (boleto id=" + idBoletoItem.ToString() + ")!!");
			}
			#endregion

			#region [ Carrega os dados ]
			boletoPlanoContasDestino.id_plano_contas_empresa = BD.readToByte(rowResultado["id_plano_contas_empresa"]);
			boletoPlanoContasDestino.id_plano_contas_grupo = BD.readToShort(rowResultado["id_plano_contas_grupo"]);
			boletoPlanoContasDestino.id_plano_contas_conta = BD.readToInt(rowResultado["id_plano_contas_conta"]);
			boletoPlanoContasDestino.natureza = BD.readToChar(rowResultado["natureza"]);
			#endregion

			return boletoPlanoContasDestino;
		}
		#endregion

		#region [ obtemBoletoPlanoContasDestinoByNumLoja ]
		/// <summary>
		/// Dado o número da loja, obtém o plano de contas para o qual deve ser lançado 
		/// o lançamento do fluxo de caixa gerado em decorrência do boleto.
		/// </summary>
		/// <param name="numeroLoja">Número da loja</param>
		/// <returns>Retorna um objeto do tipo BoletoPlanoContasDestino com os dados do plano de contas</returns>
		public static BoletoPlanoContasDestino obtemBoletoPlanoContasDestinoByNumLoja(int numeroLoja)
		{
			#region [ Declarações ]
			BoletoPlanoContasDestino boletoPlanoContasDestino = new BoletoPlanoContasDestino();
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Prepara objetos de acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			#endregion

			#region [ Consulta SQL ]
			strSql = "SELECT" +
						" id_plano_contas_empresa," +
						" id_plano_contas_grupo," +
						" id_plano_contas_conta," +
						" natureza" +
					" FROM t_LOJA" +
					" WHERE" +
						" (CONVERT(smallint, loja) = " + numeroLoja.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbResultado);
			#endregion

			#region [ Consistência ]
			if (dtbResultado.Rows.Count == 0)
			{
				throw new FinanceiroException("Não foi encontrado o registro da loja " + numeroLoja.ToString() + " ao tentar recuperar os dados do plano de contas!!");
			}
			rowResultado = dtbResultado.Rows[0];
			if (BD.readToInt(rowResultado["id_plano_contas_conta"]) == 0)
			{
				throw new FinanceiroException("A informação do plano de contas não foi preenchida adequadamente no cadastro da loja " + numeroLoja.ToString() + "!!");
			}
			#endregion

			#region [ Carrega os dados ]
			boletoPlanoContasDestino.id_plano_contas_empresa = BD.readToByte(rowResultado["id_plano_contas_empresa"]);
			boletoPlanoContasDestino.id_plano_contas_grupo = BD.readToShort(rowResultado["id_plano_contas_grupo"]);
			boletoPlanoContasDestino.id_plano_contas_conta = BD.readToInt(rowResultado["id_plano_contas_conta"]);
			boletoPlanoContasDestino.natureza = BD.readToChar(rowResultado["natureza"]);
			#endregion

			return boletoPlanoContasDestino;
		}
		#endregion

		#region [ obtemBoletoItemRateio ]
		public static DsDataSource.DtbFinBoletoItemRateioDataTable obtemBoletoItemRateio(int idBoletoItem)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbFinBoletoItemRateioDataTable dtbFinBoletoItemRateio = new DsDataSource.DtbFinBoletoItemRateioDataTable();
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT " +
						"*" +
					" FROM t_FIN_BOLETO_ITEM_RATEIO" +
					" WHERE" +
						" (id_boleto_item = " + idBoletoItem.ToString() + ")" +
					" ORDER BY" +
						" pedido";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbFinBoletoItemRateio);

			return dtbFinBoletoItemRateio;
		}
		#endregion

		#region [ obtemBoletoInformacaoPedidoLoja ]
		public static List<String> obtemBoletoInformacaoPedidoLoja(int idBoletoItem)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			List<String> listaPedidoLoja = new List<String>();
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT DISTINCT" +
						" tFBIR.pedido," +
						" tP.loja" +
					" FROM t_FIN_BOLETO_ITEM_RATEIO tFBIR" +
						" INNER JOIN t_PEDIDO tP" +
							" ON (tFBIR.pedido=tP.pedido)" +
					" WHERE" +
						" (id_boleto_item = " + idBoletoItem.ToString() + ")" +
					" ORDER BY" +
						" tFBIR.pedido";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbResultado);

			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				listaPedidoLoja.Add(dtbResultado.Rows[i]["pedido"].ToString().Trim() + "=" + dtbResultado.Rows[i]["loja"].ToString().Trim());
			}

			return listaPedidoLoja;
		}
		#endregion

		#region [ obtemListaNumeroPedidoRateio ]
		public static List<String> obtemListaNumeroPedidoRateio(int idBoleto)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			List<String> listaPedido = new List<String>();
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT DISTINCT" +
						" pedido" +
					" FROM t_FIN_BOLETO_ITEM_RATEIO" +
					" WHERE" +
						" (id_boleto = " + idBoleto.ToString() + ")" +
					" ORDER BY" +
						" pedido";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbResultado);

			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				listaPedido.Add(BD.readToString(dtbResultado.Rows[i]["pedido"]));
			}

			return listaPedido;
		}
		#endregion

		#region [ obtemRegistroPrincipalBoleto ]
		public static DsDataSource.DtbFinBoletoRow obtemRegistroPrincipalBoleto(int idBoleto)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbFinBoletoDataTable dtbFinBoleto = new DsDataSource.DtbFinBoletoDataTable();
			#endregion

			#region [ Consistência ]
			if (idBoleto == 0) return null;
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Obtém dados do registro principal ]
			strSql = "SELECT " +
						"*" +
					" FROM t_FIN_BOLETO" +
					" WHERE" +
						" (id = " + idBoleto.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbFinBoleto);
			#endregion

			if (dtbFinBoleto.Rows.Count == 0) return null;
			return (DsDataSource.DtbFinBoletoRow)dtbFinBoleto.Rows[0];
		}
		#endregion

		#region [ obtemRegistroPrincipalBoletoByIdBoletoItem ]
		public static DsDataSource.DtbFinBoletoRow obtemRegistroPrincipalBoletoByIdBoletoItem(int idBoletoItem)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			SqlDataReader drTBI;
			DsDataSource.DtbFinBoletoDataTable dtbFinBoleto = new DsDataSource.DtbFinBoletoDataTable();
			int idBoleto = 0;
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Obtém o Id do registro principal ]
			strSql = "SELECT" +
						" id_boleto" +
					" FROM t_FIN_BOLETO_ITEM" +
					" WHERE" +
						" (id = " + idBoletoItem.ToString() + ")";
			cmCommand.CommandText = strSql;
			drTBI = cmCommand.ExecuteReader();
			try
			{
				if (drTBI.Read())
				{
					idBoleto = (int)drTBI["id_boleto"];
					if (idBoleto == 0) return null;
				}
				else
				{
					return null;
				}
			}
			finally
			{
				drTBI.Close();
			}
			#endregion

			#region [ Obtém dados do registro principal ]
			strSql = "SELECT " +
						"*" +
					" FROM t_FIN_BOLETO" +
					" WHERE" +
						" (id = " + idBoleto.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbFinBoleto);
			#endregion

			if (dtbFinBoleto.Rows.Count == 0) return null;
			return (DsDataSource.DtbFinBoletoRow)dtbFinBoleto.Rows[0];
		}
		#endregion

		#region [ obtemRegistroPrincipalBoletoByNossoNumero ]
		public static DsDataSource.DtbFinBoletoRow obtemRegistroPrincipalBoletoByNossoNumero(int id_boleto_cedente, String nossoNumeroSemDigito, DateTime dtVencto)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbFinBoletoDataTable dtbFinBoleto = new DsDataSource.DtbFinBoletoDataTable();
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			#region [ Consistência ]
			if (nossoNumeroSemDigito == null) return null;
			if (nossoNumeroSemDigito.Trim().Length == 0) return null;
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			rowBoletoItem = obtemRegistroBoletoItemByNossoNumero(id_boleto_cedente, nossoNumeroSemDigito, dtVencto);
			if (rowBoletoItem == null) return null;
			if (rowBoletoItem.id_boleto == 0) return null;

			#region [ Obtém dados do registro principal ]
			strSql = "SELECT " +
						"*" +
					" FROM t_FIN_BOLETO" +
					" WHERE" +
						" (id = " + rowBoletoItem.id_boleto.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbFinBoleto);
			#endregion

			if (dtbFinBoleto.Rows.Count == 0) return null;
			return (DsDataSource.DtbFinBoletoRow)dtbFinBoleto.Rows[0];
		}
		#endregion

		#region [ obtemRegistroPrincipalBoletoByNumeroDocumento ]
		public static DsDataSource.DtbFinBoletoRow obtemRegistroPrincipalBoletoByNumeroDocumento(int id_boleto_cedente, String numeroDocumento, DateTime dtVencto)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbFinBoletoDataTable dtbFinBoleto = new DsDataSource.DtbFinBoletoDataTable();
			DsDataSource.DtbFinBoletoItemRow rowBoletoItem;
			#endregion

			#region [ Consistência ]
			if (numeroDocumento == null) return null;
			if (numeroDocumento.Trim().Length == 0) return null;
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			rowBoletoItem = obtemRegistroBoletoItemByNumeroDocumento(id_boleto_cedente, numeroDocumento, dtVencto);
			if (rowBoletoItem == null) return null;
			if (rowBoletoItem.id_boleto == 0) return null;

			#region [ Obtém dados do registro principal ]
			strSql = "SELECT " +
						"*" +
					" FROM t_FIN_BOLETO" +
					" WHERE" +
						" (id = " + rowBoletoItem.id_boleto.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbFinBoleto);
			#endregion

			if (dtbFinBoleto.Rows.Count == 0) return null;
			return (DsDataSource.DtbFinBoletoRow)dtbFinBoleto.Rows[0];
		}
		#endregion

		#region [ obtemRegistroBoletoItem ]
		public static DsDataSource.DtbFinBoletoItemRow obtemRegistroBoletoItem(int idBoletoItem)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbFinBoletoItemDataTable dtbFinBoletoItem = new DsDataSource.DtbFinBoletoItemDataTable();
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Obtém dados ]
			strSql = "SELECT " +
						"*" +
					" FROM t_FIN_BOLETO_ITEM" +
					" WHERE" +
						" (id = " + idBoletoItem.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbFinBoletoItem);
			#endregion

			if (dtbFinBoletoItem.Rows.Count == 0) return null;
			return (DsDataSource.DtbFinBoletoItemRow)dtbFinBoletoItem.Rows[0];
		}
		#endregion

		#region [ obtemRegistroBoletoItemByNossoNumero ]
		public static DsDataSource.DtbFinBoletoItemRow obtemRegistroBoletoItemByNossoNumero(int id_boleto_cedente, String nossoNumeroSemDigito, DateTime dtVencto)
		{
			#region [ Declarações ]
			int intQtdeDtVenctoIgual = 0;
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbFinBoletoItemDataTable dtbFinBoletoItem = new DsDataSource.DtbFinBoletoItemDataTable();
			DsDataSource.DtbFinBoletoItemRow rowResposta = null;
			#endregion

			#region [ Consistência ]
			if (nossoNumeroSemDigito == null) return null;
			if (nossoNumeroSemDigito.Trim().Length == 0) return null;
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Obtém dados ]
			strSql = "SELECT " +
						"tFBI.*" +
					" FROM t_FIN_BOLETO_ITEM tFBI" +
						" INNER JOIN t_FIN_BOLETO tFB ON (tFBI.id_boleto=tFB.id)" +
					" WHERE" +
						" (nosso_numero = '" + nossoNumeroSemDigito.Trim() + "')" +
						" AND (id_boleto_cedente = " + id_boleto_cedente.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbFinBoletoItem);
			#endregion

			#region [ Não encontrou nenhum registro ]
			if (dtbFinBoletoItem.Rows.Count == 0) return null;
			#endregion

			#region [ Encontrou apenas 1 registro ]
			if (dtbFinBoletoItem.Rows.Count == 1) return (DsDataSource.DtbFinBoletoItemRow)dtbFinBoletoItem.Rows[0];
			#endregion

			#region [ Encontrou mais do que 1 registro, tentar determinar qual é o correto pela data de vencimento ]
			if (dtVencto > DateTime.MinValue)
			{
				for (int i = 0; i < dtbFinBoletoItem.Rows.Count; i++)
				{
					if (((DsDataSource.DtbFinBoletoItemRow)dtbFinBoletoItem.Rows[i]).dt_vencto == dtVencto)
					{
						intQtdeDtVenctoIgual++;
						rowResposta = (DsDataSource.DtbFinBoletoItemRow)dtbFinBoletoItem.Rows[i];
					}
				}
			}
			if ((intQtdeDtVenctoIgual == 1) && (rowResposta != null)) return rowResposta;
			#endregion

			return null;
		}
		#endregion

		#region [ obtemRegistroBoletoItemByNossoNumero ]
		public static DsDataSource.DtbFinBoletoItemRow obtemRegistroBoletoItemByNossoNumero(int id_boleto_cedente, String nossoNumeroSemDigito)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbFinBoletoItemDataTable dtbFinBoletoItem = new DsDataSource.DtbFinBoletoItemDataTable();
			#endregion

			#region [ Consistência ]
			if (nossoNumeroSemDigito == null) return null;
			if (nossoNumeroSemDigito.Trim().Length == 0) return null;
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Obtém dados ]
			strSql = "SELECT " +
						"tFBI.*" +
					" FROM t_FIN_BOLETO_ITEM tFBI" +
						" INNER JOIN t_FIN_BOLETO tFB ON (tFBI.id_boleto=tFB.id)" +
					" WHERE" +
						" (nosso_numero = '" + nossoNumeroSemDigito.Trim() + "')" +
						" AND (id_boleto_cedente = " + id_boleto_cedente.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbFinBoletoItem);
			#endregion

			#region [ Não encontrou nenhum registro ]
			if (dtbFinBoletoItem.Rows.Count == 0) return null;
			#endregion

			#region [ Encontrou apenas 1 registro ]
			if (dtbFinBoletoItem.Rows.Count == 1) return (DsDataSource.DtbFinBoletoItemRow)dtbFinBoletoItem.Rows[0];
			#endregion

			return null;
		}
		#endregion

		#region [ obtemRegistroBoletoItemByNumeroDocumento ]
		public static DsDataSource.DtbFinBoletoItemRow obtemRegistroBoletoItemByNumeroDocumento(int id_boleto_cedente, String numeroDocumento, DateTime dtVencto)
		{
			#region [ Declarações ]
			int intQtdeDtVenctoIgual = 0;
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbFinBoletoItemDataTable dtbFinBoletoItem = new DsDataSource.DtbFinBoletoItemDataTable();
			DsDataSource.DtbFinBoletoItemRow rowResposta = null;
			#endregion

			#region [ Consistência ]
			if (numeroDocumento == null) return null;
			if (numeroDocumento.Trim().Length == 0) return null;
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Obtém dados ]
			strSql = "SELECT " +
						"tFBI.*" +
					" FROM t_FIN_BOLETO_ITEM tFBI" +
						" INNER JOIN t_FIN_BOLETO tFB ON (tFBI.id_boleto=tFB.id)" +
					" WHERE" +
						" (numero_documento = '" + numeroDocumento.Trim() + "')" +
						" AND (id_boleto_cedente = " + id_boleto_cedente.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbFinBoletoItem);
			#endregion

			#region [ Não encontrou nenhum registro ]
			if (dtbFinBoletoItem.Rows.Count == 0) return null;
			#endregion

			#region [ Encontrou apenas 1 registro ]
			if (dtbFinBoletoItem.Rows.Count == 1) return (DsDataSource.DtbFinBoletoItemRow)dtbFinBoletoItem.Rows[0];
			#endregion

			#region [ Encontrou mais do que 1 registro, tentar determinar qual é o correto pela data de vencimento ]
			if (dtVencto > DateTime.MinValue)
			{
				for (int i = 0; i < dtbFinBoletoItem.Rows.Count; i++)
				{
					if (((DsDataSource.DtbFinBoletoItemRow)dtbFinBoletoItem.Rows[i]).dt_vencto == dtVencto)
					{
						intQtdeDtVenctoIgual++;
						rowResposta = (DsDataSource.DtbFinBoletoItemRow)dtbFinBoletoItem.Rows[i];
					}
				}
			}
			if ((intQtdeDtVenctoIgual == 1) && (rowResposta != null)) return rowResposta;
			#endregion

			return null;
		}
		#endregion

		#region [ boletoInsere ]
		/// <summary>
		/// Grava os dados de uma nova série de boletos de um cliente, podendo conter uma ou mais parcelas.
		/// </summary>
		/// <param name="usuario">
		/// Usuário que está realizando a operação.
		/// </param>
		/// <param name="boleto">
		/// Objeto do tipo Boleto contendo os dados p/ cadastrar.
		/// </param>
		/// <param name="strDescricaoLog">
		/// Retorna texto com detalhes da operação a serem registradas no log.
		/// </param>
		/// <param name="strMsgErro">
		/// Em caso de erro, retorna mensagem com descrição.
		/// </param>
		/// <returns>
		/// true: gravação bem sucedida.
		/// false: falha na gravação.
		/// </returns>
		public static bool boletoInsere(String usuario,
										Boleto boleto,
										ref String strDescricaoLog,
										ref String strMsgErro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			bool blnGerouNsu;
			int intNsuBoleto = 0;
			int intNsuBoletoItem = 0;
			int intRetorno;
			String strOperacao = "Gravação de boleto";
			StringBuilder sbLog = new StringBuilder("");
			StringBuilder sbLogLinha;
			#endregion

			try
			{
				strMsgErro = "";

				#region [ Gera o NSU para o boleto ]
				blnGerouNsu = BD.geraNsu(Global.Cte.FIN.NSU.T_FIN_BOLETO, ref intNsuBoleto, ref strMsgErro);
				if (!blnGerouNsu)
				{
					strMsgErro = "Falha ao tentar gerar o NSU para o boleto!!\n" + strMsgErro;
					return false;
				}
				boleto.id = intNsuBoleto;
				#endregion

				#region [ Gera o NSU para as parcelas ]
				for (int i = 0; i < boleto.listaBoletoItem.Count; i++)
				{
					blnGerouNsu = BD.geraNsu(Global.Cte.FIN.NSU.T_FIN_BOLETO_ITEM, ref intNsuBoletoItem, ref strMsgErro);
					if (!blnGerouNsu)
					{
						strMsgErro = "Falha ao tentar gerar o NSU para as parcelas do boleto!!\n" + strMsgErro;
						return false;
					}
					boleto.listaBoletoItem[i].id = intNsuBoletoItem;
					boleto.listaBoletoItem[i].id_boleto = boleto.id;
					for (int j = 0; j < boleto.listaBoletoItem[i].listaBoletoItemRateio.Count; j++)
					{
						boleto.listaBoletoItem[i].listaBoletoItemRateio[j].id_boleto = boleto.id;
						boleto.listaBoletoItem[i].listaBoletoItemRateio[j].id_boleto_item = boleto.listaBoletoItem[i].id;
					}
				}
				#endregion

				try
				{
					#region [ Tenta gravar o boleto ]

					#region [ Preenche o valor dos parâmetros ]
					cmBoletoInsert.Parameters["@id"].Value = boleto.id;
					cmBoletoInsert.Parameters["@id_cliente"].Value = boleto.id_cliente;
					cmBoletoInsert.Parameters["@id_nf_parcela_pagto"].Value = boleto.id_nf_parcela_pagto;
					cmBoletoInsert.Parameters["@tipo_vinculo"].Value = boleto.tipo_vinculo;
					cmBoletoInsert.Parameters["@status"].Value = Global.Cte.FIN.CodBoletoStatus.INICIAL;
					cmBoletoInsert.Parameters["@numero_NF"].Value = boleto.numero_NF;
					cmBoletoInsert.Parameters["@num_documento_boleto_avulso"].Value = boleto.num_documento_boleto_avulso;
					cmBoletoInsert.Parameters["@qtde_parcelas"].Value = boleto.qtde_parcelas;
					cmBoletoInsert.Parameters["@id_boleto_cedente"].Value = boleto.id_boleto_cedente;
					cmBoletoInsert.Parameters["@codigo_empresa"].Value = boleto.codigo_empresa;
					cmBoletoInsert.Parameters["@nome_empresa"].Value = boleto.nome_empresa;
					cmBoletoInsert.Parameters["@num_banco"].Value = boleto.num_banco;
					cmBoletoInsert.Parameters["@nome_banco"].Value = boleto.nome_banco;
					cmBoletoInsert.Parameters["@agencia"].Value = boleto.agencia;
					cmBoletoInsert.Parameters["@digito_agencia"].Value = boleto.digito_agencia;
					cmBoletoInsert.Parameters["@conta"].Value = boleto.conta;
					cmBoletoInsert.Parameters["@digito_conta"].Value = boleto.digito_conta;
					cmBoletoInsert.Parameters["@carteira"].Value = boleto.carteira;
					cmBoletoInsert.Parameters["@juros_mora"].Value = boleto.juros_mora;
					cmBoletoInsert.Parameters["@perc_multa"].Value = boleto.perc_multa;
					cmBoletoInsert.Parameters["@primeira_instrucao"].Value = boleto.primeira_instrucao;
					cmBoletoInsert.Parameters["@segunda_instrucao"].Value = boleto.segunda_instrucao;
					cmBoletoInsert.Parameters["@qtde_dias_protesto"].Value = boleto.qtde_dias_protesto;
					cmBoletoInsert.Parameters["@qtde_dias_decurso_prazo"].Value = boleto.qtde_dias_decurso_prazo;
					cmBoletoInsert.Parameters["@tipo_sacado"].Value = boleto.tipo_sacado;
					cmBoletoInsert.Parameters["@num_inscricao_sacado"].Value = boleto.num_inscricao_sacado;
					cmBoletoInsert.Parameters["@nome_sacado"].Value = boleto.nome_sacado;
					cmBoletoInsert.Parameters["@endereco_sacado"].Value = boleto.endereco_sacado;
					cmBoletoInsert.Parameters["@cep_sacado"].Value = boleto.cep_sacado;
					cmBoletoInsert.Parameters["@bairro_sacado"].Value = boleto.bairro_sacado;
					cmBoletoInsert.Parameters["@cidade_sacado"].Value = boleto.cidade_sacado;
					cmBoletoInsert.Parameters["@uf_sacado"].Value = boleto.uf_sacado;
					cmBoletoInsert.Parameters["@email_sacado"].Value = boleto.email_sacado;
					cmBoletoInsert.Parameters["@segunda_mensagem"].Value = boleto.segunda_mensagem;
					cmBoletoInsert.Parameters["@mensagem_1"].Value = boleto.mensagem_1;
					cmBoletoInsert.Parameters["@mensagem_2"].Value = boleto.mensagem_2;
					cmBoletoInsert.Parameters["@mensagem_3"].Value = boleto.mensagem_3;
					cmBoletoInsert.Parameters["@mensagem_4"].Value = boleto.mensagem_4;
					cmBoletoInsert.Parameters["@usuario_cadastro"].Value = usuario;
					cmBoletoInsert.Parameters["@usuario_ult_atualizacao"].Value = usuario;
					#endregion

					#region [ Monta texto para o log em arquivo ]
					foreach (SqlParameter item in cmBoletoInsert.Parameters)
					{
						if (sbLog.Length > 0) sbLog.Append("; ");
						sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
					}
					#endregion

					#region [ Tenta inserir o registro ]
					strMsgErro = "";
					try
					{
						intRetorno = BD.executaNonQuery(ref cmBoletoInsert);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						strMsgErro = ex.Message;
						Global.gravaLogAtividade(strOperacao + " - Exception: " + sbLog.ToString() + "\n" + ex.ToString());
					}
					#endregion

					#region [ Gravou o registro principal? ]
					if (intRetorno == 0)
					{
						strMsgErro = "Falha ao tentar gravar o registro principal do boleto!!\n" + strMsgErro;
						return false;
					}
					#endregion

					#region [ Grava as parcelas ]
					for (int i = 0; i < boleto.listaBoletoItem.Count; i++)
					{
						#region [ Preenche o campo num_controle_participante com o Id do registro ]
						boleto.listaBoletoItem[i].num_controle_participante = Global.Cte.Etc.PREFIXO_BOLETO_NUM_CONTROLE_PARTICIPANTE + "=" + boleto.listaBoletoItem[i].id.ToString();
						#endregion

						#region [ Preenche o valor dos parâmetros ]
						cmBoletoItemInsert.Parameters["@id"].Value = boleto.listaBoletoItem[i].id;
						cmBoletoItemInsert.Parameters["@id_boleto"].Value = boleto.listaBoletoItem[i].id_boleto;
						cmBoletoItemInsert.Parameters["@num_parcela"].Value = boleto.listaBoletoItem[i].num_parcela;
						cmBoletoItemInsert.Parameters["@status"].Value = Global.Cte.FIN.CodBoletoItemStatus.INICIAL;
						cmBoletoItemInsert.Parameters["@tipo_vencimento"].Value = boleto.listaBoletoItem[i].tipo_vencimento;
						cmBoletoItemInsert.Parameters["@dt_vencto"].Value = Global.formataDataYyyyMmDdComSeparador(boleto.listaBoletoItem[i].dt_vencto);
						cmBoletoItemInsert.Parameters["@valor"].Value = boleto.listaBoletoItem[i].valor;
						cmBoletoItemInsert.Parameters["@bonificacao_por_dia"].Value = boleto.listaBoletoItem[i].bonificacao_por_dia;
						cmBoletoItemInsert.Parameters["@valor_por_dia_atraso"].Value = boleto.listaBoletoItem[i].valor_por_dia_atraso;
						cmBoletoItemInsert.Parameters["@dt_limite_desconto"].Value = Global.formataDataYyyyMmDdComSeparador(boleto.listaBoletoItem[i].dt_limite_desconto);
						cmBoletoItemInsert.Parameters["@valor_desconto"].Value = boleto.listaBoletoItem[i].valor_desconto;
						cmBoletoItemInsert.Parameters["@numero_documento"].Value = boleto.listaBoletoItem[i].numero_documento;
						cmBoletoItemInsert.Parameters["@primeira_mensagem"].Value = boleto.listaBoletoItem[i].primeira_mensagem;
						cmBoletoItemInsert.Parameters["@num_controle_participante"].Value = boleto.listaBoletoItem[i].num_controle_participante;
						cmBoletoItemInsert.Parameters["@usuario_ult_atualizacao"].Value = usuario;
						cmBoletoItemInsert.Parameters["@st_instrucao_protesto"].Value = boleto.listaBoletoItem[i].st_instrucao_protesto;
						#endregion

						#region [ Monta texto para o log em arquivo ]
						sbLogLinha = new StringBuilder("");
						foreach (SqlParameter item in cmBoletoItemInsert.Parameters)
						{
							if (sbLogLinha.Length > 0) sbLogLinha.Append("; ");
							sbLogLinha.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
						}
						if (sbLog.Length > 0) sbLog.Append("; ");
						sbLog.Append("\r\tParcela: " + sbLogLinha.ToString());
						#endregion

						#region [ Tenta inserir o registro da parcela ]
						strMsgErro = "";
						try
						{
							intRetorno = BD.executaNonQuery(ref cmBoletoItemInsert);
						}
						catch (Exception ex)
						{
							intRetorno = 0;
							strMsgErro = ex.Message;
							Global.gravaLogAtividade(strOperacao + " - Exception: " + sbLog.ToString() + "\n" + ex.ToString());
						}
						#endregion

						#region [ Gravou o registro da parcela? ]
						if (intRetorno == 0)
						{
							strMsgErro = "Falha ao tentar gravar o registro da parcela do boleto!!\n" + strMsgErro;
							return false;
						}
						#endregion

						#region [ Grava os dados do rateio desta parcela ]
						for (int j = 0; j < boleto.listaBoletoItem[i].listaBoletoItemRateio.Count; j++)
						{
							#region [ Preenche o valor dos parâmetros ]
							cmBoletoItemRateioInsert.Parameters["@id_boleto_item"].Value = boleto.listaBoletoItem[i].listaBoletoItemRateio[j].id_boleto_item;
							cmBoletoItemRateioInsert.Parameters["@pedido"].Value = boleto.listaBoletoItem[i].listaBoletoItemRateio[j].pedido;
							cmBoletoItemRateioInsert.Parameters["@id_boleto"].Value = boleto.listaBoletoItem[i].listaBoletoItemRateio[j].id_boleto;
							cmBoletoItemRateioInsert.Parameters["@valor"].Value = boleto.listaBoletoItem[i].listaBoletoItemRateio[j].valor;
							#endregion

							#region [ Monta texto para o log em arquivo ]
							sbLogLinha = new StringBuilder("");
							foreach (SqlParameter item in cmBoletoItemRateioInsert.Parameters)
							{
								if (sbLogLinha.Length > 0) sbLogLinha.Append("; ");
								sbLogLinha.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
							}
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append("\r\t\tRateio: " + sbLogLinha.ToString());
							#endregion

							#region [ Tenta inserir o registro do rateio ]
							strMsgErro = "";
							try
							{
								intRetorno = BD.executaNonQuery(ref cmBoletoItemRateioInsert);
							}
							catch (Exception ex)
							{
								intRetorno = 0;
								strMsgErro = ex.Message;
								Global.gravaLogAtividade(strOperacao + " - Exception: " + sbLog.ToString() + "\n" + ex.ToString());
							}
							#endregion

							#region [ Gravou o registro do rateio? ]
							if (intRetorno == 0)
							{
								strMsgErro = "Falha ao tentar gravar o registro do rateio da parcela do boleto!!\n" + strMsgErro;
								return false;
							}
							#endregion
						}
						#endregion
					}
					#endregion

					#region [ Operação bem sucedida ]
					try
					{
						strDescricaoLog = sbLog.ToString();
						Global.gravaLogAtividade(strOperacao + " - Sucesso: " + sbLog.ToString());
						blnSucesso = true;
					}
					catch (Exception ex)
					{
						// Para o usuário, exibe uma mensagem mais sucinta
						strMsgErro = ex.Message;
						// No log em arquivo, grava o stack de erro completo
						Global.gravaLogAtividade(strOperacao + " - Falha: " + sbLog.ToString() + "\n" + ex.ToString());
						return false;
					}
					#endregion

					#endregion
				}
				catch (Exception ex)
				{
					// Para o usuário, exibe uma mensagem mais sucinta
					strMsgErro = ex.Message;
					// No log em arquivo, grava o stack de erro completo
					Global.gravaLogAtividade(strOperacao + " - Falha: " + sbLog.ToString() + "\n" + ex.ToString());
					return false;
				}

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar gravar o boleto no banco de dados!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha: " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ boletoMovimentoInsere ]
		/// <summary>
		/// Grava o registro na tabela de movimentações de boletos (t_FIN_BOLETO_MOVIMENTO).
		/// IMPORTANTE: alguns campos podem estar vazios quando for o caso de ser um boleto
		/// desconhecido ou não identificado (tipicamente, ocorrência 17 - liquidação após baixa
		/// ou Título não registrado).
		/// </summary>
		/// <param name="usuario">Usuário que está realizando o processamento da carga do arquivo de retorno</param>
		/// <param name="id_arq_retorno">Identificação do registro associado na tabela t_FIN_BOLETO_ARQ_RETORNO</param>
		/// <param name="id_boleto">Identificação do registro da tabela t_FIN_BOLETO (pode estar zerado no caso de boleto não identificado)</param>
		/// <param name="id_boleto_item">Identificação do registro da tabela t_FIN_BOLETO_ITEM (pode estar zerado no caso de boleto não identificado)</param>
		/// <param name="identificacaoOcorrencia">Código de identificação da ocorrência</param>
		/// <param name="motivosRejeicoes">Motivos das ocorrências</param>
		/// <param name="dataOcorrencia">Data da ocorrência no banco</param>
		/// <param name="numeroDocumento">Número do documento (pode estar zerado no caso de boleto não identificado)</param>
		/// <param name="nossoNumero">Nosso número (sem dígito)</param>
		/// <param name="digitoNossoNumero">Dígito do nosso número</param>
		/// <param name="dataVencto">Data de vencimento do título</param>
		/// <param name="valorTitulo">Valor do título</param>
		/// <param name="valorDespesasCobranca">Despesas de cobrança para os códigos de ocorrência 02 (entrada confirmada) e 28 (débito de tarifas). Campo da posição 176 a 188 do registro tipo 1 do arquivo de retorno.</param>
		/// <param name="valorOutrasDespesas">Outras despesas / Custas de protesto. Campo da posição 189 a 201 do registro tipo 1 do arquivo de retorno.</param>
		/// <param name="valorIofDevido">IOF devido. Campo da posição 215 a 227 do registro tipo 1 do arquivo de retorno.</param>
		/// <param name="valorAbatimentoConcedido">Abatimento concedido sobre o título. Campo da posição 228 a 240 do registro tipo 1 do arquivo de retorno.</param>
		/// <param name="valorDescontoConcedido">Desconto concedido. Campo da posição 241 a 253 do registro tipo 1 do arquivo de retorno.</param>
		/// <param name="valorPago">Valor total recebido. Campo da posição 254 a 266 do registro tipo 1 do arquivo de retorno.</param>
		/// <param name="valorJurosMora">Juros de mora. Campo da posição 267 a 279 do registro tipo 1 do arquivo de retorno.</param>
		/// <param name="dataCredito">Data do crédito. Campo da posição 296 a 301 do registro tipo 1 do arquivo de retorno.</param>
		/// <param name="strMsgErro">No caso de erro, retorna a mensagem de erro</param>
		/// <returns>
		/// true: sucesso na gravação dos dados
		/// false: falha na gravação dos dados
		/// </returns>
		public static bool boletoMovimentoInsere(String usuario,
												 int id_arq_retorno,
												 int id_boleto,
												 int id_boleto_item,
												 String identificacaoOcorrencia,
												 String motivosRejeicoes,
												 DateTime dataOcorrencia,
												 String numeroDocumento,
												 String nossoNumero,
												 String digitoNossoNumero,
												 DateTime dataVencto,
												 decimal valorTitulo,
												 decimal valorDespesasCobranca,
												 decimal valorOutrasDespesas,
												 decimal valorIofDevido,
												 decimal valorAbatimentoConcedido,
												 decimal valorDescontoConcedido,
												 decimal valorPago,
												 decimal valorJurosMora,
												 DateTime dataCredito,
												 ref String strMsgErro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			bool blnGerouNsu;
			int intRetorno;
			int intNsuBoletoMovimento = 0;
			String strOperacao = "Gravação dos dados de movimento de boletos";
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistência ]
				if (id_arq_retorno <= 0)
				{
					strMsgErro = "Número de identificação do registro associado na tabela t_FIN_BOLETO_ARQ_RETORNO não foi informado!!";
					return false;
				}
				#endregion

				try
				{
					#region [ Gera o NSU para o registro ]
					blnGerouNsu = BD.geraNsu(Global.Cte.FIN.NSU.T_FIN_BOLETO_MOVIMENTO, ref intNsuBoletoMovimento, ref strMsgErro);
					if (!blnGerouNsu)
					{
						strMsgErro = "Falha ao tentar gerar o NSU para o registro de movimentação do boleto!!\n" + strMsgErro;
						return false;
					}
					#endregion

					#region [ Preenche o valor dos parâmetros ]
					cmBoletoMovimentoInsert.Parameters["@id"].Value = intNsuBoletoMovimento;
					cmBoletoMovimentoInsert.Parameters["@id_arq_retorno"].Value = id_arq_retorno;
					cmBoletoMovimentoInsert.Parameters["@id_boleto"].Value = id_boleto;
					cmBoletoMovimentoInsert.Parameters["@id_boleto_item"].Value = id_boleto_item;
					cmBoletoMovimentoInsert.Parameters["@identificacao_ocorrencia"].Value = identificacaoOcorrencia;
					cmBoletoMovimentoInsert.Parameters["@motivos_rejeicoes"].Value = motivosRejeicoes;
					cmBoletoMovimentoInsert.Parameters["@data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrencia);
					cmBoletoMovimentoInsert.Parameters["@numero_documento"].Value = numeroDocumento.Trim();
					cmBoletoMovimentoInsert.Parameters["@nosso_numero"].Value = nossoNumero.Trim();
					cmBoletoMovimentoInsert.Parameters["@digito_nosso_numero"].Value = digitoNossoNumero.Trim();
					cmBoletoMovimentoInsert.Parameters["@dt_vencto"].Value = Global.formataDataYyyyMmDdComSeparador(dataVencto);
					cmBoletoMovimentoInsert.Parameters["@vl_titulo"].Value = valorTitulo;
					cmBoletoMovimentoInsert.Parameters["@vl_despesas_cobranca"].Value = valorDespesasCobranca;
					cmBoletoMovimentoInsert.Parameters["@vl_outras_despesas"].Value = valorOutrasDespesas;
					cmBoletoMovimentoInsert.Parameters["@vl_IOF"].Value = valorIofDevido;
					cmBoletoMovimentoInsert.Parameters["@vl_abatimento"].Value = valorAbatimentoConcedido;
					cmBoletoMovimentoInsert.Parameters["@vl_desconto"].Value = valorDescontoConcedido;
					cmBoletoMovimentoInsert.Parameters["@vl_pago"].Value = valorPago;
					cmBoletoMovimentoInsert.Parameters["@vl_juros_mora"].Value = valorJurosMora;
					cmBoletoMovimentoInsert.Parameters["@dt_credito"].Value = Global.formataDataYyyyMmDdComSeparador(dataCredito);
					#endregion

					#region [ Tenta inserir o registro ]
					strMsgErro = "";
					try
					{
						intRetorno = BD.executaNonQuery(ref cmBoletoMovimentoInsert);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						strMsgErro = ex.Message;
						Global.gravaLogAtividade(strOperacao + " - Exception!!\n" + ex.ToString());
					}
					#endregion

					#region [ Gravou o registro? ]
					if (intRetorno == 0)
					{
						strMsgErro = "Falha ao tentar gravar o registro de movimentação do boleto!!\n" + strMsgErro;
						return false;
					}
					#endregion

					blnSucesso = true;
				}
				catch (Exception ex)
				{
					// Para o usuário, exibe uma mensagem mais sucinta
					strMsgErro = ex.Message;
					// No log em arquivo, grava o stack de erro completo
					Global.gravaLogAtividade(strOperacao + " - Falha!!\n" + ex.ToString());
					return false;
				}

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar inserir o registro de movimentação do boleto!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ boletoOcorrenciaInsere ]
		/// <summary>
		/// Grava o registro na tabela de ocorrências de boletos (t_FIN_BOLETO_OCORRENCIA).
		/// IMPORTANTE: alguns campos podem estar vazios quando for o caso de ser um boleto
		/// desconhecido ou não identificado (tipicamente, ocorrência 17 - liquidação após baixa
		/// ou Título não registrado).
		/// São gravados como ocorrências os registros do arquivo de retorno que necessitam de
		/// análise humana.
		/// Podem ocorrer as seguintes situações:
		///		1) Boletos já tratados pelo sistema, mas que precisam informar alguma situação
		///		   especial para o usuário (ex: boleto pago com valor maior que o esperado).
		///		2) Boletos com código de identificação de ocorrência desconhecido e/ou não tratado. 
		///		   É a chamada "vala comum".
		/// </summary>
		/// <param name="usuario">Usuário que está realizando o processamento da carga do arquivo de retorno</param>
		/// <param name="id_arq_retorno">Identificação do registro associado na tabela t_FIN_BOLETO_ARQ_RETORNO</param>
		/// <param name="id_boleto">Identificação do registro da tabela t_FIN_BOLETO (pode estar zerado no caso de boleto não identificado)</param>
		/// <param name="id_boleto_item">Identificação do registro da tabela t_FIN_BOLETO_ITEM (pode estar zerado no caso de boleto não identificado)</param>
		/// <param name="numeroDocumento">Número do documento (pode estar zerado no caso de boleto não identificado)</param>
		/// <param name="nossoNumero">Nosso número (sem dígito)</param>
		/// <param name="digitoNossoNumero">Dígito do nosso número</param>
		/// <param name="dataVencto">Data de vencimento do título</param>
		/// <param name="valorTitulo">Valor do título</param>
		/// <param name="identificacaoOcorrencia">Código de identificação da ocorrência</param>
		/// <param name="motivosRejeicoes">Motivos das ocorrências</param>
		/// <param name="motivoCodigoOcorrencia19">Motivo do código de ocorrência 19 (confirmação de instrução de protesto)</param>
		/// <param name="dataOcorrencia">Data da ocorrência no banco</param>
		/// <param name="obsOcorrencia">Observações e/ou detalhes sobre a ocorrência</param>
		/// <param name="linhaTextoRegistroArquivo">Registro (linha) original do arquivo de retorno na íntegra</param>
		/// <param name="strMsgErro">Retorna a mensagem de erro em caso de ocorrer erro</param>
		/// <returns>
		/// true: sucesso na gravação dos dados
		/// false: falha na gravação dos dados
		/// </returns>
		public static bool boletoOcorrenciaInsere(String usuario,
												 int id_arq_retorno,
												 int id_boleto_cedente,
												 int id_boleto,
												 int id_boleto_item,
												 byte st_divergencia_valor,
												 String numeroDocumento,
												 String nossoNumero,
												 String digitoNossoNumero,
												 DateTime dataVencto,
												 decimal valorTitulo,
												 String identificacaoOcorrencia,
												 String motivosRejeicoes,
												 String motivoCodigoOcorrencia19,
												 DateTime dataOcorrencia,
												 String obsOcorrencia,
												 String linhaTextoRegistroArquivo,
												 ref String strMsgErro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			bool blnGerouNsu;
			int intRetorno;
			int intNsuBoletoOcorrencia = 0;
			String strOperacao = "Gravação de novo registro de ocorrência para o boleto";
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistência ]
				if (id_arq_retorno <= 0)
				{
					strMsgErro = "Número de identificação do registro associado na tabela t_FIN_BOLETO_ARQ_RETORNO não foi informado!!";
					return false;
				}
				#endregion

				try
				{
					#region [ Gera o NSU para o registro ]
					blnGerouNsu = BD.geraNsu(Global.Cte.FIN.NSU.T_FIN_BOLETO_OCORRENCIA, ref intNsuBoletoOcorrencia, ref strMsgErro);
					if (!blnGerouNsu)
					{
						strMsgErro = "Falha ao tentar gerar o NSU para o registro de ocorrência para o boleto!!\n" + strMsgErro;
						return false;
					}
					#endregion

					#region [ Preenche o valor dos parâmetros ]
					cmBoletoOcorrenciaInsert.Parameters["@id"].Value = intNsuBoletoOcorrencia;
					cmBoletoOcorrenciaInsert.Parameters["@id_arq_retorno"].Value = id_arq_retorno;
					cmBoletoOcorrenciaInsert.Parameters["@id_boleto_cedente"].Value = id_boleto_cedente;
					cmBoletoOcorrenciaInsert.Parameters["@id_boleto"].Value = id_boleto;
					cmBoletoOcorrenciaInsert.Parameters["@id_boleto_item"].Value = id_boleto_item;
					cmBoletoOcorrenciaInsert.Parameters["@st_divergencia_valor"].Value = st_divergencia_valor;
					cmBoletoOcorrenciaInsert.Parameters["@numero_documento"].Value = numeroDocumento.Trim();
					cmBoletoOcorrenciaInsert.Parameters["@nosso_numero"].Value = nossoNumero.Trim();
					cmBoletoOcorrenciaInsert.Parameters["@digito_nosso_numero"].Value = digitoNossoNumero.Trim();
					cmBoletoOcorrenciaInsert.Parameters["@dt_vencto"].Value = Global.formataDataYyyyMmDdComSeparador(dataVencto);
					cmBoletoOcorrenciaInsert.Parameters["@vl_titulo"].Value = valorTitulo;
					cmBoletoOcorrenciaInsert.Parameters["@identificacao_ocorrencia"].Value = identificacaoOcorrencia;
					cmBoletoOcorrenciaInsert.Parameters["@motivos_rejeicoes"].Value = motivosRejeicoes;
					cmBoletoOcorrenciaInsert.Parameters["@motivo_ocorrencia_19"].Value = motivoCodigoOcorrencia19;
					cmBoletoOcorrenciaInsert.Parameters["@data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrencia);
					cmBoletoOcorrenciaInsert.Parameters["@obs_ocorrencia"].Value = obsOcorrencia;
					cmBoletoOcorrenciaInsert.Parameters["@registro_arq_retorno"].Value = linhaTextoRegistroArquivo;
					#endregion

					#region [ Tenta inserir o registro ]
					strMsgErro = "";
					try
					{
						intRetorno = BD.executaNonQuery(ref cmBoletoOcorrenciaInsert);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						strMsgErro = ex.Message;
						Global.gravaLogAtividade(strOperacao + " - Exception!!\n" + ex.ToString());
					}
					#endregion

					#region [ Gravou o registro? ]
					if (intRetorno == 0)
					{
						strMsgErro = "Falha ao tentar gravar o registro de ocorrência para o boleto!!\n" + strMsgErro;
						return false;
					}
					#endregion

					blnSucesso = true;
				}
				catch (Exception ex)
				{
					// Para o usuário, exibe uma mensagem mais sucinta
					strMsgErro = ex.Message;
					// No log em arquivo, grava o stack de erro completo
					Global.gravaLogAtividade(strOperacao + " - Falha!!\n" + ex.ToString());
					return false;
				}

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar inserir o registro de ocorrência para o boleto!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ boletoOcorrenciaInsere ]
		public static bool boletoOcorrenciaInsere(String usuario,
												 int id_arq_retorno,
												 int id_boleto_cedente,
												 int id_boleto,
												 int id_boleto_item,
												 String numeroDocumento,
												 String nossoNumero,
												 String digitoNossoNumero,
												 DateTime dataVencto,
												 decimal valorTitulo,
												 String identificacaoOcorrencia,
												 String motivosRejeicoes,
												 String motivoCodigoOcorrencia19,
												 DateTime dataOcorrencia,
												 String obsOcorrencia,
												 String linhaTextoRegistroArquivo,
												 ref String strMsgErro)
		{
			return boletoOcorrenciaInsere(usuario,
											id_arq_retorno,
											id_boleto_cedente,
											id_boleto,
											id_boleto_item,
											Global.Cte.FIN.StCampoFlag.FLAG_DESLIGADO,
											numeroDocumento,
											nossoNumero,
											digitoNossoNumero,
											dataVencto,
											valorTitulo,
											identificacaoOcorrencia,
											motivosRejeicoes,
											motivoCodigoOcorrencia19,
											dataOcorrencia,
											obsOcorrencia,
											linhaTextoRegistroArquivo,
											ref strMsgErro);
		}
		#endregion

		#region [ excluiBoletoEmStatusInicial ]
		public static bool excluiBoletoEmStatusInicial(String usuario,
													   int id_boleto,
													   ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Exclui boleto (somente se estiver no status inicial)";
			String strSql;
			bool blnSucesso = false;
			int intRetorno;
			SqlCommand cmComando;
			#endregion

			strMsgErro = "";
			try
			{
				cmComando = BD.criaSqlCommand();

				#region [ Exclui de t_FIN_BOLETO_ITEM_RATEIO ]
				strSql = "DELETE" +
						 " FROM t_FIN_BOLETO_ITEM_RATEIO" +
						 " WHERE" +
							" (id_boleto = " + id_boleto.ToString() + ")";
				cmComando.CommandText = strSql;
				intRetorno = BD.executaNonQuery(ref cmComando);
				#endregion

				#region [ Exclui de t_FIN_BOLETO_ITEM ]
				strSql = "DELETE" +
						 " FROM t_FIN_BOLETO_ITEM" +
						 " WHERE" +
							" (id_boleto = " + id_boleto.ToString() + ")";
				cmComando.CommandText = strSql;
				intRetorno = BD.executaNonQuery(ref cmComando);
				#endregion

				#region [ Exclui de t_FIN_BOLETO ]
				strSql = "DELETE" +
						 " FROM t_FIN_BOLETO" +
						 " WHERE" +
							" (id = " + id_boleto.ToString() + ")" +
							" AND (status = " + Global.Cte.FIN.CodBoletoStatus.INICIAL.ToString() + ")";
				cmComando.CommandText = strSql;
				intRetorno = BD.executaNonQuery(ref cmComando);
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar excluir o boleto (somente se estiver no status inicial)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ selecionaBoletosParaArqRemessa ]
		public static DataSet selecionaBoletosParaArqRemessa(short id_boleto_cedente)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataSet dsResultado = new DataSet();
			DsDataSource.DtbFinBoletoDataTable dtbFinBoleto = new DsDataSource.DtbFinBoletoDataTable();
			DsDataSource.DtbFinBoletoItemDataTable dtbFinBoletoItem = new DsDataSource.DtbFinBoletoItemDataTable();
			DsDataSource.DtbFinBoletoItemRateioDataTable dtbFinBoletoItemRateio = new DsDataSource.DtbFinBoletoItemRateioDataTable();
			DataRelation drlBoletoItem;
			DataRelation drlBoletoItemRateio;
			#endregion

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;

			strSql = "SELECT " +
						"*" +
					" FROM t_FIN_BOLETO" +
					" WHERE" +
						" (id_boleto_cedente = " + id_boleto_cedente.ToString() + ")" +
						" AND (id IN " +
								  "(" +
									"SELECT" +
										" DISTINCT id_boleto" +
									" FROM t_FIN_BOLETO_ITEM" +
									" WHERE" +
										" (status = " + Global.Cte.FIN.CodBoletoItemStatus.INICIAL.ToString() + ")" +
									")" +
							  ")" +
					" ORDER BY" +
						" id";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbFinBoleto);
			dsResultado.Tables.Add(dtbFinBoleto);

			strSql = "SELECT" +
						" tFBI.*" +
					" FROM t_FIN_BOLETO tFB" +
						" INNER JOIN t_FIN_BOLETO_ITEM tFBI" +
							" ON (tFB.id=tFBI.id_boleto)" +
					" WHERE" +
						" (tFBI.status = " + Global.Cte.FIN.CodBoletoItemStatus.INICIAL.ToString() + ")" +
						" AND (tFB.id_boleto_cedente = " + id_boleto_cedente.ToString() + ")" +
					" ORDER BY" +
						" tFB.id," +
						" tFBI.num_parcela";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbFinBoletoItem);
			dsResultado.Tables.Add(dtbFinBoletoItem);

			strSql = "SELECT" +
						" tFBIR.*" +
					" FROM t_FIN_BOLETO tFB" +
						" INNER JOIN t_FIN_BOLETO_ITEM tFBI" +
							" ON (tFB.id=tFBI.id_boleto)" +
						" INNER JOIN t_FIN_BOLETO_ITEM_RATEIO tFBIR" +
							" ON (tFBI.id=tFBIR.id_boleto_item)" +
					" WHERE" +
						" (tFBI.status = " + Global.Cte.FIN.CodBoletoItemStatus.INICIAL.ToString() + ")" +
						" AND (tFB.id_boleto_cedente = " + id_boleto_cedente.ToString() + ")" +
					" ORDER BY" +
						" tFB.id," +
						" tFBI.num_parcela," +
						" tFBIR.pedido";
			cmCommand.CommandText = strSql;
			daDataAdapter.Fill(dtbFinBoletoItemRateio);
			dsResultado.Tables.Add(dtbFinBoletoItemRateio);

			drlBoletoItem = new DataRelation("DtbFinBoleto_DtbFinBoletoItem", dsResultado.Tables["DtbFinBoleto"].Columns["id"], dsResultado.Tables["DtbFinBoletoItem"].Columns["id_boleto"]);
			dsResultado.Relations.Add(drlBoletoItem);
			drlBoletoItemRateio = new DataRelation("DtbFinBoletoItem_DtbFinBoletoItemRateio", dsResultado.Tables["DtbFinBoletoItem"].Columns["id"], dsResultado.Tables["DtbFinBoletoItemRateio"].Columns["id_boleto_item"]);
			dsResultado.Relations.Add(drlBoletoItemRateio);

			return dsResultado;
		}
		#endregion

		#region [ marcaBoletoEnviadoRemessaBanco ]
		public static bool marcaBoletoEnviadoRemessaBanco(String usuario,
														  int id,
														  int id_boleto_arq_remessa,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Marca o registro do boleto como gravado no arquivo de remessa";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoMarcaEnviadoRemessaBanco.Parameters["@id"].Value = id;
				cmBoletoMarcaEnviadoRemessaBanco.Parameters["@id_boleto_arq_remessa"].Value = id_boleto_arq_remessa;
				cmBoletoMarcaEnviadoRemessaBanco.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoMarcaEnviadoRemessaBanco);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar marcar o registro do boleto como já enviado no arquivo de remessa!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ marcaBoletoItemEnviadoRemessaBanco ]
		public static bool marcaBoletoItemEnviadoRemessaBanco(String usuario,
														  int id,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Marca o registro da parcela do boleto como gravado no arquivo de remessa";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemMarcaEnviadoRemessaBanco.Parameters["@id"].Value = id;
				cmBoletoItemMarcaEnviadoRemessaBanco.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemMarcaEnviadoRemessaBanco);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar marcar o registro da parcela do boleto como já enviado no arquivo de remessa!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ marcaBoletoCanceladoManual ]
		public static bool marcaBoletoCanceladoManual(String usuario,
													int id,
													ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Marca o registro do boleto como cancelado manualmente";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoMarcaCanceladoManual.Parameters["@id"].Value = id;
				cmBoletoMarcaCanceladoManual.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoMarcaCanceladoManual);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar marcar o registro do boleto como cancelado manualmente!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ marcaBoletoItemCanceladoManual ]
		public static bool marcaBoletoItemCanceladoManual(String usuario,
														int id,
														ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Marca o registro da parcela do boleto como cancelado manualmente";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemMarcaCanceladoManual.Parameters["@id"].Value = id;
				cmBoletoItemMarcaCanceladoManual.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemMarcaCanceladoManual);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar marcar o registro da parcela do boleto como cancelado manualmente!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ marcaBoletoItemCanceladoManualByIdBoleto ]
		public static bool marcaBoletoItemCanceladoManualByIdBoleto(String usuario,
														int id_boleto,
														ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Marca os registros das parcelas de uma série de boletos como cancelado manualmente";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemMarcaCanceladoManualByIdBoleto.Parameters["@id_boleto"].Value = id_boleto;
				cmBoletoItemMarcaCanceladoManualByIdBoleto.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemMarcaCanceladoManualByIdBoleto);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno > 0)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar marcar os registros das parcelas de uma série de boletos como cancelado manualmente!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ boletoArqRemessaInsere ]
		/// <summary>
		/// Grava o registro em t_FIN_BOLETO_ARQ_REMESSA para manter o histórico dos arquivos de remessa gerados.
		/// </summary>
		/// <param name="usuario">
		/// Usuário que gerou o arquivo de remessa.
		/// </param>
		/// <param name="boletoArqRemessa">
		/// Objeto do tipo BoletoArqRemessa com os dados básicos do arquivo de remessa gerado.
		/// </param>
		/// <param name="strMsgErro">
		/// Em caso de erro, retorna mensagem com descrição.
		/// </param>
		/// <returns>
		/// true: gravação bem sucedida.
		/// false: falha na gravação.
		/// </returns>
		public static bool boletoArqRemessaInsere(String usuario,
												  BoletoArqRemessa boletoArqRemessa,
												  ref String strMsgErro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			bool blnGerouNsu;
			int intNsuBoletoArqRemessa = 0;
			int intRetorno;
			String strOperacao = "Gravação de histórico de arquivos de remessa";
			#endregion

			try
			{
				strMsgErro = "";

				#region [ Gera o NSU? ]
				if (boletoArqRemessa.id <= 0)
				{
					blnGerouNsu = BD.geraNsu(Global.Cte.FIN.NSU.T_FIN_BOLETO_ARQ_REMESSA, ref intNsuBoletoArqRemessa, ref strMsgErro);
					if (!blnGerouNsu)
					{
						strMsgErro = "Falha ao tentar gerar o NSU para o registro de histórico de arquivos de remessa!!\n" + strMsgErro;
						return false;
					}
					boletoArqRemessa.id = intNsuBoletoArqRemessa;
				}
				#endregion

				try
				{
					#region [ Tenta gravar os dados ]

					#region [ Preenche o valor dos parâmetros ]
					cmBoletoArqRemessaInsert.Parameters["@id"].Value = boletoArqRemessa.id;
					cmBoletoArqRemessaInsert.Parameters["@nsu_arq_remessa"].Value = boletoArqRemessa.nsu_arq_remessa;
					cmBoletoArqRemessaInsert.Parameters["@usuario_geracao"].Value = usuario;
					cmBoletoArqRemessaInsert.Parameters["@qtde_registros"].Value = boletoArqRemessa.qtde_registros;
					cmBoletoArqRemessaInsert.Parameters["@qtde_serie_boletos"].Value = boletoArqRemessa.qtde_serie_boletos;
					cmBoletoArqRemessaInsert.Parameters["@id_boleto_cedente"].Value = boletoArqRemessa.id_boleto_cedente;
					cmBoletoArqRemessaInsert.Parameters["@codigo_empresa"].Value = boletoArqRemessa.codigo_empresa;
					cmBoletoArqRemessaInsert.Parameters["@nome_empresa"].Value = boletoArqRemessa.nome_empresa;
					cmBoletoArqRemessaInsert.Parameters["@num_banco"].Value = boletoArqRemessa.num_banco;
					cmBoletoArqRemessaInsert.Parameters["@nome_banco"].Value = boletoArqRemessa.nome_banco;
					cmBoletoArqRemessaInsert.Parameters["@agencia"].Value = boletoArqRemessa.agencia;
					cmBoletoArqRemessaInsert.Parameters["@digito_agencia"].Value = boletoArqRemessa.digito_agencia;
					cmBoletoArqRemessaInsert.Parameters["@conta"].Value = boletoArqRemessa.conta;
					cmBoletoArqRemessaInsert.Parameters["@digito_conta"].Value = boletoArqRemessa.digito_conta;
					cmBoletoArqRemessaInsert.Parameters["@carteira"].Value = boletoArqRemessa.carteira;
					cmBoletoArqRemessaInsert.Parameters["@vl_total"].Value = boletoArqRemessa.vl_total;
					cmBoletoArqRemessaInsert.Parameters["@duracao_proc_em_seg"].Value = boletoArqRemessa.duracao_proc_em_seg;
					cmBoletoArqRemessaInsert.Parameters["@nome_arq_remessa"].Value = boletoArqRemessa.nome_arq_remessa;
					cmBoletoArqRemessaInsert.Parameters["@caminho_arq_remessa"].Value = boletoArqRemessa.caminho_arq_remessa;
					cmBoletoArqRemessaInsert.Parameters["@st_geracao"].Value = boletoArqRemessa.st_geracao;
					cmBoletoArqRemessaInsert.Parameters["@msg_erro_geracao"].Value = boletoArqRemessa.msg_erro_geracao;
					#endregion

					#region [ Tenta inserir o registro ]
					strMsgErro = "";
					try
					{
						intRetorno = BD.executaNonQuery(ref cmBoletoArqRemessaInsert);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						strMsgErro = ex.Message;
						Global.gravaLogAtividade(strOperacao + " - Exception!!\n" + ex.ToString());
					}
					#endregion

					#region [ Gravou o registro? ]
					if (intRetorno == 0)
					{
						strMsgErro = "Falha ao tentar gravar o registro de histórico de arquivos de remessa!!\n" + strMsgErro;
						return false;
					}
					#endregion

					#endregion

					blnSucesso = true;
				}
				catch (Exception ex)
				{
					// Para o usuário, exibe uma mensagem mais sucinta
					strMsgErro = ex.Message;
					// No log em arquivo, grava o stack de erro completo
					Global.gravaLogAtividade(strOperacao + " - Falha!!\n" + ex.ToString());
					return false;
				}

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar gravar o registro do histórico de arquivos de remessa!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ boletoArqRetornoInsere ]
		/// <summary>
		/// Grava o registro em t_FIN_BOLETO_ARQ_RETORNO para manter o histórico dos arquivos de retorno carregados.
		/// </summary>
		/// <param name="usuario">
		/// Usuário que carregou o arquivo de retorno.
		/// </param>
		/// <param name="boletoArqRetorno">
		/// Objeto do tipo BoletoArqRetorno com os dados básicos do arquivo de retorno.
		/// </param>
		/// <param name="strMsgErro">
		/// Em caso de erro, retorna mensagem com descrição.
		/// </param>
		/// <returns>
		/// true: gravação bem sucedida.
		/// false: falha na gravação.
		/// </returns>
		public static bool boletoArqRetornoInsere(String usuario,
												  BoletoArqRetorno boletoArqRetorno,
												  ref String strMsgErro)
		{
			#region [ Declarações ]
			bool blnSucesso = false;
			int intRetorno;
			String strOperacao = "Gravação de histórico de arquivos de retorno";
			#endregion

			try
			{
				strMsgErro = "";

				#region [ Consistência ]
				if (boletoArqRetorno.id <= 0)
				{
					strMsgErro = "NSU não fornecido para o novo registro da tabela t_FIN_BOLETO_ARQ_RETORNO!!\n" + strMsgErro;
					return false;
				}
				#endregion

				try
				{
					#region [ Tenta gravar os dados ]

					#region [ Preenche o valor dos parâmetros ]
					cmBoletoArqRetornoInsert.Parameters["@id"].Value = boletoArqRetorno.id;
					cmBoletoArqRetornoInsert.Parameters["@id_boleto_cedente"].Value = boletoArqRetorno.id_boleto_cedente;
					cmBoletoArqRetornoInsert.Parameters["@usuario_processamento"].Value = usuario;
					cmBoletoArqRetornoInsert.Parameters["@qtde_registros"].Value = boletoArqRetorno.qtde_registros;
					cmBoletoArqRetornoInsert.Parameters["@codigo_empresa"].Value = boletoArqRetorno.codigo_empresa.Trim();
					cmBoletoArqRetornoInsert.Parameters["@nome_empresa"].Value = boletoArqRetorno.nome_empresa.Trim();
					cmBoletoArqRetornoInsert.Parameters["@num_banco"].Value = boletoArqRetorno.num_banco;
					cmBoletoArqRetornoInsert.Parameters["@nome_banco"].Value = boletoArqRetorno.nome_banco.Trim();
					cmBoletoArqRetornoInsert.Parameters["@data_gravacao_arquivo"].Value = boletoArqRetorno.data_gravacao_arquivo;
					cmBoletoArqRetornoInsert.Parameters["@dt_gravacao_arquivo"].Value = Global.formataDataYyyyMmDdComSeparador(Global.converteDdMmYyParaDateTime(boletoArqRetorno.data_gravacao_arquivo));
					cmBoletoArqRetornoInsert.Parameters["@numero_aviso_bancario"].Value = boletoArqRetorno.numero_aviso_bancario;
					cmBoletoArqRetornoInsert.Parameters["@data_credito"].Value = boletoArqRetorno.data_credito;
					cmBoletoArqRetornoInsert.Parameters["@dt_credito"].Value = Global.formataDataYyyyMmDdComSeparador(Global.converteDdMmYyParaDateTime(boletoArqRetorno.data_credito));
					cmBoletoArqRetornoInsert.Parameters["@qtdeTitulosEmCobranca"].Value = boletoArqRetorno.qtdeTitulosEmCobranca;
					cmBoletoArqRetornoInsert.Parameters["@valorTotalEmCobranca"].Value = boletoArqRetorno.valorTotalEmCobranca;
					cmBoletoArqRetornoInsert.Parameters["@qtdeRegsOcorrencia02ConfirmacaoEntradas"].Value = boletoArqRetorno.qtdeRegsOcorrencia02ConfirmacaoEntradas;
					cmBoletoArqRetornoInsert.Parameters["@valorRegsOcorrencia02ConfirmacaoEntradas"].Value = boletoArqRetorno.valorRegsOcorrencia02ConfirmacaoEntradas;
					cmBoletoArqRetornoInsert.Parameters["@valorRegsOcorrencia06Liquidacao"].Value = boletoArqRetorno.valorRegsOcorrencia06Liquidacao;
					cmBoletoArqRetornoInsert.Parameters["@qtdeRegsOcorrencia06Liquidacao"].Value = boletoArqRetorno.qtdeRegsOcorrencia06Liquidacao;
					cmBoletoArqRetornoInsert.Parameters["@valorRegsOcorrencia06"].Value = boletoArqRetorno.valorRegsOcorrencia06;
					cmBoletoArqRetornoInsert.Parameters["@qtdeRegsOcorrencia09e10TitulosBaixados"].Value = boletoArqRetorno.qtdeRegsOcorrencia09e10TitulosBaixados;
					cmBoletoArqRetornoInsert.Parameters["@valorRegsOcorrencia09e10TitulosBaixados"].Value = boletoArqRetorno.valorRegsOcorrencia09e10TitulosBaixados;
					cmBoletoArqRetornoInsert.Parameters["@qtdeRegsOcorrencia13AbatimentoCancelado"].Value = boletoArqRetorno.qtdeRegsOcorrencia13AbatimentoCancelado;
					cmBoletoArqRetornoInsert.Parameters["@valorRegsOcorrencia13AbatimentoCancelado"].Value = boletoArqRetorno.valorRegsOcorrencia13AbatimentoCancelado;
					cmBoletoArqRetornoInsert.Parameters["@qtdeRegsOcorrencia14VenctoAlterado"].Value = boletoArqRetorno.qtdeRegsOcorrencia14VenctoAlterado;
					cmBoletoArqRetornoInsert.Parameters["@valorRegsOcorrencia14VenctoAlterado"].Value = boletoArqRetorno.valorRegsOcorrencia14VenctoAlterado;
					cmBoletoArqRetornoInsert.Parameters["@qtdeRegsOcorrencia12AbatimentoConcedido"].Value = boletoArqRetorno.qtdeRegsOcorrencia12AbatimentoConcedido;
					cmBoletoArqRetornoInsert.Parameters["@valorRegsOcorrencia12AbatimentoConcedido"].Value = boletoArqRetorno.valorRegsOcorrencia12AbatimentoConcedido;
					cmBoletoArqRetornoInsert.Parameters["@qtdeRegsOcorrencia19ConfirmacaoInstrucaoProtesto"].Value = boletoArqRetorno.qtdeRegsOcorrencia19ConfirmacaoInstrucaoProtesto;
					cmBoletoArqRetornoInsert.Parameters["@valorRegsOcorrencia19ConfirmacaoInstrucaoProtesto"].Value = boletoArqRetorno.valorRegsOcorrencia19ConfirmacaoInstrucaoProtesto;
					cmBoletoArqRetornoInsert.Parameters["@valorTotalRateiosEfetuados"].Value = boletoArqRetorno.valorTotalRateiosEfetuados;
					cmBoletoArqRetornoInsert.Parameters["@qtdeTotalRateiosEfetuados"].Value = boletoArqRetorno.qtdeTotalRateiosEfetuados;
					cmBoletoArqRetornoInsert.Parameters["@duracao_proc_em_seg"].Value = boletoArqRetorno.duracao_proc_em_seg;
					cmBoletoArqRetornoInsert.Parameters["@nome_arq_retorno"].Value = boletoArqRetorno.nome_arq_retorno;
					cmBoletoArqRetornoInsert.Parameters["@caminho_arq_retorno"].Value = boletoArqRetorno.caminho_arq_retorno;
					cmBoletoArqRetornoInsert.Parameters["@st_processamento"].Value = boletoArqRetorno.st_processamento;
					cmBoletoArqRetornoInsert.Parameters["@msg_erro_processamento"].Value = boletoArqRetorno.msg_erro_processamento;
					#endregion

					#region [ Tenta inserir o registro ]
					strMsgErro = "";
					try
					{
						intRetorno = BD.executaNonQuery(ref cmBoletoArqRetornoInsert);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						strMsgErro = ex.Message;
						Global.gravaLogAtividade(strOperacao + " - Exception!!\n" + ex.ToString());
					}
					#endregion

					#region [ Gravou o registro? ]
					if (intRetorno == 0)
					{
						strMsgErro = "Falha ao tentar gravar o registro de histórico de arquivos de retorno!!\n" + strMsgErro;
						return false;
					}
					#endregion

					#endregion

					blnSucesso = true;
				}
				catch (Exception ex)
				{
					// Para o usuário, exibe uma mensagem mais sucinta
					strMsgErro = ex.Message;
					// No log em arquivo, grava o stack de erro completo
					Global.gravaLogAtividade(strOperacao + " - Falha!!\n" + ex.ToString());
					return false;
				}

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar gravar o registro do histórico de arquivos de retorno!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ boletoArqRetornoAtualiza ]
		public static bool boletoArqRetornoAtualiza(String usuario,
													int idArqRetorno,
													short stProcessamento,
													int duracaoProcessamentoEmSegundos,
													String msgErroProcessamento,
													ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro com os dados do arquivo de retorno";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoArqRetornoUpdate.Parameters["@id"].Value = idArqRetorno;
				cmBoletoArqRetornoUpdate.Parameters["@st_processamento"].Value = stProcessamento;
				cmBoletoArqRetornoUpdate.Parameters["@duracao_proc_em_seg"].Value = duracaoProcessamentoEmSegundos;
				cmBoletoArqRetornoUpdate.Parameters["@msg_erro_processamento"].Value = Texto.leftStr(msgErroProcessamento, 1024);
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoArqRetornoUpdate);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar o registro com os dados do arquivo de retorno!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ boletoArqRetornoJaCarregado ]
		/// <summary>
		/// O nome do arquivo de retorno possui o formato CBDDMMNN.RET, sendo que NN é apenas um índice, portanto,
		/// o nome do arquivo de retorno pode se repetir a cada ano.
		/// </summary>
		/// <param name="nomeArqRetorno">Nome do arquivo de retorno, sem o path</param>
		/// <param name="codigoEmpresa">Código da empresa fornecido pelo banco</param>
		/// <param name="dataGravacaoArquivoDdMmYy">Informação que consta no header do arquivo de retorno.</param>
		/// <param name="dtHrProcessamentoAnterior">No caso do arquivo já ter sido carregado, informa a data e hora em que isso ocorreu</param>
		/// <param name="usuarioProcessamentoAnterior">No caso do arquivo já ter sido carregado, informa o usuário que realizou a carga</param>
		/// <returns>
		/// true: o arquivo de retorno já foi carregado com sucesso anteriormente
		/// false: o arquivo de retorno ainda não foi carregado
		/// </returns>
		public static bool boletoArqRetornoJaCarregado(String nomeArqRetorno,
													   String codigoEmpresa,
													   String dataGravacaoArquivoDdMmYy,
													   out DateTime dtHrProcessamentoAnterior,
													   out String usuarioProcessamentoAnterior)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Inicialização ]
			dtHrProcessamentoAnterior = DateTime.MinValue;
			usuarioProcessamentoAnterior = "";
			#endregion

			#region [ Consistências ]
			if (nomeArqRetorno == null) throw new FinanceiroException("O nome do arquivo de retorno não foi fornecido!!");
			if (nomeArqRetorno.Trim().Length == 0) throw new FinanceiroException("O nome do arquivo de retorno não foi informado!!");

			if (codigoEmpresa == null) throw new FinanceiroException("O código da empresa não foi fornecido!!");
			if (codigoEmpresa.Trim().Length == 0) throw new FinanceiroException("O código da empresa não foi informado!!");

			if (dataGravacaoArquivoDdMmYy == null) throw new FinanceiroException("A data da gravação do arquivo não foi fornecida!!");
			if (dataGravacaoArquivoDdMmYy.Trim().Length == 0) throw new FinanceiroException("A data da gravação do arquivo não foi informada!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Executa a consulta ]
			// O nome do arquivo de retorno é no formato CBDDMMNN.RET, sendo que NN é apenas um índice, portanto, o nome pode se
			// repetir a cada ano. Além disso, empresas cedentes diferentes provavelmente terão arquivos de retorno com mesmo
			// nome diariamente.
			strSql = "SELECT" +
						" codigo_empresa," +
						" data_gravacao_arquivo," +
						" dt_hr_processamento," +
						" usuario_processamento" +
					" FROM t_FIN_BOLETO_ARQ_RETORNO" +
					" WHERE" +
						" (codigo_empresa = '" + codigoEmpresa + "')" +
						" AND (nome_arq_retorno = '" + nomeArqRetorno.Trim().ToUpper() + "')" +
						" AND (data_gravacao_arquivo = '" + dataGravacaoArquivoDdMmYy + "')" +
						" AND (st_processamento = " + Global.Cte.FIN.CodBoletoArqRetornoStProcessamento.SUCESSO.ToString() + ")";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);
			#endregion

			#region [ O arquivo já foi carregado anteriormente? ]
			if (dtbResultado.Rows.Count == 0) return false;
			#endregion

			#region [ Retorna dados do processamento anterior ]
			rowResultado = dtbResultado.Rows[0];
			dtHrProcessamentoAnterior = BD.readToDateTime(rowResultado["dt_hr_processamento"]);
			usuarioProcessamentoAnterior = BD.readToString(rowResultado["usuario_processamento"]);
			#endregion

			return true;
		}
		#endregion

		#region [ boletoArqRetornoObtemDtGravacaoUltArqCarregadoComSucesso ]
		public static bool boletoArqRetornoObtemDtGravacaoUltArqCarregadoComSucesso(
										String codigoEmpresa,
										out DateTime dtGravacaoUltArqCarregadoComSucesso,
										out String nomeUltArqRetornoCarregadoComSucesso,
										out DateTime dtHrProcessamentoUltArqCarregadoComSucesso,
										out String usuarioProcessamentoUltArqCarregadoComSucesso)
		{
			#region [ Declarações ]
			String strSql;
			String strDtGravacaoUltArqCarregadoComSucesso;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Inicialização ]
			dtGravacaoUltArqCarregadoComSucesso = DateTime.MinValue;
			nomeUltArqRetornoCarregadoComSucesso = "";
			dtHrProcessamentoUltArqCarregadoComSucesso = DateTime.MinValue;
			usuarioProcessamentoUltArqCarregadoComSucesso = "";
			#endregion

			#region [ Consistências ]
			if (codigoEmpresa == null) throw new FinanceiroException("O código da empresa não foi fornecido!!");
			if (codigoEmpresa.Trim().Length == 0) throw new FinanceiroException("O código da empresa não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Executa a consulta ]
			strSql = "SELECT TOP 1" +
						" data_gravacao_arquivo," +
						" nome_arq_retorno," +
						" dt_hr_processamento," +
						" usuario_processamento" +
					" FROM t_FIN_BOLETO_ARQ_RETORNO" +
					" WHERE" +
						" (codigo_empresa = '" + codigoEmpresa.Trim() + "')" +
						" AND (st_processamento = " + Global.Cte.FIN.CodBoletoArqRetornoStProcessamento.SUCESSO.ToString() + ")" +
					" ORDER BY" +
						" dt_processamento DESC";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);
			#endregion

			#region [ Nenhum arquivo encontrado ]
			if (dtbResultado.Rows.Count == 0) return false;
			#endregion

			#region [ Analisa o último arquivo processado ]
			rowResultado = dtbResultado.Rows[0];

			nomeUltArqRetornoCarregadoComSucesso = BD.readToString(rowResultado["nome_arq_retorno"]);

			#region [ Data está no formato esperado? ]
			strDtGravacaoUltArqCarregadoComSucesso = Global.digitos(BD.readToString(rowResultado["data_gravacao_arquivo"]));
			if (strDtGravacaoUltArqCarregadoComSucesso.Length != 6) return false;
			dtGravacaoUltArqCarregadoComSucesso = Global.converteDdMmYyParaDateTime(strDtGravacaoUltArqCarregadoComSucesso);
			#endregion

			dtHrProcessamentoUltArqCarregadoComSucesso = BD.readToDateTime(rowResultado["dt_hr_processamento"]);
			usuarioProcessamentoUltArqCarregadoComSucesso = BD.readToString(rowResultado["usuario_processamento"]);
			#endregion

			return true;
		}
		#endregion

		#region [ boletoArqRetornoObtemUltimaDtCredito ]
		public static bool boletoArqRetornoObtemUltimaDtCredito(
										out DateTime dtCreditoArqRetorno,
										out String nomeArqRetorno,
										out DateTime dtHrProcessamentoArqRetorno,
										out String usuarioProcessamentoArqRetorno)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Inicialização ]
			dtCreditoArqRetorno = DateTime.MinValue;
			nomeArqRetorno = "";
			dtHrProcessamentoArqRetorno = DateTime.MinValue;
			usuarioProcessamentoArqRetorno = "";
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Executa a consulta ]
			strSql = "SELECT TOP 1" +
						" dt_credito," +
						" nome_arq_retorno," +
						" dt_hr_processamento," +
						" usuario_processamento" +
					" FROM t_FIN_BOLETO_ARQ_RETORNO" +
					" WHERE" +
						" (st_processamento = " + Global.Cte.FIN.CodBoletoArqRetornoStProcessamento.SUCESSO.ToString() + ")" +
						" AND (dt_credito IS NOT NULL)" +
					" ORDER BY" +
						" dt_credito DESC";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);
			#endregion

			#region [ Nenhum arquivo encontrado ]
			if (dtbResultado.Rows.Count == 0) return false;
			#endregion

			#region [ Analisa o último arquivo processado ]
			rowResultado = dtbResultado.Rows[0];

			dtCreditoArqRetorno = BD.readToDateTime(rowResultado["dt_credito"]);
			nomeArqRetorno = BD.readToString(rowResultado["nome_arq_retorno"]);
			dtHrProcessamentoArqRetorno = BD.readToDateTime(rowResultado["dt_hr_processamento"]);
			usuarioProcessamentoArqRetorno = BD.readToString(rowResultado["usuario_processamento"]);
			#endregion

			return true;
		}
		#endregion

		#region [ boletoArqRetornoObtemUltimaDtGravacaoArquivo ]
		public static bool boletoArqRetornoObtemUltimaDtGravacaoArquivo(
										out DateTime dtGravacaoArquivoArqRetorno,
										out String nomeArqRetorno,
										out DateTime dtHrProcessamentoArqRetorno,
										out String usuarioProcessamentoArqRetorno)
		{
			#region [ Declarações ]
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Inicialização ]
			dtGravacaoArquivoArqRetorno = DateTime.MinValue;
			nomeArqRetorno = "";
			dtHrProcessamentoArqRetorno = DateTime.MinValue;
			usuarioProcessamentoArqRetorno = "";
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Executa a consulta ]
			strSql = "SELECT TOP 1" +
						" dt_gravacao_arquivo," +
						" nome_arq_retorno," +
						" dt_hr_processamento," +
						" usuario_processamento" +
					" FROM t_FIN_BOLETO_ARQ_RETORNO" +
					" WHERE" +
						" (st_processamento = " + Global.Cte.FIN.CodBoletoArqRetornoStProcessamento.SUCESSO.ToString() + ")" +
						" AND (dt_gravacao_arquivo IS NOT NULL)" +
					" ORDER BY" +
						" dt_gravacao_arquivo DESC";
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);
			#endregion

			#region [ Nenhum arquivo encontrado ]
			if (dtbResultado.Rows.Count == 0) return false;
			#endregion

			#region [ Analisa o último arquivo processado ]
			rowResultado = dtbResultado.Rows[0];

			dtGravacaoArquivoArqRetorno = BD.readToDateTime(rowResultado["dt_gravacao_arquivo"]);
			nomeArqRetorno = BD.readToString(rowResultado["nome_arq_retorno"]);
			dtHrProcessamentoArqRetorno = BD.readToDateTime(rowResultado["dt_hr_processamento"]);
			usuarioProcessamentoArqRetorno = BD.readToString(rowResultado["usuario_processamento"]);
			#endregion

			return true;
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia02EntradaConfirmada ]
		public static bool atualizaBoletoItemOcorrencia02EntradaConfirmada(
								String usuario,
								int idBoletoItem,
								String nossoNumero,
								String digitoNossoNumero,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								String motivoCodigoOcorrencia19,
								DateTime dataOcorrenciaBanco,
								String codigoBarras,
								String linhaDigitavel,
								decimal vlTarifaRegistro,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto devido a ocorrência 02 (entrada confirmada)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrencia02.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrencia02.Parameters["@nosso_numero"].Value = nossoNumero;
				cmBoletoItemAtualizaOcorrencia02.Parameters["@digito_nosso_numero"].Value = digitoNossoNumero;
				cmBoletoItemAtualizaOcorrencia02.Parameters["@codigo_barras"].Value = codigoBarras;
				cmBoletoItemAtualizaOcorrencia02.Parameters["@linha_digitavel"].Value = linhaDigitavel;
				cmBoletoItemAtualizaOcorrencia02.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrencia02.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrencia02.Parameters["@ult_motivo_ocorrencia_19"].Value = motivoCodigoOcorrencia19;
				cmBoletoItemAtualizaOcorrencia02.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia02.Parameters["@dt_entrada_confirmada"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia02.Parameters["@vl_tarifa_registro"].Value = vlTarifaRegistro;
				cmBoletoItemAtualizaOcorrencia02.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrencia02);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 02 (entrada confirmada)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia06LiquidacaoNormal ]
		public static bool atualizaBoletoItemOcorrencia06LiquidacaoNormal(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								DateTime dataOcorrenciaBanco,
								decimal vlAbatimentoConcedido,
								decimal vlDescontoConcedido,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto devido a ocorrência 06 (liquidação normal)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrencia06.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrencia06.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrencia06.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrencia06.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia06.Parameters["@vl_abatimento_concedido"].Value = vlAbatimentoConcedido;
				cmBoletoItemAtualizaOcorrencia06.Parameters["@vl_desconto_concedido"].Value = vlDescontoConcedido;
				cmBoletoItemAtualizaOcorrencia06.Parameters["@st_boleto_ocorrencia_06"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmBoletoItemAtualizaOcorrencia06.Parameters["@dt_ocorrencia_banco_boleto_ocorrencia_06"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia06.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrencia06);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 06 (liquidação normal)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia09BaixadoAutoViaArq ]
		public static bool atualizaBoletoItemOcorrencia09BaixadoAutoViaArq(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								DateTime dataOcorrenciaBanco,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto devido a ocorrência 09 (baixado automaticamente via arquivo)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrencia09.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrencia09.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrencia09.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrencia09.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia09.Parameters["@st_boleto_baixado"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmBoletoItemAtualizaOcorrencia09.Parameters["@dt_ocorrencia_banco_boleto_baixado"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia09.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrencia09);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 09 (baixado automaticamente via arquivo)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia10BaixadoConfInstrAgencia ]
		public static bool atualizaBoletoItemOcorrencia10BaixadoConfInstrAgencia(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								DateTime dataOcorrenciaBanco,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto devido a ocorrência 10 (baixado conforme instruções da agência)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrencia10.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrencia10.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrencia10.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrencia10.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia10.Parameters["@st_boleto_baixado"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmBoletoItemAtualizaOcorrencia10.Parameters["@dt_ocorrencia_banco_boleto_baixado"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia10.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrencia10);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 10 (baixado conforme instruções da agência)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia12AbatimentoConcedido ]
		public static bool atualizaBoletoItemOcorrencia12AbatimentoConcedido(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								DateTime dataOcorrenciaBanco,
								decimal vlAbatimentoConcedido,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto devido a ocorrência 12 (abatimento concedido)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrencia12.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrencia12.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrencia12.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrencia12.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia12.Parameters["@vl_abatimento_concedido"].Value = vlAbatimentoConcedido;
				cmBoletoItemAtualizaOcorrencia12.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrencia12);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 12 (abatimento concedido)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia13AbatimentoCancelado ]
		public static bool atualizaBoletoItemOcorrencia13AbatimentoCancelado(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								DateTime dataOcorrenciaBanco,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto devido a ocorrência 13 (abatimento cancelado)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrencia13.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrencia13.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrencia13.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrencia13.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia13.Parameters["@vl_abatimento_concedido"].Value = 0m;
				cmBoletoItemAtualizaOcorrencia13.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrencia13);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 13 (abatimento cancelado)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia14VenctoAlterado ]
		public static bool atualizaBoletoItemOcorrencia14VenctoAlterado(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								DateTime dataOcorrenciaBanco,
								DateTime dtNovoVencto,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto devido a ocorrência 14 (vencimento alterado)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrencia14.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrencia14.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrencia14.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrencia14.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia14.Parameters["@dt_vencto"].Value = Global.formataDataYyyyMmDdComSeparador(dtNovoVencto);
				cmBoletoItemAtualizaOcorrencia14.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrencia14);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 14 (vencimento alterado)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia15LiquidacaoEmCartorio ]
		public static bool atualizaBoletoItemOcorrencia15LiquidacaoEmCartorio(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								DateTime dataOcorrenciaBanco,
								decimal vlAbatimentoConcedido,
								decimal vlDescontoConcedido,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto devido a ocorrência 15 (liquidação em cartório)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrencia15.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrencia15.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrencia15.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrencia15.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia15.Parameters["@vl_abatimento_concedido"].Value = vlAbatimentoConcedido;
				cmBoletoItemAtualizaOcorrencia15.Parameters["@vl_desconto_concedido"].Value = vlDescontoConcedido;
				cmBoletoItemAtualizaOcorrencia15.Parameters["@st_boleto_ocorrencia_15"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmBoletoItemAtualizaOcorrencia15.Parameters["@dt_ocorrencia_banco_boleto_ocorrencia_15"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia15.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrencia15);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 15 (liquidação em cartório)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia16TituloPagoEmCheque ]
		public static bool atualizaBoletoItemOcorrencia16TituloPagoEmCheque(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								DateTime dataOcorrenciaBanco,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto devido a ocorrência 16 (título pago em cheque)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrencia16.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrencia16.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrencia16.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrencia16.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia16.Parameters["@st_boleto_pago_cheque"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmBoletoItemAtualizaOcorrencia16.Parameters["@dt_ocorrencia_banco_boleto_pago_cheque"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia16.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrencia16);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 16 (título pago em cheque)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia17LiqAposBaixaOuTitNaoRegistrado ]
		public static bool atualizaBoletoItemOcorrencia17LiqAposBaixaOuTitNaoRegistrado(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								DateTime dataOcorrenciaBanco,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto devido a ocorrência 17 (liquidação após baixa ou título não registrado)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrencia17.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrencia17.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrencia17.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrencia17.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia17.Parameters["@st_boleto_ocorrencia_17"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmBoletoItemAtualizaOcorrencia17.Parameters["@dt_ocorrencia_banco_boleto_ocorrencia_17"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia17.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrencia17);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 17 (liquidação após baixa ou título não registrado)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia19ConfirmacaoRecebInstProtesto ]
		public static bool atualizaBoletoItemOcorrencia19ConfirmacaoRecebInstProtesto(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								String motivoCodigoOcorrencia19,
								DateTime dataOcorrenciaBanco,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto devido a ocorrência 19 (confirmação receb. inst. de protesto)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrencia19.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrencia19.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrencia19.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrencia19.Parameters["@ult_motivo_ocorrencia_19"].Value = motivoCodigoOcorrencia19;
				cmBoletoItemAtualizaOcorrencia19.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia19.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrencia19);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 19 (confirmação receb. inst. de protesto)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia22TituloComPagamentoCancelado ]
		public static bool atualizaBoletoItemOcorrencia22TituloComPagamentoCancelado(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								DateTime dataOcorrenciaBanco,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto devido a ocorrência 22 (título com pagamento cancelado)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrencia22.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrencia22.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrencia22.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrencia22.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia22.Parameters["@st_boleto_pago_cheque"].Value = Global.Cte.FIN.StCampoFlag.FLAG_DESLIGADO;
				cmBoletoItemAtualizaOcorrencia22.Parameters["@dt_ocorrencia_banco_boleto_pago_cheque"].Value = "";
				cmBoletoItemAtualizaOcorrencia22.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrencia22);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 22 (título com pagamento cancelado)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia23EntradaTituloEmCartorio ]
		public static bool atualizaBoletoItemOcorrencia23EntradaTituloEmCartorio(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								DateTime dataOcorrenciaBanco,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto devido a ocorrência 23 (entrada do título em cartório)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrencia23.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrencia23.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrencia23.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrencia23.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia23.Parameters["@st_boleto_ocorrencia_23"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmBoletoItemAtualizaOcorrencia23.Parameters["@dt_ocorrencia_banco_boleto_ocorrencia_23"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia23.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrencia23);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 23 (entrada do título em cartório)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia28DebitoTarifasCustas ]
		public static bool atualizaBoletoItemOcorrencia28DebitoTarifasCustas(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								DateTime dataOcorrenciaBanco,
								LinhaRegistroTipo1ArquivoRetorno linhaRegistro,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto devido a ocorrência 28 (débito de tarifas/custas)";
			bool blnSucesso = false;
			int intRetorno;
			String strClausulaSet = "";
			String strSql;
			SqlCommand cmCommand;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Monta a cláusula SET do SQL da atualização ]
				if (Global.existeMotivoOcorrencia(linhaRegistro.motivosRejeicoes.valor, "03"))
				{
					#region [ Tarifa de sustação (motivo 03) usando campo despesas de cobrança ]
					if (strClausulaSet.Length > 0) strClausulaSet += ",";
					strClausulaSet += " vl_tarifa_sustacao = " + Global.sqlFormataDecimal(Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor));
					#endregion
				}
				else if (Global.existeMotivoOcorrencia(linhaRegistro.motivosRejeicoes.valor, "04"))
				{
					#region [ Tarifa de protesto (motivo 04) usando campo despesas de cobrança ]
					if (strClausulaSet.Length > 0) strClausulaSet += ",";
					strClausulaSet += " vl_tarifa_protesto = " + Global.sqlFormataDecimal(Global.decodificaCampoMonetario(linhaRegistro.valorDespesasCobranca.valor));
					#endregion
				}

				#region [ Custas de protesto (motivo 08) usando campo outras despesas ]
				if (Global.existeMotivoOcorrencia(linhaRegistro.motivosRejeicoes.valor, "08"))
				{
					if (strClausulaSet.Length > 0) strClausulaSet += ",";
					strClausulaSet += " vl_custas_protesto = " + Global.sqlFormataDecimal(Global.decodificaCampoMonetario(linhaRegistro.valorOutrasDespesas.valor));
				}
				#endregion

				#endregion

				#region [ Há atualizações para fazer? ]
				if (strClausulaSet.Length == 0) return true;
				#endregion

				#region [ Completa a cláusula SET com o preenchimento dos campos complementares ]
				if (strClausulaSet.Length > 0) strClausulaSet += ",";
				strClausulaSet += " ult_identificacao_ocorrencia = '" + identificacaoOcorrencia + "'," +
								  " ult_motivos_rejeicoes = '" + motivosRejeicoes + "'," +
								  " ult_data_ocorrencia_banco = " + Global.sqlMontaDateTimeParaSqlDateTime(dataOcorrenciaBanco) + "," +
								  " ult_data_carga_arq_retorno = " + Global.sqlMontaGetdateSomenteData() + "," +
								  " dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + "," +
								  " dt_hr_ult_atualizacao = getdate()," +
								  " usuario_ult_atualizacao = '" + usuario + "'";
				#endregion

				#region [ Monta o SQL ]
				strSql = "UPDATE t_FIN_BOLETO_ITEM" +
						 " SET " + strClausulaSet +
						 " WHERE" +
							" (id = " + idBoletoItem.ToString() + ")";
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					cmCommand = BD.criaSqlCommand();
					cmCommand.CommandText = strSql;
					intRetorno = BD.executaNonQuery(ref cmCommand);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 28 (débito de tarifas/custas)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia34RetiradoCartorioManutencaoCarteira ]
		public static bool atualizaBoletoItemOcorrencia34RetiradoCartorioManutencaoCarteira(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								DateTime dataOcorrenciaBanco,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto devido a ocorrência 34 (retirado de cartório e manutenção carteira)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrencia34.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrencia34.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrencia34.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrencia34.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia34.Parameters["@st_boleto_ocorrencia_34"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmBoletoItemAtualizaOcorrencia34.Parameters["@dt_ocorrencia_banco_boleto_ocorrencia_34"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia34.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrencia34);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento da ocorrência 34 (retirado de cartório e manutenção carteira)!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrenciaValaComum ]
		public static bool atualizaBoletoItemOcorrenciaValaComum(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								String motivoCodigoOcorrencia19,
								DateTime dataOcorrenciaBanco,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto com tratamento de vala comum para a ocorrência " + identificacaoOcorrencia;
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrenciaValaComum.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrenciaValaComum.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrenciaValaComum.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrenciaValaComum.Parameters["@ult_motivo_ocorrencia_19"].Value = motivoCodigoOcorrencia19;
				cmBoletoItemAtualizaOcorrenciaValaComum.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrenciaValaComum.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrenciaValaComum);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do registro durante tratamento de vala comum para a ocorrência " + identificacaoOcorrencia + "!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaBoletoItemOcorrencia24CepIrregular ]
		public static bool atualizaBoletoItemOcorrencia24CepIrregular(
								String usuario,
								int idBoletoItem,
								String identificacaoOcorrencia,
								String motivosRejeicoes,
								String motivoCodigoOcorrencia19,
								DateTime dataOcorrenciaBanco,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o registro da parcela do boleto com os dados da última ocorrência (" + identificacaoOcorrencia + ")";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoItemAtualizaOcorrencia24.Parameters["@id"].Value = idBoletoItem;
				cmBoletoItemAtualizaOcorrencia24.Parameters["@ult_identificacao_ocorrencia"].Value = identificacaoOcorrencia;
				cmBoletoItemAtualizaOcorrencia24.Parameters["@ult_motivos_rejeicoes"].Value = motivosRejeicoes;
				cmBoletoItemAtualizaOcorrencia24.Parameters["@ult_motivo_ocorrencia_19"].Value = motivoCodigoOcorrencia19;
				cmBoletoItemAtualizaOcorrencia24.Parameters["@ult_data_ocorrencia_banco"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmBoletoItemAtualizaOcorrencia24.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoItemAtualizaOcorrencia24);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar o boleto com os dados da última ocorrência (" + identificacaoOcorrencia + ")!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ corrigeBoletoOcorrencia24CepIrregular ]
		public static bool corrigeBoletoOcorrencia24CepIrregular(
								String usuario,
								int idBoleto,
								String endereco,
								String bairro,
								String cep,
								String cidade,
								String uf,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Corrige o endereço do sacado devido à ocorrência 24 (CEP irregular) e reseta status para reenviar no arquivo de remessa";
			int intRetorno;
			String strSql;
			SqlCommand cmComando;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoCorrigeOcorrencia24CepIrregular.Parameters["@id"].Value = idBoleto;
				cmBoletoCorrigeOcorrencia24CepIrregular.Parameters["@endereco_sacado"].Value = endereco;
				cmBoletoCorrigeOcorrencia24CepIrregular.Parameters["@bairro_sacado"].Value = bairro;
				cmBoletoCorrigeOcorrencia24CepIrregular.Parameters["@cep_sacado"].Value = Global.digitos(cep);
				cmBoletoCorrigeOcorrencia24CepIrregular.Parameters["@cidade_sacado"].Value = cidade;
				cmBoletoCorrigeOcorrencia24CepIrregular.Parameters["@uf_sacado"].Value = uf;
				cmBoletoCorrigeOcorrencia24CepIrregular.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar os dados do endereço ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoCorrigeOcorrencia24CepIrregular);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno != 1)
				{
					strMsgErro = "Falha ao tentar atualizar o endereço no registro do boleto durante o tratamento da ocorrência 24 (CEP irregular)!!";
					return false;
				}
				#endregion

				#region [ Tenta resetar o status dos registros das parcelas ]
				cmComando = BD.criaSqlCommand();
				strSql = "UPDATE t_FIN_BOLETO_ITEM SET " +
							"status = " + Global.Cte.FIN.CodBoletoItemStatus.INICIAL.ToString() + ", " +
							"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
							"dt_hr_ult_atualizacao = getdate(), " +
							"usuario_ult_atualizacao = '" + usuario + "' " +
						"WHERE " +
							"(id_boleto = " + idBoleto.ToString() + ")" +
							" AND (status = " + Global.Cte.FIN.CodBoletoItemStatus.BOLETO_REJEITADO_CEP_IRREGULAR.ToString() + ")";
				cmComando.CommandText = strSql;

				try
				{
					intRetorno = BD.executaNonQuery(ref cmComando);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno <= 0)
				{
					strMsgErro = "Falha ao tentar resetar o status dos registros das parcelas para reenviar no arquivo de remessa durante o tratamento da ocorrência 24 (CEP irregular): nenhuma parcela estava em situação que permitisse o reset!!";
					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ marcaBoletoOcorrenciaComoJaTratada ]
		public static bool marcaBoletoOcorrenciaComoJaTratada(
								String usuario,
								int idBoletoOcorrencia,
								String comentarioOcorrenciaTratada,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Marca a ocorrência de boleto como já tratada";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoOcorrenciaMarcaComoJaTratada.Parameters["@id"].Value = idBoletoOcorrencia;
				cmBoletoOcorrenciaMarcaComoJaTratada.Parameters["@comentario_ocorrencia_tratada"].Value = Texto.leftStr(comentarioOcorrenciaTratada, Global.Cte.FIN.TamanhoCampo.COMENTARIO_OCORRENCIA_TRATADA);
				cmBoletoOcorrenciaMarcaComoJaTratada.Parameters["@usuario_ocorrencia_tratada"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoOcorrenciaMarcaComoJaTratada);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno == 1)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar marcar a ocorrência de boleto como já tratada!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ marcaBoletoOcorrenciasComoJaTratadasByIdBoleto ]
		public static bool marcaBoletoOcorrenciasComoJaTratadasByIdBoleto(
								String usuario,
								int idBoleto,
								String comentarioOcorrenciaTratada,
								ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Marca as ocorrências de uma série de boletos como já tratadas";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmBoletoOcorrenciasMarcaComoJaTratadasByIdBoleto.Parameters["@id_boleto"].Value = idBoleto;
				cmBoletoOcorrenciasMarcaComoJaTratadasByIdBoleto.Parameters["@comentario_ocorrencia_tratada"].Value = Texto.leftStr(comentarioOcorrenciaTratada, Global.Cte.FIN.TamanhoCampo.COMENTARIO_OCORRENCIA_TRATADA);
				cmBoletoOcorrenciasMarcaComoJaTratadasByIdBoleto.Parameters["@usuario_ocorrencia_tratada"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmBoletoOcorrenciasMarcaComoJaTratadasByIdBoleto);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				if (intRetorno > 0)
				{
					blnSucesso = true;
				}
				else
				{
					blnSucesso = false;
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar marcar as ocorrências de uma série de boletos como já tratadas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + " - Falha!!" + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#endregion
	}
}
