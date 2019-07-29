using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Threading;

namespace FinanceiroService
{
	class PedidoDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmUpdatePedidoStatusAnaliseCredito;
		private static SqlCommand cmUpdatePedidoStatusPagto;
		private static SqlCommand cmUpdatePedidoVlPagoFamilia;
		private static SqlCommand cmUpdatePedidoCancela;
		private static SqlCommand cmUpdatePedidoStEntrega;
		private static SqlCommand cmInsertPedidoBlocoNotas;
		private static SqlCommand cmInsertComDataHoraPedidoBlocoNotas;
		private static SqlCommand cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento;
		private static SqlCommand cmInsertPedidoPagamento;
		private static SqlCommand cmInsertPedidoPagamentoBoletoEC;
		private static SqlCommand cmInsertFinPedidoHistPagto;
		private static SqlCommand cmInsertFinPedidoHistPagtoBoletoEC;
		private static SqlCommand cmUpdateFinPedidoHistPagtoCampoStatus;
		private static SqlCommand cmUpdateStPedidoNovoAnaliseCreditoMsgAlerta;
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
		static PedidoDAO()
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

			#region [ cmUpdatePedidoStatusAnaliseCredito ]
			strSql = "UPDATE t_PEDIDO SET " +
						"analise_credito = @analise_credito, " +
						"analise_credito_data = getdate(), " +
						"analise_credito_usuario = @analise_credito_usuario " +
					"WHERE (pedido = @pedido)";
			cmUpdatePedidoStatusAnaliseCredito = BD.criaSqlCommand();
			cmUpdatePedidoStatusAnaliseCredito.CommandText = strSql;
			cmUpdatePedidoStatusAnaliseCredito.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmUpdatePedidoStatusAnaliseCredito.Parameters.Add("@analise_credito", SqlDbType.SmallInt);
			cmUpdatePedidoStatusAnaliseCredito.Parameters.Add("@analise_credito_usuario", SqlDbType.VarChar, 10);
			cmUpdatePedidoStatusAnaliseCredito.Prepare();
			#endregion

			#region [ cmUpdatePedidoStatusPagto ]
			strSql = "UPDATE t_PEDIDO SET " +
						"st_pagto = @st_pagto, " +
						"dt_st_pagto = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"dt_hr_st_pagto = getdate(), " +
						"usuario_st_pagto = '" + Global.Cte.LogBd.Usuario.ID_USUARIO_LOG + "' " +
					"WHERE (pedido = @pedido)";
			cmUpdatePedidoStatusPagto = BD.criaSqlCommand();
			cmUpdatePedidoStatusPagto.CommandText = strSql;
			cmUpdatePedidoStatusPagto.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmUpdatePedidoStatusPagto.Parameters.Add("@st_pagto", SqlDbType.VarChar, 1);
			cmUpdatePedidoStatusPagto.Prepare();
			#endregion

			#region [ cmUpdatePedidoVlPagoFamilia ]
			strSql = "UPDATE t_PEDIDO SET " +
						"vl_pago_familia = @vl_pago_familia " +
					"WHERE (pedido = @pedido)";
			cmUpdatePedidoVlPagoFamilia = BD.criaSqlCommand();
			cmUpdatePedidoVlPagoFamilia.CommandText = strSql;
			cmUpdatePedidoVlPagoFamilia.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmUpdatePedidoVlPagoFamilia.Parameters.Add("@vl_pago_familia", SqlDbType.Money);
			cmUpdatePedidoVlPagoFamilia.Prepare();
			#endregion

			#region [ cmUpdatePedidoCancela ]
			strSql = "UPDATE t_PEDIDO SET " +
						"st_entrega = @st_entrega, " +
						"cancelado_data = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"cancelado_data_hora = getdate(), " +
						"cancelado_usuario = @cancelado_usuario, " +
						"cancelado_auto_status = 1, " +
						"cancelado_auto_data = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"cancelado_auto_data_hora = getdate(), " +
						"cancelado_auto_motivo = @cancelado_auto_motivo, " +
						"cancelado_codigo_motivo = @cancelado_codigo_motivo, " +
						"cancelado_codigo_sub_motivo = @cancelado_codigo_sub_motivo, " +
						"cancelado_motivo = @cancelado_motivo " +
					"WHERE (pedido = @pedido)";
			cmUpdatePedidoCancela = BD.criaSqlCommand();
			cmUpdatePedidoCancela.CommandText = strSql;
			cmUpdatePedidoCancela.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmUpdatePedidoCancela.Parameters.Add("@st_entrega", SqlDbType.VarChar, 3);
			cmUpdatePedidoCancela.Parameters.Add("@cancelado_usuario", SqlDbType.VarChar, 10);
			cmUpdatePedidoCancela.Parameters.Add("@cancelado_auto_motivo", SqlDbType.VarChar, 160);
			cmUpdatePedidoCancela.Parameters.Add("@cancelado_codigo_motivo", SqlDbType.VarChar, 3);
			cmUpdatePedidoCancela.Parameters.Add("@cancelado_codigo_sub_motivo", SqlDbType.VarChar, 3);
			cmUpdatePedidoCancela.Parameters.Add("@cancelado_motivo", SqlDbType.VarChar, 800);
			cmUpdatePedidoCancela.Prepare();
			#endregion

			#region [ cmUpdatePedidoStEntrega ]
			strSql = "UPDATE t_PEDIDO SET" +
						" st_entrega = @st_entrega" +
					" WHERE" +
						" (pedido = @pedido)";
			cmUpdatePedidoStEntrega = BD.criaSqlCommand();
			cmUpdatePedidoStEntrega.CommandText = strSql;
			cmUpdatePedidoStEntrega.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmUpdatePedidoStEntrega.Parameters.Add("@st_entrega", SqlDbType.VarChar, 3);
			cmUpdatePedidoStEntrega.Prepare();
			#endregion

			#region [ cmInsertPedidoBlocoNotas ]
			strSql = "INSERT INTO t_PEDIDO_BLOCO_NOTAS (" +
						"id, " +
						"pedido, " +
						"usuario, " +
						"nivel_acesso, " +
						"mensagem" +
					") VALUES (" +
						"@id, " +
						"@pedido, " +
						"@usuario, " +
						"@nivel_acesso, " +
						"@mensagem" +
					")";
			cmInsertPedidoBlocoNotas = BD.criaSqlCommand();
			cmInsertPedidoBlocoNotas.CommandText = strSql;
			cmInsertPedidoBlocoNotas.Parameters.Add("@id", SqlDbType.Int);
			cmInsertPedidoBlocoNotas.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmInsertPedidoBlocoNotas.Parameters.Add("@usuario", SqlDbType.VarChar, 10);
			cmInsertPedidoBlocoNotas.Parameters.Add("@nivel_acesso", SqlDbType.SmallInt);
			cmInsertPedidoBlocoNotas.Parameters.Add("@mensagem", SqlDbType.VarChar, 400);
			cmInsertPedidoBlocoNotas.Prepare();
			#endregion

			#region [ cmInsertComDataHoraPedidoBlocoNotas ]
			strSql = "INSERT INTO t_PEDIDO_BLOCO_NOTAS (" +
						"id, " +
						"dt_cadastro, " +
						"dt_hr_cadastro, " +
						"pedido, " +
						"usuario, " +
						"nivel_acesso, " +
						"mensagem" +
					") VALUES (" +
						"@id, " +
						"@dt_cadastro, " +
						"@dt_hr_cadastro, " +
						"@pedido, " +
						"@usuario, " +
						"@nivel_acesso, " +
						"@mensagem" +
					")";
			cmInsertComDataHoraPedidoBlocoNotas = BD.criaSqlCommand();
			cmInsertComDataHoraPedidoBlocoNotas.CommandText = strSql;
			cmInsertComDataHoraPedidoBlocoNotas.Parameters.Add("@id", SqlDbType.Int);
			cmInsertComDataHoraPedidoBlocoNotas.Parameters.Add("@dt_cadastro", SqlDbType.VarChar, 19);
			cmInsertComDataHoraPedidoBlocoNotas.Parameters.Add("@dt_hr_cadastro", SqlDbType.VarChar, 19);
			cmInsertComDataHoraPedidoBlocoNotas.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmInsertComDataHoraPedidoBlocoNotas.Parameters.Add("@usuario", SqlDbType.VarChar, 10);
			cmInsertComDataHoraPedidoBlocoNotas.Parameters.Add("@nivel_acesso", SqlDbType.SmallInt);
			cmInsertComDataHoraPedidoBlocoNotas.Parameters.Add("@mensagem", SqlDbType.VarChar, 400);
			cmInsertComDataHoraPedidoBlocoNotas.Prepare();
			#endregion

			#region [ cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento ]
			strSql = "UPDATE t_PEDIDO SET" +
						" vl_total_familia = @vl_total_familia," +
						" vl_total_NF = @vl_total_NF," +
						" vl_total_RA = @vl_total_RA," +
						" vl_total_RA_liquido = @vl_total_RA_liquido," +
						" qtde_parcelas_desagio_RA = @qtde_parcelas_desagio_RA," +
						" st_tem_desagio_RA = @st_tem_desagio_RA" +
					" WHERE" +
						" (pedido = @pedido)" +
						" AND (pedido = pedido_base)";
			cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento = BD.criaSqlCommand();
			cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.CommandText = strSql;
			cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.Parameters.Add("@vl_total_familia", SqlDbType.Money);
			cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.Parameters.Add("@vl_total_NF", SqlDbType.Money);
			cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.Parameters.Add("@vl_total_RA", SqlDbType.Money);
			cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.Parameters.Add("@vl_total_RA_liquido", SqlDbType.Money);
			cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.Parameters.Add("@qtde_parcelas_desagio_RA", SqlDbType.SmallInt);
			cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.Parameters.Add("@st_tem_desagio_RA", SqlDbType.SmallInt);
			cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.Prepare();
			#endregion

			#region [ cmInsertPedidoPagamento ]
			strSql = "INSERT INTO t_PEDIDO_PAGAMENTO (" +
						"id, " +
						"pedido, " +
						"data, " +
						"hora, " +
						"valor, " +
						"tipo_pagto, " +
						"usuario, " +
						"id_pagto_gw_pag_payment" +
					") VALUES (" +
						"@id, " +
						"@pedido, " +
						"@data, " +
						"@hora, " +
						"@valor, " +
						"@tipo_pagto, " +
						"@usuario, " +
						"@id_pagto_gw_pag_payment" +
					")";
			cmInsertPedidoPagamento = BD.criaSqlCommand();
			cmInsertPedidoPagamento.CommandText = strSql;
			cmInsertPedidoPagamento.Parameters.Add("@id", SqlDbType.VarChar, 12);
			cmInsertPedidoPagamento.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmInsertPedidoPagamento.Parameters.Add("@data", SqlDbType.VarChar, 19);
			cmInsertPedidoPagamento.Parameters.Add("@hora", SqlDbType.VarChar, 6);
			cmInsertPedidoPagamento.Parameters.Add("@valor", SqlDbType.Money);
			cmInsertPedidoPagamento.Parameters.Add("@tipo_pagto", SqlDbType.VarChar, 1);
			cmInsertPedidoPagamento.Parameters.Add("@usuario", SqlDbType.VarChar, 10);
			cmInsertPedidoPagamento.Parameters.Add("@id_pagto_gw_pag_payment", SqlDbType.Int);
			cmInsertPedidoPagamento.Prepare();
			#endregion

			#region [ cmInsertPedidoPagamentoBoletoEC ]
			strSql = "INSERT INTO t_PEDIDO_PAGAMENTO (" +
						"id, " +
						"pedido, " +
						"data, " +
						"hora, " +
						"valor, " +
						"tipo_pagto, " +
						"usuario, " +
						"id_braspag_webhook_complementar" +
					") VALUES (" +
						"@id, " +
						"@pedido, " +
						"@data, " +
						"@hora, " +
						"@valor, " +
						"@tipo_pagto, " +
						"@usuario, " +
						"@id_braspag_webhook_complementar" +
					")";
			cmInsertPedidoPagamentoBoletoEC = BD.criaSqlCommand();
			cmInsertPedidoPagamentoBoletoEC.CommandText = strSql;
			cmInsertPedidoPagamentoBoletoEC.Parameters.Add("@id", SqlDbType.VarChar, 12);
			cmInsertPedidoPagamentoBoletoEC.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmInsertPedidoPagamentoBoletoEC.Parameters.Add("@data", SqlDbType.VarChar, 19);
			cmInsertPedidoPagamentoBoletoEC.Parameters.Add("@hora", SqlDbType.VarChar, 6);
			cmInsertPedidoPagamentoBoletoEC.Parameters.Add("@valor", SqlDbType.Money);
			cmInsertPedidoPagamentoBoletoEC.Parameters.Add("@tipo_pagto", SqlDbType.VarChar, 1);
			cmInsertPedidoPagamentoBoletoEC.Parameters.Add("@usuario", SqlDbType.VarChar, 10);
			cmInsertPedidoPagamentoBoletoEC.Parameters.Add("@id_braspag_webhook_complementar", SqlDbType.Int);
			cmInsertPedidoPagamentoBoletoEC.Prepare();
			#endregion

			#region [ cmInsertFinPedidoHistPagto ]
			strSql = "INSERT INTO t_FIN_PEDIDO_HIST_PAGTO (" +
						"id, " +
						"pedido, " +
						"status, " +
						"ctrl_pagto_id_parcela, " +
						"ctrl_pagto_modulo, " +
						"dt_operacao, " +
						"valor_total, " +
						"valor_rateado, " +
						"descricao, " +
						"usuario_cadastro, " +
						"usuario_ult_atualizacao" +
					") VALUES (" +
						"@id, " +
						"@pedido, " +
						"@status, " +
						"@ctrl_pagto_id_parcela, " +
						"@ctrl_pagto_modulo, " +
						Global.sqlMontaGetdateSomenteData() + ", " +
						"@valor_total, " +
						"@valor_rateado, " +
						"@descricao, " +
						"@usuario_cadastro, " +
						"@usuario_ult_atualizacao" +
					")";
			cmInsertFinPedidoHistPagto = BD.criaSqlCommand();
			cmInsertFinPedidoHistPagto.CommandText = strSql;
			cmInsertFinPedidoHistPagto.Parameters.Add("@id", SqlDbType.Int);
			cmInsertFinPedidoHistPagto.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmInsertFinPedidoHistPagto.Parameters.Add("@status", SqlDbType.TinyInt);
			cmInsertFinPedidoHistPagto.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmInsertFinPedidoHistPagto.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmInsertFinPedidoHistPagto.Parameters.Add("@valor_total", SqlDbType.Money);
			cmInsertFinPedidoHistPagto.Parameters.Add("@valor_rateado", SqlDbType.Money);
			cmInsertFinPedidoHistPagto.Parameters.Add("@descricao", SqlDbType.VarChar, 160);
			cmInsertFinPedidoHistPagto.Parameters.Add("@usuario_cadastro", SqlDbType.VarChar, 10);
			cmInsertFinPedidoHistPagto.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmInsertFinPedidoHistPagto.Prepare();
			#endregion

			#region [ cmInsertFinPedidoHistPagtoBoletoEC ]
			strSql = "INSERT INTO t_FIN_PEDIDO_HIST_PAGTO (" +
						"id, " +
						"pedido, " +
						"status, " +
						"ctrl_pagto_id_parcela, " +
						"ctrl_pagto_modulo, " +
						"dt_vencto, " +
						"dt_credito, " +
						"valor_total, " +
						"valor_rateado, " +
						"valor_pago, " +
						"descricao, " +
						"usuario_cadastro, " +
						"usuario_ult_atualizacao" +
					") VALUES (" +
						"@id, " +
						"@pedido, " +
						"@status, " +
						"@ctrl_pagto_id_parcela, " +
						"@ctrl_pagto_modulo, " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_vencto") + ", " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_credito") + ", " +
						"@valor_total, " +
						"@valor_rateado, " +
						"@valor_pago, " +
						"@descricao, " +
						"@usuario_cadastro, " +
						"@usuario_ult_atualizacao" +
					")";
			cmInsertFinPedidoHistPagtoBoletoEC = BD.criaSqlCommand();
			cmInsertFinPedidoHistPagtoBoletoEC.CommandText = strSql;
			cmInsertFinPedidoHistPagtoBoletoEC.Parameters.Add("@id", SqlDbType.Int);
			cmInsertFinPedidoHistPagtoBoletoEC.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmInsertFinPedidoHistPagtoBoletoEC.Parameters.Add("@status", SqlDbType.TinyInt);
			cmInsertFinPedidoHistPagtoBoletoEC.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmInsertFinPedidoHistPagtoBoletoEC.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmInsertFinPedidoHistPagtoBoletoEC.Parameters.Add("@dt_vencto", SqlDbType.VarChar, 19);
			cmInsertFinPedidoHistPagtoBoletoEC.Parameters.Add("@dt_credito", SqlDbType.VarChar, 19);
			cmInsertFinPedidoHistPagtoBoletoEC.Parameters.Add("@valor_total", SqlDbType.Money);
			cmInsertFinPedidoHistPagtoBoletoEC.Parameters.Add("@valor_rateado", SqlDbType.Money);
			cmInsertFinPedidoHistPagtoBoletoEC.Parameters.Add("@valor_pago", SqlDbType.Money);
			cmInsertFinPedidoHistPagtoBoletoEC.Parameters.Add("@descricao", SqlDbType.VarChar, 160);
			cmInsertFinPedidoHistPagtoBoletoEC.Parameters.Add("@usuario_cadastro", SqlDbType.VarChar, 10);
			cmInsertFinPedidoHistPagtoBoletoEC.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmInsertFinPedidoHistPagtoBoletoEC.Prepare();
			#endregion

			#region [ cmUpdateFinPedidoHistPagtoCampoStatus ]
			strSql = "UPDATE t_FIN_PEDIDO_HIST_PAGTO SET" +
						" status = @status," +
						" dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + "," +
						" usuario_ult_atualizacao = @usuario_ult_atualizacao" +
					" WHERE" +
						" (id = @id)";
			cmUpdateFinPedidoHistPagtoCampoStatus = BD.criaSqlCommand();
			cmUpdateFinPedidoHistPagtoCampoStatus.CommandText = strSql;
			cmUpdateFinPedidoHistPagtoCampoStatus.Parameters.Add("@id", SqlDbType.Int);
			cmUpdateFinPedidoHistPagtoCampoStatus.Parameters.Add("@status", SqlDbType.TinyInt);
			cmUpdateFinPedidoHistPagtoCampoStatus.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmUpdateFinPedidoHistPagtoCampoStatus.Prepare();
			#endregion

			#region [ cmUpdateStPedidoNovoAnaliseCreditoMsgAlerta ]
			strSql = "UPDATE t_PEDIDO SET" +
						" st_pedido_novo_analise_credito_msg_alerta = @st_pedido_novo_analise_credito_msg_alerta," +
						" dt_hr_pedido_novo_analise_credito_msg_alerta = getdate()" +
					" WHERE" +
						" (pedido = @pedido)";
			cmUpdateStPedidoNovoAnaliseCreditoMsgAlerta = BD.criaSqlCommand();
			cmUpdateStPedidoNovoAnaliseCreditoMsgAlerta.CommandText = strSql;
			cmUpdateStPedidoNovoAnaliseCreditoMsgAlerta.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmUpdateStPedidoNovoAnaliseCreditoMsgAlerta.Parameters.Add("@st_pedido_novo_analise_credito_msg_alerta", SqlDbType.TinyInt);
			cmUpdateStPedidoNovoAnaliseCreditoMsgAlerta.Prepare();
			#endregion
		}
		#endregion

		#region [ getPedido ]
		/// <summary>
		/// Retorna um objeto Pedido contendo os dados lidos do BD
		/// </summary>
		/// <param name="numeroPedido">
		/// Número do pedido
		/// </param>
		/// <returns>
		/// Retorna um objeto Pedido contendo os dados lidos do BD
		/// </returns>
		public static Pedido getPedido(String numeroPedido)
		{
			#region [ Declarações ]
			String strSql;
			String numeroPedidoBase;
			decimal razaoValorPedidoFilhote = 0m;
			decimal vlBoletoDestePedido = 0m;
			decimal vlFormaPagtoDestePedido = 0m;
			decimal vlDiferencaArredondamento;
			decimal vlPagtoEmCartao = 0m;
			bool blnCartaoPagtoIntegral = false;
			Pedido pedido = new Pedido();
			PedidoItem pedidoItem;
			PedidoItemDevolvido pedidoItemDevolvido;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (numeroPedido == null) throw new Exception("Nº do pedido a ser consultado não foi fornecido!!");
			if (numeroPedido.Length == 0) throw new Exception("Nº do pedido a ser consultado não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Inicialização ]
			numeroPedido = numeroPedido.Trim();
			numeroPedidoBase = Global.retornaNumeroPedidoBase(numeroPedido);
			#endregion

			#region [ Pesquisa pedido ]

			#region [ Monta Select ]
			strSql = "SELECT" +
						" t_PEDIDO.*," +
						" t_LOJA.razao_social AS loja_razao_social," +
						" t_LOJA.nome AS loja_nome," +
						" t_USUARIO_VENDEDOR.nome AS vendedor_nome," +
						" t_ORCAMENTISTA_E_INDICADOR.desempenho_nota AS indicador_desempenho_nota" +
					" FROM t_PEDIDO" +
						" INNER JOIN t_LOJA ON (t_PEDIDO.loja=t_LOJA.loja)" +
						" INNER JOIN t_USUARIO AS t_USUARIO_VENDEDOR ON (t_PEDIDO.vendedor=t_USUARIO_VENDEDOR.usuario)" +
						" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" +
					" WHERE" +
						" (pedido = '" + numeroPedido + "')";
			#endregion

			#region [ Executa a consulta ]
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);
			#endregion

			if (dtbResultado.Rows.Count == 0) throw new Exception("Pedido nº " + numeroPedido + " não foi encontrado!!");

			#region [ Carrega os dados ]
			rowResultado = dtbResultado.Rows[0];
			pedido.pedido = BD.readToString(rowResultado["pedido"]);
			pedido.loja = BD.readToString(rowResultado["loja"]);
			pedido.loja_razao_social = BD.readToString(rowResultado["loja_razao_social"]);
			pedido.loja_nome = BD.readToString(rowResultado["loja_nome"]);
			pedido.data = BD.readToDateTime(rowResultado["data"]);
			pedido.hora = BD.readToString(rowResultado["hora"]);
			pedido.data_hora = BD.readToDateTime(rowResultado["data_hora"]);
			pedido.id_cliente = BD.readToString(rowResultado["id_cliente"]);
			pedido.midia = BD.readToString(rowResultado["midia"]);
			pedido.servicos = BD.readToString(rowResultado["servicos"]);
			pedido.vl_servicos = BD.readToDecimal(rowResultado["vl_servicos"]);
			pedido.vendedor = BD.readToString(rowResultado["vendedor"]);
			pedido.vendedor_nome = BD.readToString(rowResultado["vendedor_nome"]);
			pedido.st_entrega = BD.readToString(rowResultado["st_entrega"]);
			pedido.entregue_data = BD.readToDateTime(rowResultado["entregue_data"]);
			pedido.entregue_usuario = BD.readToString(rowResultado["entregue_usuario"]);
			pedido.cancelado_data = BD.readToDateTime(rowResultado["cancelado_data"]);
			pedido.cancelado_usuario = BD.readToString(rowResultado["cancelado_usuario"]);
			pedido.st_pagto = BD.readToString(rowResultado["st_pagto"]);
			pedido.st_recebido = BD.readToString(rowResultado["st_recebido"]);
			pedido.obs_1 = BD.readToString(rowResultado["obs_1"]);
			pedido.obs_2 = BD.readToString(rowResultado["obs_2"]);
			pedido.qtde_parcelas = BD.readToShort(rowResultado["qtde_parcelas"]);
			pedido.forma_pagto = BD.readToString(rowResultado["forma_pagto"]);
			pedido.vl_total_familia = BD.readToDecimal(rowResultado["vl_total_familia"]);
			pedido.vl_pago_familia = BD.readToDecimal(rowResultado["vl_pago_familia"]);
			pedido.split_status = BD.readToShort(rowResultado["split_status"]);
			pedido.split_data = BD.readToDateTime(rowResultado["split_data"]);
			pedido.split_hora = BD.readToString(rowResultado["split_hora"]);
			pedido.split_usuario = BD.readToString(rowResultado["split_usuario"]);
			pedido.a_entregar_status = BD.readToShort(rowResultado["a_entregar_status"]);
			pedido.a_entregar_data_marcada = BD.readToDateTime(rowResultado["a_entregar_data_marcada"]);
			pedido.a_entregar_data = BD.readToDateTime(rowResultado["a_entregar_data"]);
			pedido.a_entregar_hora = BD.readToString(rowResultado["a_entregar_hora"]);
			pedido.a_entregar_usuario = BD.readToString(rowResultado["a_entregar_usuario"]);
			pedido.loja_indicou = BD.readToString(rowResultado["loja_indicou"]);
			pedido.comissao_loja_indicou = BD.readToSingle(rowResultado["comissao_loja_indicou"]);
			pedido.venda_externa = BD.readToShort(rowResultado["venda_externa"]);
			pedido.vl_frete = BD.readToDecimal(rowResultado["vl_frete"]);
			pedido.transportadora_id = BD.readToString(rowResultado["transportadora_id"]);
			pedido.transportadora_data = BD.readToDateTime(rowResultado["transportadora_data"]);
			pedido.transportadora_usuario = BD.readToString(rowResultado["transportadora_usuario"]);
			pedido.analise_credito = BD.readToShort(rowResultado["analise_credito"]);
			pedido.analise_credito_data = BD.readToDateTime(rowResultado["analise_credito_data"]);
			pedido.analise_credito_usuario = BD.readToString(rowResultado["analise_credito_usuario"]);
			pedido.tipo_parcelamento = BD.readToShort(rowResultado["tipo_parcelamento"]);
			pedido.av_forma_pagto = BD.readToShort(rowResultado["av_forma_pagto"]);
			pedido.pc_qtde_parcelas = BD.readToShort(rowResultado["pc_qtde_parcelas"]);
			pedido.pc_valor_parcela = BD.readToDecimal(rowResultado["pc_valor_parcela"]);
			pedido.pc_maquineta_qtde_parcelas = BD.readToShort(rowResultado["pc_maquineta_qtde_parcelas"]);
			pedido.pc_maquineta_valor_parcela = BD.readToDecimal(rowResultado["pc_maquineta_valor_parcela"]);
			pedido.pce_forma_pagto_entrada = BD.readToShort(rowResultado["pce_forma_pagto_entrada"]);
			pedido.pce_forma_pagto_prestacao = BD.readToShort(rowResultado["pce_forma_pagto_prestacao"]);
			pedido.pce_entrada_valor = BD.readToDecimal(rowResultado["pce_entrada_valor"]);
			pedido.pce_prestacao_qtde = BD.readToShort(rowResultado["pce_prestacao_qtde"]);
			pedido.pce_prestacao_valor = BD.readToDecimal(rowResultado["pce_prestacao_valor"]);
			pedido.pce_prestacao_periodo = BD.readToShort(rowResultado["pce_prestacao_periodo"]);
			pedido.pse_forma_pagto_prim_prest = BD.readToShort(rowResultado["pse_forma_pagto_prim_prest"]);
			pedido.pse_forma_pagto_demais_prest = BD.readToShort(rowResultado["pse_forma_pagto_demais_prest"]);
			pedido.pse_prim_prest_valor = BD.readToDecimal(rowResultado["pse_prim_prest_valor"]);
			pedido.pse_prim_prest_apos = BD.readToShort(rowResultado["pse_prim_prest_apos"]);
			pedido.pse_demais_prest_qtde = BD.readToShort(rowResultado["pse_demais_prest_qtde"]);
			pedido.pse_demais_prest_valor = BD.readToDecimal(rowResultado["pse_demais_prest_valor"]);
			pedido.pse_demais_prest_periodo = BD.readToShort(rowResultado["pse_demais_prest_periodo"]);
			pedido.pu_forma_pagto = BD.readToShort(rowResultado["pu_forma_pagto"]);
			pedido.pu_valor = BD.readToDecimal(rowResultado["pu_valor"]);
			pedido.pu_vencto_apos = BD.readToShort(rowResultado["pu_vencto_apos"]);
			pedido.indicador = BD.readToString(rowResultado["indicador"]);
			pedido.indicador_desempenho_nota = BD.readToString(rowResultado["indicador_desempenho_nota"]);
			pedido.vl_total_NF = BD.readToDecimal(rowResultado["vl_total_NF"]);
			pedido.vl_total_RA = BD.readToDecimal(rowResultado["vl_total_RA"]);
			pedido.perc_RT = BD.readToSingle(rowResultado["perc_RT"]);
			pedido.st_orc_virou_pedido = BD.readToShort(rowResultado["st_orc_virou_pedido"]);
			pedido.orcamento = BD.readToString(rowResultado["orcamento"]);
			pedido.orcamentista = BD.readToString(rowResultado["orcamentista"]);
			pedido.comissao_paga = BD.readToShort(rowResultado["comissao_paga"]);
			pedido.comissao_paga_ult_op = BD.readToString(rowResultado["comissao_paga_ult_op"]);
			pedido.comissao_paga_data = BD.readToDateTime(rowResultado["comissao_paga_data"]);
			pedido.comissao_paga_usuario = BD.readToString(rowResultado["comissao_paga_usuario"]);
			pedido.perc_desagio_RA = BD.readToSingle(rowResultado["perc_desagio_RA"]);
			pedido.perc_limite_RA_sem_desagio = BD.readToSingle(rowResultado["perc_limite_RA_sem_desagio"]);
			pedido.vl_total_RA_liquido = BD.readToDecimal(rowResultado["vl_total_RA_liquido"]);
			pedido.st_tem_desagio_RA = BD.readToShort(rowResultado["st_tem_desagio_RA"]);
			pedido.qtde_parcelas_desagio_RA = BD.readToShort(rowResultado["qtde_parcelas_desagio_RA"]);
			pedido.transportadora_num_coleta = BD.readToString(rowResultado["transportadora_num_coleta"]);
			pedido.transportadora_contato = BD.readToString(rowResultado["transportadora_contato"]);
			pedido.st_end_entrega = BD.readToShort(rowResultado["st_end_entrega"]);
			pedido.endEtg_endereco = BD.readToString(rowResultado["EndEtg_endereco"]);
			pedido.endEtg_endereco_numero = BD.readToString(rowResultado["EndEtg_endereco_numero"]);
			pedido.endEtg_endereco_complemento = BD.readToString(rowResultado["EndEtg_endereco_complemento"]);
			pedido.endEtg_bairro = BD.readToString(rowResultado["EndEtg_bairro"]);
			pedido.endEtg_cidade = BD.readToString(rowResultado["EndEtg_cidade"]);
			pedido.endEtg_uf = BD.readToString(rowResultado["EndEtg_uf"]);
			pedido.endEtg_cep = BD.readToString(rowResultado["EndEtg_cep"]);
			pedido.st_etg_imediata = BD.readToShort(rowResultado["st_etg_imediata"]);
			pedido.etg_imediata_data = BD.readToDateTime(rowResultado["etg_imediata_data"]);
			pedido.etg_imediata_usuario = BD.readToString(rowResultado["etg_imediata_usuario"]);
			pedido.frete_status = BD.readToShort(rowResultado["frete_status"]);
			pedido.frete_valor = BD.readToDecimal(rowResultado["frete_valor"]);
			pedido.frete_data = BD.readToDateTime(rowResultado["frete_data"]);
			pedido.frete_usuario = BD.readToString(rowResultado["frete_usuario"]);
			pedido.stBemUsoConsumo = BD.readToShort(rowResultado["StBemUsoConsumo"]);
			pedido.pedidoRecebidoStatus = BD.readToShort(rowResultado["PedidoRecebidoStatus"]);
			pedido.pedidoRecebidoData = BD.readToDateTime(rowResultado["PedidoRecebidoData"]);
			pedido.pedidoRecebidoUsuarioUltAtualiz = BD.readToString(rowResultado["PedidoRecebidoUsuarioUltAtualiz"]);
			pedido.pedidoRecebidoDtHrUltAtualiz = BD.readToDateTime(rowResultado["PedidoRecebidoDtHrUltAtualiz"]);
			pedido.instaladorInstalaStatus = BD.readToShort(rowResultado["InstaladorInstalaStatus"]);
			pedido.instaladorInstalaUsuarioUltAtualiz = BD.readToString(rowResultado["InstaladorInstalaUsuarioUltAtualiz"]);
			pedido.instaladorInstalaDtHrUltAtualiz = BD.readToDateTime(rowResultado["InstaladorInstalaDtHrUltAtualiz"]);
			pedido.custoFinancFornecTipoParcelamento = BD.readToString(rowResultado["custoFinancFornecTipoParcelamento"]);
			pedido.custoFinancFornecQtdeParcelas = BD.readToShort(rowResultado["custoFinancFornecQtdeParcelas"]);
			pedido.tamanho_num_pedido = BD.readToInt(rowResultado["tamanho_num_pedido"]);
			pedido.pedido_base = BD.readToString(rowResultado["pedido_base"]);
			pedido.st_forma_pagto_somente_cartao = BD.readToByte(rowResultado["st_forma_pagto_somente_cartao"]);
			pedido.id_nfe_emitente = BD.readToInt(rowResultado["id_nfe_emitente"]);
			pedido.st_auto_split = BD.readToByte(rowResultado["st_auto_split"]);
			#endregion

			#endregion

			#region [ Pesquisa pedido-base? ]
			if (Global.isPedidoFilhote(numeroPedido))
			{
				#region [ Monta Select do pedido-base ]
				strSql = "SELECT" +
							" t_PEDIDO.*," +
							" t_ORCAMENTISTA_E_INDICADOR.desempenho_nota AS indicador_desempenho_nota" +
						" FROM t_PEDIDO" +
							" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" +
						" WHERE" +
							" (pedido = '" + numeroPedidoBase + "')";
				#endregion

				#region [ Executa a consulta ]
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daDataAdapter.Fill(dtbResultado);
				#endregion

				if (dtbResultado.Rows.Count == 0) throw new Exception("Pedido-base nº " + numeroPedidoBase + " não foi encontrado!!");

				#region [ Carrega os dados do pedido-base ]
				rowResultado = dtbResultado.Rows[0];
				pedido.st_pagto = BD.readToString(rowResultado["st_pagto"]);
				pedido.st_recebido = BD.readToString(rowResultado["st_recebido"]);
				pedido.vl_total_familia = BD.readToDecimal(rowResultado["vl_total_familia"]);
				pedido.vl_pago_familia = BD.readToDecimal(rowResultado["vl_pago_familia"]);
				pedido.analise_credito = BD.readToShort(rowResultado["analise_credito"]);
				pedido.analise_credito_data = BD.readToDateTime(rowResultado["analise_credito_data"]);
				pedido.analise_credito_usuario = BD.readToString(rowResultado["analise_credito_usuario"]);
				pedido.tipo_parcelamento = BD.readToShort(rowResultado["tipo_parcelamento"]);
				pedido.av_forma_pagto = BD.readToShort(rowResultado["av_forma_pagto"]);
				pedido.pc_qtde_parcelas = BD.readToShort(rowResultado["pc_qtde_parcelas"]);
				pedido.pc_valor_parcela = BD.readToDecimal(rowResultado["pc_valor_parcela"]);
				pedido.pc_maquineta_qtde_parcelas = BD.readToShort(rowResultado["pc_maquineta_qtde_parcelas"]);
				pedido.pc_maquineta_valor_parcela = BD.readToDecimal(rowResultado["pc_maquineta_valor_parcela"]);
				pedido.pce_forma_pagto_entrada = BD.readToShort(rowResultado["pce_forma_pagto_entrada"]);
				pedido.pce_forma_pagto_prestacao = BD.readToShort(rowResultado["pce_forma_pagto_prestacao"]);
				pedido.pce_entrada_valor = BD.readToDecimal(rowResultado["pce_entrada_valor"]);
				pedido.pce_prestacao_qtde = BD.readToShort(rowResultado["pce_prestacao_qtde"]);
				pedido.pce_prestacao_valor = BD.readToDecimal(rowResultado["pce_prestacao_valor"]);
				pedido.pce_prestacao_periodo = BD.readToShort(rowResultado["pce_prestacao_periodo"]);
				pedido.pse_forma_pagto_prim_prest = BD.readToShort(rowResultado["pse_forma_pagto_prim_prest"]);
				pedido.pse_forma_pagto_demais_prest = BD.readToShort(rowResultado["pse_forma_pagto_demais_prest"]);
				pedido.pse_prim_prest_valor = BD.readToDecimal(rowResultado["pse_prim_prest_valor"]);
				pedido.pse_prim_prest_apos = BD.readToShort(rowResultado["pse_prim_prest_apos"]);
				pedido.pse_demais_prest_qtde = BD.readToShort(rowResultado["pse_demais_prest_qtde"]);
				pedido.pse_demais_prest_valor = BD.readToDecimal(rowResultado["pse_demais_prest_valor"]);
				pedido.pse_demais_prest_periodo = BD.readToShort(rowResultado["pse_demais_prest_periodo"]);
				pedido.pu_forma_pagto = BD.readToShort(rowResultado["pu_forma_pagto"]);
				pedido.pu_valor = BD.readToDecimal(rowResultado["pu_valor"]);
				pedido.pu_vencto_apos = BD.readToShort(rowResultado["pu_vencto_apos"]);
				pedido.custoFinancFornecTipoParcelamento = BD.readToString(rowResultado["custoFinancFornecTipoParcelamento"]);
				pedido.custoFinancFornecQtdeParcelas = BD.readToShort(rowResultado["custoFinancFornecQtdeParcelas"]);
				pedido.indicador = BD.readToString(rowResultado["indicador"]);
				pedido.indicador_desempenho_nota = BD.readToString(rowResultado["indicador_desempenho_nota"]);
				pedido.vl_total_NF = BD.readToDecimal(rowResultado["vl_total_NF"]);
				pedido.vl_total_RA = BD.readToDecimal(rowResultado["vl_total_RA"]);
				pedido.perc_RT = BD.readToSingle(rowResultado["perc_RT"]);
				pedido.st_orc_virou_pedido = BD.readToShort(rowResultado["st_orc_virou_pedido"]);
				pedido.orcamento = BD.readToString(rowResultado["orcamento"]);
				pedido.orcamentista = BD.readToString(rowResultado["orcamentista"]);
				pedido.comissao_paga = BD.readToShort(rowResultado["comissao_paga"]);
				pedido.comissao_paga_ult_op = BD.readToString(rowResultado["comissao_paga_ult_op"]);
				pedido.comissao_paga_data = BD.readToDateTime(rowResultado["comissao_paga_data"]);
				pedido.comissao_paga_usuario = BD.readToString(rowResultado["comissao_paga_usuario"]);
				pedido.perc_desagio_RA = BD.readToSingle(rowResultado["perc_desagio_RA"]);
				pedido.perc_limite_RA_sem_desagio = BD.readToSingle(rowResultado["perc_limite_RA_sem_desagio"]);
				pedido.vl_total_RA_liquido = BD.readToDecimal(rowResultado["vl_total_RA_liquido"]);
				pedido.st_tem_desagio_RA = BD.readToShort(rowResultado["st_tem_desagio_RA"]);
				pedido.qtde_parcelas_desagio_RA = BD.readToShort(rowResultado["qtde_parcelas_desagio_RA"]);
				pedido.st_forma_pagto_somente_cartao = BD.readToByte(rowResultado["st_forma_pagto_somente_cartao"]);
				#endregion
			}
			#endregion

			#region [ Pesquisa itens do pedido ]

			#region [ Monta Select ]
			strSql = "SELECT " +
						"*" +
					" FROM t_PEDIDO_ITEM" +
					" WHERE" +
						" (pedido = '" + numeroPedido + "')" +
					" ORDER BY" +
						" sequencia";
			#endregion

			#region [ Executa a consulta ]
			cmCommand.CommandText = strSql;
			dtbResultado.Reset();
			daDataAdapter.Fill(dtbResultado);
			#endregion

			if (dtbResultado.Rows.Count == 0) throw new Exception("Itens do pedido nº " + numeroPedido + " não foram encontrados!!");

			#region [ Carrega os dados ]
			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				rowResultado = dtbResultado.Rows[i];
				pedidoItem = new PedidoItem();
				pedidoItem.pedido = BD.readToString(rowResultado["pedido"]);
				pedidoItem.fabricante = BD.readToString(rowResultado["fabricante"]);
				pedidoItem.produto = BD.readToString(rowResultado["produto"]);
				pedidoItem.qtde = BD.readToShort(rowResultado["qtde"]);
				pedidoItem.desc_dado = BD.readToSingle(rowResultado["desc_dado"]);
				pedidoItem.preco_venda = BD.readToDecimal(rowResultado["preco_venda"]);
				pedidoItem.preco_fabricante = BD.readToDecimal(rowResultado["preco_fabricante"]);
				pedidoItem.preco_lista = BD.readToDecimal(rowResultado["preco_lista"]);
				pedidoItem.margem = BD.readToSingle(rowResultado["margem"]);
				pedidoItem.desc_max = BD.readToSingle(rowResultado["desc_max"]);
				pedidoItem.comissao = BD.readToSingle(rowResultado["comissao"]);
				pedidoItem.descricao = BD.readToString(rowResultado["descricao"]);
				pedidoItem.ean = BD.readToString(rowResultado["ean"]);
				pedidoItem.grupo = BD.readToString(rowResultado["grupo"]);
				pedidoItem.peso = BD.readToSingle(rowResultado["peso"]);
				pedidoItem.qtde_volumes = BD.readToShort(rowResultado["qtde_volumes"]);
				pedidoItem.abaixo_min_status = BD.readToShort(rowResultado["abaixo_min_status"]);
				pedidoItem.abaixo_min_autorizacao = BD.readToString(rowResultado["abaixo_min_autorizacao"]);
				pedidoItem.abaixo_min_autorizador = BD.readToString(rowResultado["abaixo_min_autorizador"]);
				pedidoItem.sequencia = BD.readToShort(rowResultado["sequencia"]);
				pedidoItem.markup_fabricante = BD.readToSingle(rowResultado["markup_fabricante"]);
				pedidoItem.preco_NF = BD.readToDecimal(rowResultado["preco_NF"]);
				pedidoItem.abaixo_min_superv_autorizador = BD.readToString(rowResultado["abaixo_min_superv_autorizador"]);
				pedidoItem.vl_custo2 = BD.readToDecimal(rowResultado["vl_custo2"]);
				pedidoItem.descricao_html = BD.readToString(rowResultado["descricao_html"]);
				pedidoItem.custoFinancFornecCoeficiente = BD.readToSingle(rowResultado["custoFinancFornecCoeficiente"]);
				pedidoItem.custoFinancFornecPrecoListaBase = BD.readToDecimal(rowResultado["custoFinancFornecPrecoListaBase"]);
				pedido.listaPedidoItem.Add(pedidoItem);

				pedido.vlTotalPrecoNfDestePedido += pedidoItem.qtde * pedidoItem.preco_NF;
				pedido.vlTotalPrecoVendaDestePedido += pedidoItem.qtde * pedidoItem.preco_venda;
			}
			#endregion

			#endregion

			#region [ Pesquisa itens devolvidos ]

			#region [ Monta Select ]
			strSql = "SELECT " +
						"*" +
					" FROM t_PEDIDO_ITEM_DEVOLVIDO" +
					" WHERE" +
						" (pedido = '" + numeroPedido + "')" +
					" ORDER BY" +
						" devolucao_data," +
						" devolucao_hora";
			#endregion

			#region [ Executa a consulta ]
			cmCommand.CommandText = strSql;
			dtbResultado.Reset();
			daDataAdapter.Fill(dtbResultado);
			#endregion

			#region [ Carrega os dados ]
			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				rowResultado = dtbResultado.Rows[i];
				pedidoItemDevolvido = new PedidoItemDevolvido();
				pedidoItemDevolvido.id = BD.readToString(rowResultado["id"]);
				pedidoItemDevolvido.devolucao_data = BD.readToDateTime(rowResultado["devolucao_data"]);
				pedidoItemDevolvido.devolucao_hora = BD.readToString(rowResultado["devolucao_hora"]);
				pedidoItemDevolvido.devolucao_usuario = BD.readToString(rowResultado["devolucao_usuario"]);
				pedidoItemDevolvido.pedido = BD.readToString(rowResultado["pedido"]);
				pedidoItemDevolvido.fabricante = BD.readToString(rowResultado["fabricante"]);
				pedidoItemDevolvido.produto = BD.readToString(rowResultado["produto"]);
				pedidoItemDevolvido.qtde = BD.readToShort(rowResultado["qtde"]);
				pedidoItemDevolvido.desc_dado = BD.readToSingle(rowResultado["desc_dado"]);
				pedidoItemDevolvido.preco_venda = BD.readToDecimal(rowResultado["preco_venda"]);
				pedidoItemDevolvido.preco_fabricante = BD.readToDecimal(rowResultado["preco_fabricante"]);
				pedidoItemDevolvido.preco_lista = BD.readToDecimal(rowResultado["preco_lista"]);
				pedidoItemDevolvido.margem = BD.readToSingle(rowResultado["margem"]);
				pedidoItemDevolvido.desc_max = BD.readToSingle(rowResultado["desc_max"]);
				pedidoItemDevolvido.comissao = BD.readToSingle(rowResultado["comissao"]);
				pedidoItemDevolvido.descricao = BD.readToString(rowResultado["descricao"]);
				pedidoItemDevolvido.ean = BD.readToString(rowResultado["ean"]);
				pedidoItemDevolvido.grupo = BD.readToString(rowResultado["grupo"]);
				pedidoItemDevolvido.peso = BD.readToSingle(rowResultado["peso"]);
				pedidoItemDevolvido.qtde_volumes = BD.readToShort(rowResultado["qtde_volumes"]);
				pedidoItemDevolvido.abaixo_min_status = BD.readToShort(rowResultado["abaixo_min_status"]);
				pedidoItemDevolvido.abaixo_min_autorizacao = BD.readToString(rowResultado["abaixo_min_autorizacao"]);
				pedidoItemDevolvido.abaixo_min_autorizador = BD.readToString(rowResultado["abaixo_min_autorizador"]);
				pedidoItemDevolvido.markup_fabricante = BD.readToSingle(rowResultado["markup_fabricante"]);
				pedidoItemDevolvido.motivo = BD.readToString(rowResultado["motivo"]);
				pedidoItemDevolvido.preco_NF = BD.readToDecimal(rowResultado["preco_NF"]);
				pedidoItemDevolvido.comissao_descontada = BD.readToShort(rowResultado["comissao_descontada"]);
				pedidoItemDevolvido.comissao_descontada_ult_op = BD.readToString(rowResultado["comissao_descontada_ult_op"]);
				pedidoItemDevolvido.comissao_descontada_data = BD.readToDateTime(rowResultado["comissao_descontada_data"]);
				pedidoItemDevolvido.comissao_descontada_usuario = BD.readToString(rowResultado["comissao_descontada_usuario"]);
				pedidoItemDevolvido.abaixo_min_superv_autorizador = BD.readToString(rowResultado["abaixo_min_superv_autorizador"]);
				pedidoItemDevolvido.vl_custo2 = BD.readToDecimal(rowResultado["vl_custo2"]);
				pedidoItemDevolvido.descricao_html = BD.readToString(rowResultado["descricao_html"]);
				pedidoItemDevolvido.custoFinancFornecCoeficiente = BD.readToSingle(rowResultado["custoFinancFornecCoeficiente"]);
				pedidoItemDevolvido.custoFinancFornecPrecoListaBase = BD.readToDecimal(rowResultado["custoFinancFornecPrecoListaBase"]);
				pedido.listaPedidoItemDevolvido.Add(pedidoItemDevolvido);

				pedido.vlTotalPrecoNfDestePedido -= pedidoItemDevolvido.qtde * pedidoItemDevolvido.preco_NF;
				pedido.vlTotalPrecoVendaDestePedido -= pedidoItemDevolvido.qtde * pedidoItemDevolvido.preco_venda;
			}
			#endregion

			#endregion

			#region [ Calcula valor total já pago ]

			#region [ Monta Select ]
			strSql = "SELECT" +
						" Coalesce(SUM(valor), 0) AS vl_total" +
					" FROM t_PEDIDO_PAGAMENTO" +
					" WHERE" +
						" (pedido LIKE '" + numeroPedidoBase + BD.CARACTER_CURINGA_TODOS + "')";
			#endregion

			#region [ Executa a consulta ]
			cmCommand.CommandText = strSql;
			dtbResultado.Reset();
			daDataAdapter.Fill(dtbResultado);
			#endregion

			if (dtbResultado.Rows.Count == 0) throw new Exception("Falha ao calcular o valor total já pago!!");

			#region [ Carrega os dados ]
			rowResultado = dtbResultado.Rows[0];
			pedido.vlTotalFamiliaPago = BD.readToDecimal(rowResultado["vl_total"]);
			#endregion

			#endregion

			#region [ Calcula o valor total da família de pedidos ]

			#region [ Monta Select ]
			strSql = "SELECT" +
						" Coalesce(SUM(qtde*preco_venda), 0) AS vl_total," +
						" Coalesce(SUM(qtde*preco_NF), 0) AS vl_total_NF" +
					" FROM t_PEDIDO_ITEM INNER JOIN t_PEDIDO" +
						" ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" +
					" WHERE" +
						" (st_entrega<>'" + Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO + "')" +
						" AND (t_PEDIDO.pedido LIKE '" + numeroPedidoBase + BD.CARACTER_CURINGA_TODOS + "')";
			#endregion

			#region [ Executa a consulta ]
			cmCommand.CommandText = strSql;
			dtbResultado.Reset();
			daDataAdapter.Fill(dtbResultado);
			#endregion

			if (dtbResultado.Rows.Count == 0) throw new Exception("Falha ao calcular o valor total da família de pedidos!!");

			#region [ Carrega os dados ]
			rowResultado = dtbResultado.Rows[0];
			pedido.vlTotalFamiliaPrecoVenda = BD.readToDecimal(rowResultado["vl_total"]);
			pedido.vlTotalFamiliaPrecoNF = BD.readToDecimal(rowResultado["vl_total_NF"]);
			#endregion

			#endregion

			#region [ Calcula o valor total em devoluções da família de pedidos ]

			#region [ Monta Select ]
			strSql = "SELECT" +
						" Coalesce(SUM(qtde*preco_venda), 0) AS vl_total," +
						" Coalesce(SUM(qtde*preco_NF), 0) AS vl_total_NF" +
					" FROM t_PEDIDO_ITEM_DEVOLVIDO" +
					" WHERE" +
						" (pedido LIKE '" + numeroPedidoBase + BD.CARACTER_CURINGA_TODOS + "')";
			#endregion

			#region [ Executa a consulta ]
			cmCommand.CommandText = strSql;
			dtbResultado.Reset();
			daDataAdapter.Fill(dtbResultado);
			#endregion

			if (dtbResultado.Rows.Count == 0) throw new Exception("Falha ao calcular o valor total em devoluções da família de pedidos!!");

			#region [ Carrega os dados ]
			rowResultado = dtbResultado.Rows[0];
			pedido.vlTotalFamiliaDevolucaoPrecoVenda = BD.readToDecimal(rowResultado["vl_total"]);
			pedido.vlTotalFamiliaDevolucaoPrecoNF = BD.readToDecimal(rowResultado["vl_total_NF"]);

			pedido.vlTotalFamiliaPrecoVenda -= pedido.vlTotalFamiliaDevolucaoPrecoVenda;
			pedido.vlTotalFamiliaPrecoNF -= pedido.vlTotalFamiliaDevolucaoPrecoNF;
			#endregion

			#endregion

			#region [ Calcula o valor em boletos deste pedido ]
			if (pedido.vlTotalFamiliaPrecoNF == 0)
			{
				razaoValorPedidoFilhote = 0m;
			}
			else
			{
				razaoValorPedidoFilhote = pedido.vlTotalPrecoNfDestePedido / pedido.vlTotalFamiliaPrecoNF;
			}

			#region [ Calcula o valor proporcional, pois pode ser um pedido filhote ]
			if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_A_VISTA)
			{
				#region [ À vista ]
				vlFormaPagtoDestePedido = pedido.vlTotalPrecoNfDestePedido;
				if ((pedido.av_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
					|| (pedido.av_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO_AV))
				{
					vlBoletoDestePedido = pedido.vlTotalPrecoNfDestePedido;
				}
				#endregion
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELA_UNICA)
			{
				#region [ Parcela única ]
				vlFormaPagtoDestePedido = Global.arredondaParaMonetario(pedido.pu_valor * razaoValorPedidoFilhote);
				if (pedido.pu_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
				{
					vlBoletoDestePedido = Global.arredondaParaMonetario(pedido.pu_valor * razaoValorPedidoFilhote);
				}
				#endregion
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA)
			{
				#region [ Parcelado com entrada ]
				vlFormaPagtoDestePedido = Global.arredondaParaMonetario(pedido.pce_entrada_valor * razaoValorPedidoFilhote);
				if (pedido.pce_forma_pagto_entrada == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
				{
					vlBoletoDestePedido = Global.arredondaParaMonetario(pedido.pce_entrada_valor * razaoValorPedidoFilhote);
				}

				vlFormaPagtoDestePedido += Global.arredondaParaMonetario((pedido.pce_prestacao_qtde * pedido.pce_prestacao_valor) * razaoValorPedidoFilhote);
				if (pedido.pce_forma_pagto_prestacao == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
				{
					vlBoletoDestePedido += Global.arredondaParaMonetario((pedido.pce_prestacao_qtde * pedido.pce_prestacao_valor) * razaoValorPedidoFilhote);
				}
				#endregion
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA)
			{
				#region [ Parcelado sem entrada ]
				vlFormaPagtoDestePedido = Global.arredondaParaMonetario(pedido.pse_prim_prest_valor * razaoValorPedidoFilhote);
				if (pedido.pse_forma_pagto_prim_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
				{
					vlBoletoDestePedido = Global.arredondaParaMonetario(pedido.pse_prim_prest_valor * razaoValorPedidoFilhote);
				}

				vlFormaPagtoDestePedido += Global.arredondaParaMonetario((pedido.pse_demais_prest_qtde * pedido.pse_demais_prest_valor) * razaoValorPedidoFilhote);
				if (pedido.pse_forma_pagto_demais_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
				{
					vlBoletoDestePedido += Global.arredondaParaMonetario((pedido.pse_demais_prest_qtde * pedido.pse_demais_prest_valor) * razaoValorPedidoFilhote);
				}
				#endregion
			}
			#endregion

			vlDiferencaArredondamento = pedido.vlTotalPrecoNfDestePedido - vlFormaPagtoDestePedido;

			pedido.vlTotalFormaPagtoDestePedido = vlFormaPagtoDestePedido;
			pedido.vlTotalBoletoDestePedido = vlBoletoDestePedido;
			if (Math.Abs(vlDiferencaArredondamento) <= 1) pedido.vlTotalBoletoDestePedido += vlDiferencaArredondamento;
			#endregion

			#region [ Calcula o valor que será pago através de cartão ]
			if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_A_VISTA)
			{
				if (pedido.av_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
				{
					blnCartaoPagtoIntegral = true;
					vlPagtoEmCartao = pedido.vl_total_NF;
				}
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_CARTAO)
			{
				blnCartaoPagtoIntegral = true;
				vlPagtoEmCartao = pedido.pc_qtde_parcelas * pedido.pc_valor_parcela;
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA)
			{
				// NOP
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA)
			{
				if (pedido.pce_forma_pagto_entrada == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
				{
					vlPagtoEmCartao = pedido.pce_entrada_valor;
				}
				if (pedido.pce_forma_pagto_prestacao == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
				{
					vlPagtoEmCartao += pedido.pce_prestacao_qtde * pedido.pce_prestacao_valor;
				}
				if ((pedido.pce_forma_pagto_entrada == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO) && (pedido.pce_forma_pagto_prestacao == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO))
				{
					blnCartaoPagtoIntegral = true;
				}
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA)
			{
				if (pedido.pse_forma_pagto_prim_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
				{
					vlPagtoEmCartao = pedido.pse_prim_prest_valor;
				}
				if (pedido.pse_forma_pagto_demais_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
				{
					vlPagtoEmCartao += pedido.pse_demais_prest_qtde * pedido.pse_demais_prest_valor;
				}
				if ((pedido.pse_forma_pagto_prim_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO) && (pedido.pse_forma_pagto_demais_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO))
				{
					blnCartaoPagtoIntegral = true;
				}
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELA_UNICA)
			{
				if (pedido.pu_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
				{
					vlPagtoEmCartao = pedido.pu_valor;
					blnCartaoPagtoIntegral = true;
				}
			}

			if (blnCartaoPagtoIntegral)
			{
				pedido.vlPagtoEmCartao = pedido.vl_total_NF;
			}
			else
			{
				pedido.vlPagtoEmCartao = vlPagtoEmCartao;
			}
			#endregion

			return pedido;
		}
		#endregion

		#region [ getPedidoConsolidadoFamilia ]
		/// <summary>
		/// Retorna um objeto Pedido contendo os dados do pedido consolidados por família de pedidos (tratamento para auto-split)
		/// </summary>
		/// <param name="numeroPedido">
		/// Número do pedido (pedido-pai ou pedido-filhote, o método sempre irá se basear automaticamente no nº do pedido-pai)
		/// </param>
		/// <returns>
		/// Retorna um objeto Pedido contendo os dados do pedido consolidados por família de pedidos (tratamento para auto-split)
		/// </returns>
		public static Pedido getPedidoConsolidadoFamilia(String numeroPedido)
		{
			#region [ Declarações ]
			String strSql;
			String numeroPedidoBase;
			short iSequencia = 0;
			decimal vlBoletoConsolidadoFamilia = 0m;
			decimal vlFormaPagtoConsolidadoFamilia = 0m;
			decimal vlDiferencaArredondamento;
			decimal vlPagtoEmCartao = 0m;
			bool blnCartaoPagtoIntegral = false;
			Pedido pedido = new Pedido();
			PedidoItem pedidoItem;
			PedidoItemDevolvido pedidoItemDevolvido;
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow rowResultado;
			#endregion

			#region [ Consistências ]
			if (numeroPedido == null) throw new Exception("Nº do pedido a ser consultado não foi fornecido!!");
			if (numeroPedido.Length == 0) throw new Exception("Nº do pedido a ser consultado não foi informado!!");
			#endregion

			#region [ Prepara acesso ao BD ]
			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();
			#endregion

			#region [ Inicialização ]
			numeroPedido = numeroPedido.Trim();
			numeroPedidoBase = Global.retornaNumeroPedidoBase(numeroPedido);
			#endregion

			#region [ Pesquisa pedido ]

			#region [ Monta Select ]
			strSql = "SELECT" +
						" t_PEDIDO.*," +
						" t_LOJA.razao_social AS loja_razao_social," +
						" t_LOJA.nome AS loja_nome," +
						" t_USUARIO_VENDEDOR.nome AS vendedor_nome," +
						" t_ORCAMENTISTA_E_INDICADOR.desempenho_nota AS indicador_desempenho_nota" +
					" FROM t_PEDIDO" +
						" INNER JOIN t_LOJA ON (t_PEDIDO.loja=t_LOJA.loja)" +
						" INNER JOIN t_USUARIO AS t_USUARIO_VENDEDOR ON (t_PEDIDO.vendedor=t_USUARIO_VENDEDOR.usuario)" +
						" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" +
					" WHERE" +
						" (pedido = '" + numeroPedidoBase + "')";
			#endregion

			#region [ Executa a consulta ]
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbResultado);
			#endregion

			if (dtbResultado.Rows.Count == 0) throw new Exception("Pedido nº " + numeroPedidoBase + " não foi encontrado!!");

			#region [ Carrega os dados ]
			rowResultado = dtbResultado.Rows[0];
			pedido.pedido = BD.readToString(rowResultado["pedido"]);
			pedido.loja = BD.readToString(rowResultado["loja"]);
			pedido.loja_razao_social = BD.readToString(rowResultado["loja_razao_social"]);
			pedido.loja_nome = BD.readToString(rowResultado["loja_nome"]);
			pedido.data = BD.readToDateTime(rowResultado["data"]);
			pedido.hora = BD.readToString(rowResultado["hora"]);
			pedido.data_hora = BD.readToDateTime(rowResultado["data_hora"]);
			pedido.id_cliente = BD.readToString(rowResultado["id_cliente"]);
			pedido.midia = BD.readToString(rowResultado["midia"]);
			pedido.servicos = BD.readToString(rowResultado["servicos"]);
			pedido.vl_servicos = BD.readToDecimal(rowResultado["vl_servicos"]);
			pedido.vendedor = BD.readToString(rowResultado["vendedor"]);
			pedido.vendedor_nome = BD.readToString(rowResultado["vendedor_nome"]);
			pedido.st_entrega = BD.readToString(rowResultado["st_entrega"]);
			pedido.entregue_data = BD.readToDateTime(rowResultado["entregue_data"]);
			pedido.entregue_usuario = BD.readToString(rowResultado["entregue_usuario"]);
			pedido.cancelado_data = BD.readToDateTime(rowResultado["cancelado_data"]);
			pedido.cancelado_usuario = BD.readToString(rowResultado["cancelado_usuario"]);
			pedido.st_pagto = BD.readToString(rowResultado["st_pagto"]);
			pedido.st_recebido = BD.readToString(rowResultado["st_recebido"]);
			pedido.obs_1 = BD.readToString(rowResultado["obs_1"]);
			pedido.obs_2 = BD.readToString(rowResultado["obs_2"]);
			pedido.qtde_parcelas = BD.readToShort(rowResultado["qtde_parcelas"]);
			pedido.forma_pagto = BD.readToString(rowResultado["forma_pagto"]);
			pedido.vl_total_familia = BD.readToDecimal(rowResultado["vl_total_familia"]);
			pedido.vl_pago_familia = BD.readToDecimal(rowResultado["vl_pago_familia"]);
			pedido.split_status = BD.readToShort(rowResultado["split_status"]);
			pedido.split_data = BD.readToDateTime(rowResultado["split_data"]);
			pedido.split_hora = BD.readToString(rowResultado["split_hora"]);
			pedido.split_usuario = BD.readToString(rowResultado["split_usuario"]);
			pedido.a_entregar_status = BD.readToShort(rowResultado["a_entregar_status"]);
			pedido.a_entregar_data_marcada = BD.readToDateTime(rowResultado["a_entregar_data_marcada"]);
			pedido.a_entregar_data = BD.readToDateTime(rowResultado["a_entregar_data"]);
			pedido.a_entregar_hora = BD.readToString(rowResultado["a_entregar_hora"]);
			pedido.a_entregar_usuario = BD.readToString(rowResultado["a_entregar_usuario"]);
			pedido.loja_indicou = BD.readToString(rowResultado["loja_indicou"]);
			pedido.comissao_loja_indicou = BD.readToSingle(rowResultado["comissao_loja_indicou"]);
			pedido.venda_externa = BD.readToShort(rowResultado["venda_externa"]);
			pedido.vl_frete = BD.readToDecimal(rowResultado["vl_frete"]);
			pedido.transportadora_id = BD.readToString(rowResultado["transportadora_id"]);
			pedido.transportadora_data = BD.readToDateTime(rowResultado["transportadora_data"]);
			pedido.transportadora_usuario = BD.readToString(rowResultado["transportadora_usuario"]);
			pedido.analise_credito = BD.readToShort(rowResultado["analise_credito"]);
			pedido.analise_credito_data = BD.readToDateTime(rowResultado["analise_credito_data"]);
			pedido.analise_credito_usuario = BD.readToString(rowResultado["analise_credito_usuario"]);
			pedido.tipo_parcelamento = BD.readToShort(rowResultado["tipo_parcelamento"]);
			pedido.av_forma_pagto = BD.readToShort(rowResultado["av_forma_pagto"]);
			pedido.pc_qtde_parcelas = BD.readToShort(rowResultado["pc_qtde_parcelas"]);
			pedido.pc_valor_parcela = BD.readToDecimal(rowResultado["pc_valor_parcela"]);
			pedido.pc_maquineta_qtde_parcelas = BD.readToShort(rowResultado["pc_maquineta_qtde_parcelas"]);
			pedido.pc_maquineta_valor_parcela = BD.readToDecimal(rowResultado["pc_maquineta_valor_parcela"]);
			pedido.pce_forma_pagto_entrada = BD.readToShort(rowResultado["pce_forma_pagto_entrada"]);
			pedido.pce_forma_pagto_prestacao = BD.readToShort(rowResultado["pce_forma_pagto_prestacao"]);
			pedido.pce_entrada_valor = BD.readToDecimal(rowResultado["pce_entrada_valor"]);
			pedido.pce_prestacao_qtde = BD.readToShort(rowResultado["pce_prestacao_qtde"]);
			pedido.pce_prestacao_valor = BD.readToDecimal(rowResultado["pce_prestacao_valor"]);
			pedido.pce_prestacao_periodo = BD.readToShort(rowResultado["pce_prestacao_periodo"]);
			pedido.pse_forma_pagto_prim_prest = BD.readToShort(rowResultado["pse_forma_pagto_prim_prest"]);
			pedido.pse_forma_pagto_demais_prest = BD.readToShort(rowResultado["pse_forma_pagto_demais_prest"]);
			pedido.pse_prim_prest_valor = BD.readToDecimal(rowResultado["pse_prim_prest_valor"]);
			pedido.pse_prim_prest_apos = BD.readToShort(rowResultado["pse_prim_prest_apos"]);
			pedido.pse_demais_prest_qtde = BD.readToShort(rowResultado["pse_demais_prest_qtde"]);
			pedido.pse_demais_prest_valor = BD.readToDecimal(rowResultado["pse_demais_prest_valor"]);
			pedido.pse_demais_prest_periodo = BD.readToShort(rowResultado["pse_demais_prest_periodo"]);
			pedido.pu_forma_pagto = BD.readToShort(rowResultado["pu_forma_pagto"]);
			pedido.pu_valor = BD.readToDecimal(rowResultado["pu_valor"]);
			pedido.pu_vencto_apos = BD.readToShort(rowResultado["pu_vencto_apos"]);
			pedido.indicador = BD.readToString(rowResultado["indicador"]);
			pedido.indicador_desempenho_nota = BD.readToString(rowResultado["indicador_desempenho_nota"]);
			pedido.vl_total_NF = BD.readToDecimal(rowResultado["vl_total_NF"]);
			pedido.vl_total_RA = BD.readToDecimal(rowResultado["vl_total_RA"]);
			pedido.perc_RT = BD.readToSingle(rowResultado["perc_RT"]);
			pedido.st_orc_virou_pedido = BD.readToShort(rowResultado["st_orc_virou_pedido"]);
			pedido.orcamento = BD.readToString(rowResultado["orcamento"]);
			pedido.orcamentista = BD.readToString(rowResultado["orcamentista"]);
			pedido.comissao_paga = BD.readToShort(rowResultado["comissao_paga"]);
			pedido.comissao_paga_ult_op = BD.readToString(rowResultado["comissao_paga_ult_op"]);
			pedido.comissao_paga_data = BD.readToDateTime(rowResultado["comissao_paga_data"]);
			pedido.comissao_paga_usuario = BD.readToString(rowResultado["comissao_paga_usuario"]);
			pedido.perc_desagio_RA = BD.readToSingle(rowResultado["perc_desagio_RA"]);
			pedido.perc_limite_RA_sem_desagio = BD.readToSingle(rowResultado["perc_limite_RA_sem_desagio"]);
			pedido.vl_total_RA_liquido = BD.readToDecimal(rowResultado["vl_total_RA_liquido"]);
			pedido.st_tem_desagio_RA = BD.readToShort(rowResultado["st_tem_desagio_RA"]);
			pedido.qtde_parcelas_desagio_RA = BD.readToShort(rowResultado["qtde_parcelas_desagio_RA"]);
			pedido.transportadora_num_coleta = BD.readToString(rowResultado["transportadora_num_coleta"]);
			pedido.transportadora_contato = BD.readToString(rowResultado["transportadora_contato"]);
			pedido.st_end_entrega = BD.readToShort(rowResultado["st_end_entrega"]);
			pedido.endEtg_endereco = BD.readToString(rowResultado["EndEtg_endereco"]);
			pedido.endEtg_endereco_numero = BD.readToString(rowResultado["EndEtg_endereco_numero"]);
			pedido.endEtg_endereco_complemento = BD.readToString(rowResultado["EndEtg_endereco_complemento"]);
			pedido.endEtg_bairro = BD.readToString(rowResultado["EndEtg_bairro"]);
			pedido.endEtg_cidade = BD.readToString(rowResultado["EndEtg_cidade"]);
			pedido.endEtg_uf = BD.readToString(rowResultado["EndEtg_uf"]);
			pedido.endEtg_cep = BD.readToString(rowResultado["EndEtg_cep"]);
			pedido.st_etg_imediata = BD.readToShort(rowResultado["st_etg_imediata"]);
			pedido.etg_imediata_data = BD.readToDateTime(rowResultado["etg_imediata_data"]);
			pedido.etg_imediata_usuario = BD.readToString(rowResultado["etg_imediata_usuario"]);
			pedido.frete_status = BD.readToShort(rowResultado["frete_status"]);
			pedido.frete_valor = BD.readToDecimal(rowResultado["frete_valor"]);
			pedido.frete_data = BD.readToDateTime(rowResultado["frete_data"]);
			pedido.frete_usuario = BD.readToString(rowResultado["frete_usuario"]);
			pedido.stBemUsoConsumo = BD.readToShort(rowResultado["StBemUsoConsumo"]);
			pedido.pedidoRecebidoStatus = BD.readToShort(rowResultado["PedidoRecebidoStatus"]);
			pedido.pedidoRecebidoData = BD.readToDateTime(rowResultado["PedidoRecebidoData"]);
			pedido.pedidoRecebidoUsuarioUltAtualiz = BD.readToString(rowResultado["PedidoRecebidoUsuarioUltAtualiz"]);
			pedido.pedidoRecebidoDtHrUltAtualiz = BD.readToDateTime(rowResultado["PedidoRecebidoDtHrUltAtualiz"]);
			pedido.instaladorInstalaStatus = BD.readToShort(rowResultado["InstaladorInstalaStatus"]);
			pedido.instaladorInstalaUsuarioUltAtualiz = BD.readToString(rowResultado["InstaladorInstalaUsuarioUltAtualiz"]);
			pedido.instaladorInstalaDtHrUltAtualiz = BD.readToDateTime(rowResultado["InstaladorInstalaDtHrUltAtualiz"]);
			pedido.custoFinancFornecTipoParcelamento = BD.readToString(rowResultado["custoFinancFornecTipoParcelamento"]);
			pedido.custoFinancFornecQtdeParcelas = BD.readToShort(rowResultado["custoFinancFornecQtdeParcelas"]);
			pedido.tamanho_num_pedido = BD.readToInt(rowResultado["tamanho_num_pedido"]);
			pedido.pedido_base = BD.readToString(rowResultado["pedido_base"]);
			pedido.st_forma_pagto_somente_cartao = BD.readToByte(rowResultado["st_forma_pagto_somente_cartao"]);
			pedido.id_nfe_emitente = BD.readToInt(rowResultado["id_nfe_emitente"]);
			pedido.st_auto_split = BD.readToByte(rowResultado["st_auto_split"]);
			#endregion

			#endregion

			#region [ Pesquisa itens da família de pedidos ]

			#region [ Monta Select ]
			strSql = "SELECT" +
						" fabricante," +
						" produto," +
						" SUM(qtde) AS total_qtde," +
						" SUM(qtde*preco_venda) AS total_preco_venda," +
						" SUM(qtde*preco_NF) AS total_preco_NF" +
					" FROM t_PEDIDO_ITEM tPI" +
						" INNER JOIN t_PEDIDO tP ON (tPI.pedido = tP.pedido)" +
					" WHERE" +
						" (tPI.pedido LIKE '" + numeroPedidoBase + BD.CARACTER_CURINGA_TODOS + "')" +
						" AND (tP.st_entrega <> '" + Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO + "')" +
					" GROUP BY" +
						" fabricante," +
						" produto" +
					" ORDER BY" +
						" fabricante," +
						" produto";
			#endregion

			#region [ Executa a consulta ]
			cmCommand.CommandText = strSql;
			dtbResultado.Reset();
			daDataAdapter.Fill(dtbResultado);
			#endregion

			if (dtbResultado.Rows.Count == 0) throw new Exception("Itens do pedido nº " + numeroPedidoBase + " não foram encontrados!!");

			#region [ Carrega os dados ]
			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				rowResultado = dtbResultado.Rows[i];
				pedidoItem = new PedidoItem();
				pedidoItem.pedido = numeroPedidoBase;
				pedidoItem.fabricante = BD.readToString(rowResultado["fabricante"]);
				pedidoItem.produto = BD.readToString(rowResultado["produto"]);
				// Não é possível fazer o cast direto p/ short, ocorre um exception
				pedidoItem.qtde = (short)BD.readToInt(rowResultado["total_qtde"]);

				// UTILIZA O PREÇO MÉDIO DE VENDA E NF, POIS OS VALORES PODEM TER SIDO EDITADOS E ESTAREM DIFERENTES NOS PEDIDOS PAI E FILHOTES
				if (pedidoItem.qtde == 0)
				{
					pedidoItem.preco_venda = 0;
					pedidoItem.preco_NF = 0;
				}
				else
				{
					pedidoItem.preco_venda = BD.readToDecimal(rowResultado["total_preco_venda"]) / pedidoItem.qtde;
					pedidoItem.preco_NF = BD.readToDecimal(rowResultado["total_preco_NF"]) / pedidoItem.qtde;
				}

				iSequencia++;
				pedidoItem.sequencia = iSequencia;
				pedido.listaPedidoItem.Add(pedidoItem);
			}

			foreach (PedidoItem item in pedido.listaPedidoItem)
			{
				strSql = "SELECT TOP 1 " +
							"*" +
						" FROM t_PEDIDO_ITEM" +
						" WHERE" +
							" (pedido LIKE '" + numeroPedidoBase + BD.CARACTER_CURINGA_TODOS + "')" +
							" AND (fabricante = '" + item.fabricante + "')" +
							" AND (produto = '" + item.produto + "')" +
						" ORDER BY" +
							" pedido";
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daDataAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count > 0)
				{
					rowResultado = dtbResultado.Rows[0];
					item.desc_dado = BD.readToSingle(rowResultado["desc_dado"]);
					item.preco_venda = BD.readToDecimal(rowResultado["preco_venda"]);
					item.preco_fabricante = BD.readToDecimal(rowResultado["preco_fabricante"]);
					item.preco_lista = BD.readToDecimal(rowResultado["preco_lista"]);
					item.margem = BD.readToSingle(rowResultado["margem"]);
					item.desc_max = BD.readToSingle(rowResultado["desc_max"]);
					item.comissao = BD.readToSingle(rowResultado["comissao"]);
					item.descricao = BD.readToString(rowResultado["descricao"]);
					item.ean = BD.readToString(rowResultado["ean"]);
					item.grupo = BD.readToString(rowResultado["grupo"]);
					item.peso = BD.readToSingle(rowResultado["peso"]);
					item.qtde_volumes = BD.readToShort(rowResultado["qtde_volumes"]);
					item.abaixo_min_status = BD.readToShort(rowResultado["abaixo_min_status"]);
					item.abaixo_min_autorizacao = BD.readToString(rowResultado["abaixo_min_autorizacao"]);
					item.abaixo_min_autorizador = BD.readToString(rowResultado["abaixo_min_autorizador"]);
					item.markup_fabricante = BD.readToSingle(rowResultado["markup_fabricante"]);
					item.preco_NF = BD.readToDecimal(rowResultado["preco_NF"]);
					item.abaixo_min_superv_autorizador = BD.readToString(rowResultado["abaixo_min_superv_autorizador"]);
					item.vl_custo2 = BD.readToDecimal(rowResultado["vl_custo2"]);
					item.descricao_html = BD.readToString(rowResultado["descricao_html"]);
					item.custoFinancFornecCoeficiente = BD.readToSingle(rowResultado["custoFinancFornecCoeficiente"]);
					item.custoFinancFornecPrecoListaBase = BD.readToDecimal(rowResultado["custoFinancFornecPrecoListaBase"]);

					pedido.vlTotalPrecoNfDestePedido += item.qtde * item.preco_NF;
					pedido.vlTotalPrecoVendaDestePedido += item.qtde * item.preco_venda;
				}
			}
			#endregion

			#endregion

			#region [ Pesquisa itens devolvidos ]

			#region [ Monta Select ]
			strSql = "SELECT " +
						"*" +
					" FROM t_PEDIDO_ITEM_DEVOLVIDO" +
					" WHERE" +
						" (pedido LIKE '" + numeroPedidoBase + BD.CARACTER_CURINGA_TODOS + "')" +
					" ORDER BY" +
						" devolucao_data," +
						" devolucao_hora";
			#endregion

			#region [ Executa a consulta ]
			cmCommand.CommandText = strSql;
			dtbResultado.Reset();
			daDataAdapter.Fill(dtbResultado);
			#endregion

			#region [ Carrega os dados ]
			for (int i = 0; i < dtbResultado.Rows.Count; i++)
			{
				rowResultado = dtbResultado.Rows[i];
				pedidoItemDevolvido = new PedidoItemDevolvido();
				pedidoItemDevolvido.id = BD.readToString(rowResultado["id"]);
				pedidoItemDevolvido.devolucao_data = BD.readToDateTime(rowResultado["devolucao_data"]);
				pedidoItemDevolvido.devolucao_hora = BD.readToString(rowResultado["devolucao_hora"]);
				pedidoItemDevolvido.devolucao_usuario = BD.readToString(rowResultado["devolucao_usuario"]);
				pedidoItemDevolvido.pedido = BD.readToString(rowResultado["pedido"]);
				pedidoItemDevolvido.fabricante = BD.readToString(rowResultado["fabricante"]);
				pedidoItemDevolvido.produto = BD.readToString(rowResultado["produto"]);
				pedidoItemDevolvido.qtde = BD.readToShort(rowResultado["qtde"]);
				pedidoItemDevolvido.desc_dado = BD.readToSingle(rowResultado["desc_dado"]);
				pedidoItemDevolvido.preco_venda = BD.readToDecimal(rowResultado["preco_venda"]);
				pedidoItemDevolvido.preco_fabricante = BD.readToDecimal(rowResultado["preco_fabricante"]);
				pedidoItemDevolvido.preco_lista = BD.readToDecimal(rowResultado["preco_lista"]);
				pedidoItemDevolvido.margem = BD.readToSingle(rowResultado["margem"]);
				pedidoItemDevolvido.desc_max = BD.readToSingle(rowResultado["desc_max"]);
				pedidoItemDevolvido.comissao = BD.readToSingle(rowResultado["comissao"]);
				pedidoItemDevolvido.descricao = BD.readToString(rowResultado["descricao"]);
				pedidoItemDevolvido.ean = BD.readToString(rowResultado["ean"]);
				pedidoItemDevolvido.grupo = BD.readToString(rowResultado["grupo"]);
				pedidoItemDevolvido.peso = BD.readToSingle(rowResultado["peso"]);
				pedidoItemDevolvido.qtde_volumes = BD.readToShort(rowResultado["qtde_volumes"]);
				pedidoItemDevolvido.abaixo_min_status = BD.readToShort(rowResultado["abaixo_min_status"]);
				pedidoItemDevolvido.abaixo_min_autorizacao = BD.readToString(rowResultado["abaixo_min_autorizacao"]);
				pedidoItemDevolvido.abaixo_min_autorizador = BD.readToString(rowResultado["abaixo_min_autorizador"]);
				pedidoItemDevolvido.markup_fabricante = BD.readToSingle(rowResultado["markup_fabricante"]);
				pedidoItemDevolvido.motivo = BD.readToString(rowResultado["motivo"]);
				pedidoItemDevolvido.preco_NF = BD.readToDecimal(rowResultado["preco_NF"]);
				pedidoItemDevolvido.comissao_descontada = BD.readToShort(rowResultado["comissao_descontada"]);
				pedidoItemDevolvido.comissao_descontada_ult_op = BD.readToString(rowResultado["comissao_descontada_ult_op"]);
				pedidoItemDevolvido.comissao_descontada_data = BD.readToDateTime(rowResultado["comissao_descontada_data"]);
				pedidoItemDevolvido.comissao_descontada_usuario = BD.readToString(rowResultado["comissao_descontada_usuario"]);
				pedidoItemDevolvido.abaixo_min_superv_autorizador = BD.readToString(rowResultado["abaixo_min_superv_autorizador"]);
				pedidoItemDevolvido.vl_custo2 = BD.readToDecimal(rowResultado["vl_custo2"]);
				pedidoItemDevolvido.descricao_html = BD.readToString(rowResultado["descricao_html"]);
				pedidoItemDevolvido.custoFinancFornecCoeficiente = BD.readToSingle(rowResultado["custoFinancFornecCoeficiente"]);
				pedidoItemDevolvido.custoFinancFornecPrecoListaBase = BD.readToDecimal(rowResultado["custoFinancFornecPrecoListaBase"]);
				pedido.listaPedidoItemDevolvido.Add(pedidoItemDevolvido);

				pedido.vlTotalPrecoNfDestePedido -= pedidoItemDevolvido.qtde * pedidoItemDevolvido.preco_NF;
				pedido.vlTotalPrecoVendaDestePedido -= pedidoItemDevolvido.qtde * pedidoItemDevolvido.preco_venda;
			}
			#endregion

			#endregion

			#region [ Calcula valor total já pago ]

			#region [ Monta Select ]
			strSql = "SELECT" +
						" Coalesce(SUM(valor), 0) AS vl_total" +
					" FROM t_PEDIDO_PAGAMENTO" +
					" WHERE" +
						" (pedido LIKE '" + numeroPedidoBase + BD.CARACTER_CURINGA_TODOS + "')";
			#endregion

			#region [ Executa a consulta ]
			cmCommand.CommandText = strSql;
			dtbResultado.Reset();
			daDataAdapter.Fill(dtbResultado);
			#endregion

			if (dtbResultado.Rows.Count == 0) throw new Exception("Falha ao calcular o valor total já pago!!");

			#region [ Carrega os dados ]
			rowResultado = dtbResultado.Rows[0];
			pedido.vlTotalFamiliaPago = BD.readToDecimal(rowResultado["vl_total"]);
			#endregion

			#endregion

			#region [ Calcula o valor total da família de pedidos ]

			#region [ Monta Select ]
			strSql = "SELECT" +
						" Coalesce(SUM(qtde*preco_venda), 0) AS vl_total," +
						" Coalesce(SUM(qtde*preco_NF), 0) AS vl_total_NF" +
					" FROM t_PEDIDO_ITEM INNER JOIN t_PEDIDO" +
						" ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" +
					" WHERE" +
						" (st_entrega<>'" + Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO + "')" +
						" AND (t_PEDIDO.pedido LIKE '" + numeroPedidoBase + BD.CARACTER_CURINGA_TODOS + "')";
			#endregion

			#region [ Executa a consulta ]
			cmCommand.CommandText = strSql;
			dtbResultado.Reset();
			daDataAdapter.Fill(dtbResultado);
			#endregion

			if (dtbResultado.Rows.Count == 0) throw new Exception("Falha ao calcular o valor total da família de pedidos!!");

			#region [ Carrega os dados ]
			rowResultado = dtbResultado.Rows[0];
			pedido.vlTotalFamiliaPrecoVenda = BD.readToDecimal(rowResultado["vl_total"]);
			pedido.vlTotalFamiliaPrecoNF = BD.readToDecimal(rowResultado["vl_total_NF"]);
			#endregion

			#endregion

			#region [ Calcula o valor total em devoluções da família de pedidos ]

			#region [ Monta Select ]
			strSql = "SELECT" +
						" Coalesce(SUM(qtde*preco_venda), 0) AS vl_total," +
						" Coalesce(SUM(qtde*preco_NF), 0) AS vl_total_NF" +
					" FROM t_PEDIDO_ITEM_DEVOLVIDO" +
					" WHERE" +
						" (pedido LIKE '" + numeroPedidoBase + BD.CARACTER_CURINGA_TODOS + "')";
			#endregion

			#region [ Executa a consulta ]
			cmCommand.CommandText = strSql;
			dtbResultado.Reset();
			daDataAdapter.Fill(dtbResultado);
			#endregion

			if (dtbResultado.Rows.Count == 0) throw new Exception("Falha ao calcular o valor total em devoluções da família de pedidos!!");

			#region [ Carrega os dados ]
			rowResultado = dtbResultado.Rows[0];
			pedido.vlTotalFamiliaDevolucaoPrecoVenda = BD.readToDecimal(rowResultado["vl_total"]);
			pedido.vlTotalFamiliaDevolucaoPrecoNF = BD.readToDecimal(rowResultado["vl_total_NF"]);

			pedido.vlTotalFamiliaPrecoVenda -= pedido.vlTotalFamiliaDevolucaoPrecoVenda;
			pedido.vlTotalFamiliaPrecoNF -= pedido.vlTotalFamiliaDevolucaoPrecoNF;
			#endregion

			#endregion

			#region [ Calcula o valor em boletos (consolidado por família de pedidos) ]

			#region [ Calcula o valor proporcional, pois pode ser um pedido filhote ]
			if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_A_VISTA)
			{
				#region [ À vista ]
				vlFormaPagtoConsolidadoFamilia = pedido.vlTotalFamiliaPrecoNF;
				if ((pedido.av_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
					|| (pedido.av_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO_AV))
				{
					vlBoletoConsolidadoFamilia = pedido.vlTotalFamiliaPrecoNF;
				}
				#endregion
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELA_UNICA)
			{
				#region [ Parcela única ]
				vlFormaPagtoConsolidadoFamilia = Global.arredondaParaMonetario(pedido.pu_valor);
				if (pedido.pu_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
				{
					vlBoletoConsolidadoFamilia = Global.arredondaParaMonetario(pedido.pu_valor);
				}
				#endregion
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA)
			{
				#region [ Parcelado com entrada ]
				vlFormaPagtoConsolidadoFamilia = Global.arredondaParaMonetario(pedido.pce_entrada_valor);
				if (pedido.pce_forma_pagto_entrada == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
				{
					vlBoletoConsolidadoFamilia = Global.arredondaParaMonetario(pedido.pce_entrada_valor);
				}

				vlFormaPagtoConsolidadoFamilia += Global.arredondaParaMonetario(pedido.pce_prestacao_qtde * pedido.pce_prestacao_valor);
				if (pedido.pce_forma_pagto_prestacao == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
				{
					vlBoletoConsolidadoFamilia += Global.arredondaParaMonetario(pedido.pce_prestacao_qtde * pedido.pce_prestacao_valor);
				}
				#endregion
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA)
			{
				#region [ Parcelado sem entrada ]
				vlFormaPagtoConsolidadoFamilia = Global.arredondaParaMonetario(pedido.pse_prim_prest_valor);
				if (pedido.pse_forma_pagto_prim_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
				{
					vlBoletoConsolidadoFamilia = Global.arredondaParaMonetario(pedido.pse_prim_prest_valor);
				}

				vlFormaPagtoConsolidadoFamilia += Global.arredondaParaMonetario(pedido.pse_demais_prest_qtde * pedido.pse_demais_prest_valor);
				if (pedido.pse_forma_pagto_demais_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_BOLETO)
				{
					vlBoletoConsolidadoFamilia += Global.arredondaParaMonetario(pedido.pse_demais_prest_qtde * pedido.pse_demais_prest_valor);
				}
				#endregion
			}
			#endregion

			vlDiferencaArredondamento = pedido.vlTotalPrecoNfDestePedido - vlFormaPagtoConsolidadoFamilia;

			pedido.vlTotalFormaPagtoDestePedido = vlFormaPagtoConsolidadoFamilia;
			pedido.vlTotalBoletoDestePedido = vlBoletoConsolidadoFamilia;
			if (Math.Abs(vlDiferencaArredondamento) <= 1) pedido.vlTotalBoletoDestePedido += vlDiferencaArredondamento;
			#endregion

			#region [ Calcula o valor que será pago através de cartão ]
			if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_A_VISTA)
			{
				if (pedido.av_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
				{
					blnCartaoPagtoIntegral = true;
					vlPagtoEmCartao = pedido.vl_total_NF;
				}
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_CARTAO)
			{
				blnCartaoPagtoIntegral = true;
				vlPagtoEmCartao = pedido.pc_qtde_parcelas * pedido.pc_valor_parcela;
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA)
			{
				// NOP
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA)
			{
				if (pedido.pce_forma_pagto_entrada == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
				{
					vlPagtoEmCartao = pedido.pce_entrada_valor;
				}
				if (pedido.pce_forma_pagto_prestacao == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
				{
					vlPagtoEmCartao += pedido.pce_prestacao_qtde * pedido.pce_prestacao_valor;
				}
				if ((pedido.pce_forma_pagto_entrada == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO) && (pedido.pce_forma_pagto_prestacao == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO))
				{
					blnCartaoPagtoIntegral = true;
				}
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA)
			{
				if (pedido.pse_forma_pagto_prim_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
				{
					vlPagtoEmCartao = pedido.pse_prim_prest_valor;
				}
				if (pedido.pse_forma_pagto_demais_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
				{
					vlPagtoEmCartao += pedido.pse_demais_prest_qtde * pedido.pse_demais_prest_valor;
				}
				if ((pedido.pse_forma_pagto_prim_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO) && (pedido.pse_forma_pagto_demais_prest == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO))
				{
					blnCartaoPagtoIntegral = true;
				}
			}
			else if (pedido.tipo_parcelamento == Global.Cte.TipoParcelamentoPedido.COD_FORMA_PAGTO_PARCELA_UNICA)
			{
				if (pedido.pu_forma_pagto == Global.Cte.CodFormaPagtoPedido.ID_FORMA_PAGTO_CARTAO)
				{
					vlPagtoEmCartao = pedido.pu_valor;
					blnCartaoPagtoIntegral = true;
				}
			}

			if (blnCartaoPagtoIntegral)
			{
				pedido.vlPagtoEmCartao = pedido.vl_total_NF;
			}
			else
			{
				pedido.vlPagtoEmCartao = vlPagtoEmCartao;
			}
			#endregion

			return pedido;
		}
		#endregion

		#region [ atualizaPedidoCancela ]
		public static bool atualizaPedidoCancela(String usuario,
												 String pedido,
												 String cancelado_auto_motivo,
												 String codigo_sub_motivo,
												 String cancelado_motivo,
												 out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.atualizaPedidoCancela()";
			string strMsg;
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (pedido == null)
				{
					strMsgErro = "Não foi informado o número do pedido!";
					return false;
				}

				if (pedido.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o número do pedido!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePedidoCancela.Parameters["@pedido"].Value = pedido;
				cmUpdatePedidoCancela.Parameters["@st_entrega"].Value = Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO;
				cmUpdatePedidoCancela.Parameters["@cancelado_usuario"].Value = usuario;
				cmUpdatePedidoCancela.Parameters["@cancelado_auto_motivo"].Value = Texto.leftStr(cancelado_auto_motivo, 160);
				cmUpdatePedidoCancela.Parameters["@cancelado_codigo_motivo"].Value = Global.Cte.PedidoCanceladoCodigoMotivo.CANCELAMENTO_AUTOMATICO.GetValue();
				cmUpdatePedidoCancela.Parameters["@cancelado_codigo_sub_motivo"].Value = (codigo_sub_motivo == null ? "" : codigo_sub_motivo);
				cmUpdatePedidoCancela.Parameters["@cancelado_motivo"].Value = Texto.leftStr(cancelado_motivo, 800);
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePedidoCancela);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					strMsgErro = NOME_DESTA_ROTINA + " - Tentativa resultou em exception!!\n" + ex.ToString();
					Global.gravaLogAtividade(strMsgErro);
				}

				if (intRetorno == 1)
				{
					blnSucesso = true;
					strMsg = "Pedido " + pedido + " cancelado automaticamente (motivo: " + cancelado_auto_motivo + ")";
					GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_CANCELAMENTO_AUTOMATICO_PEDIDO, pedido, strMsg, out strMsgErro);
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
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao tentar atualizar o pedido com o status 'cancelado'!!" + strMsgErro;
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ atualizaPedidoStatusAnaliseCredito ]
		public static bool atualizaPedidoStatusAnaliseCredito(string pedido,
													int? statusAnaliseCredito,
													string usuario,
													out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.atualizaPedidoStatusAnaliseCredito()";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (pedido == null)
				{
					strMsgErro = "Não foi informado o número do pedido!!";
					return false;
				}

				if (pedido.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o número do pedido!!";
					return false;
				}

				if (statusAnaliseCredito == null)
				{
					strMsgErro = "Não foi informado o novo status da análise de crédito!!";
					return false;
				}
				#endregion

				if (usuario == null) usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				if (usuario.Trim().Length == 0) usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;

				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePedidoStatusAnaliseCredito.Parameters["@pedido"].Value = pedido;
				cmUpdatePedidoStatusAnaliseCredito.Parameters["@analise_credito"].Value = statusAnaliseCredito;
				cmUpdatePedidoStatusAnaliseCredito.Parameters["@analise_credito_usuario"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePedidoStatusAnaliseCredito);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					strMsgErro = NOME_DESTA_ROTINA + " - Tentativa resultou em exception!!\n" + ex.ToString();
					Global.gravaLogAtividade(strMsgErro);
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
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao tentar atualizar o status da análise de crédito do pedido (pedido=" + pedido + ")!!" + strMsgErro;
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ atualizaPedidoStatusPagto ]
		public static bool atualizaPedidoStatusPagto(string pedido,
													string st_pagto,
													out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.atualizaPedidoStatusPagto()";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (pedido == null)
				{
					strMsgErro = "Não foi informado o número do pedido!!";
					return false;
				}

				if (pedido.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o número do pedido!!";
					return false;
				}

				if (st_pagto == null)
				{
					strMsgErro = "Não foi informado o novo status de pagamento!!";
					return false;
				}

				if (st_pagto.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o novo status de pagamento!!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePedidoStatusPagto.Parameters["@pedido"].Value = pedido;
				cmUpdatePedidoStatusPagto.Parameters["@st_pagto"].Value = st_pagto;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePedidoStatusPagto);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					strMsgErro = NOME_DESTA_ROTINA + " - Tentativa resultou em exception!!\n" + ex.ToString();
					Global.gravaLogAtividade(strMsgErro);
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
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao tentar atualizar o status de pagamento do pedido (pedido=" + pedido + ")!!" + strMsgErro;
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ atualizaPedidoVlPagoFamilia ]
		public static bool atualizaPedidoVlPagoFamilia(string pedido,
													decimal vl_pago_familia,
													out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.atualizaPedidoVlPagoFamilia()";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (pedido == null)
				{
					strMsgErro = "Não foi informado o número do pedido!!";
					return false;
				}

				if (pedido.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o número do pedido!!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePedidoVlPagoFamilia.Parameters["@pedido"].Value = pedido;
				cmUpdatePedidoVlPagoFamilia.Parameters["@vl_pago_familia"].Value = vl_pago_familia;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePedidoVlPagoFamilia);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					strMsgErro = NOME_DESTA_ROTINA + " - Tentativa resultou em exception!!\n" + ex.ToString();
					Global.gravaLogAtividade(strMsgErro);
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
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao tentar atualizar o valor pago (família de pedidos) do pedido (pedido=" + pedido + ")!!" + strMsgErro;
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ atualizaPedidoStEntrega ]
		public static bool atualizaPedidoStEntrega(String pedido,
													String st_entrega,
													out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.atualizaPedidoStEntrega()";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistências ]
				if (pedido == null)
				{
					strMsgErro = "Não foi informado o número do pedido!!";
					return false;
				}

				if (pedido.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o número do pedido!!";
					return false;
				}

				if (st_entrega == null)
				{
					strMsgErro = "Não foi informado o novo status de entrega!!";
					return false;
				}

				if (st_entrega.ToString().Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o novo status de entrega!!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePedidoStEntrega.Parameters["@pedido"].Value = pedido;
				cmUpdatePedidoStEntrega.Parameters["@st_entrega"].Value = st_entrega;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePedidoStEntrega);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					strMsgErro = NOME_DESTA_ROTINA + " - Tentativa resultou em exception!!\n" + ex.ToString();
					Global.gravaLogAtividade(strMsgErro);
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
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao tentar atualizar o status de entrega do pedido (pedido=" + pedido + ")!!" + strMsgErro;
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ atualizaPedidoBaseCamposValoresTotaisDevidoCancelamento ]
		public static bool atualizaPedidoBaseCamposValoresTotaisDevidoCancelamento(
										String pedido,
										Decimal vl_total_familia,
										Decimal vl_total_NF,
										Decimal vl_total_RA,
										Decimal vl_total_RA_liquido,
										int qtde_parcelas_desagio_RA,
										int st_tem_desagio_RA,
										out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.atualizaPedidoBaseCamposValoresTotaisDevidoCancelamento()";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			#region [ Inicialização dos parâmetros ]
			strMsgErro = "";
			#endregion

			try
			{
				#region [ Consistências ]
				if (pedido == null)
				{
					strMsgErro = "Não foi informado o número do pedido!!";
					return false;
				}

				if (pedido.Trim().Length == 0)
				{
					strMsgErro = "Não foi fornecido o número do pedido!!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.Parameters["@pedido"].Value = pedido;
				cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.Parameters["@vl_total_familia"].Value = vl_total_familia;
				cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.Parameters["@vl_total_NF"].Value = vl_total_NF;
				cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.Parameters["@vl_total_RA"].Value = vl_total_RA;
				cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.Parameters["@vl_total_RA_liquido"].Value = vl_total_RA_liquido;
				cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.Parameters["@qtde_parcelas_desagio_RA"].Value = qtde_parcelas_desagio_RA;
				cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento.Parameters["@st_tem_desagio_RA"].Value = st_tem_desagio_RA;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePedidoBaseCamposValoresTotaisDevidoCancelamento);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					strMsgErro = NOME_DESTA_ROTINA + " - Tentativa resultou em exception!!\n" + ex.ToString();
					Global.gravaLogAtividade(strMsgErro);
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
					if (strMsgErro.Length > 0) strMsgErro = "\n" + strMsgErro;
					strMsgErro = "Falha ao tentar atualizar os valores totais do pedido-base devido ao cancelamento do pedido " + pedido + "!!" + strMsgErro;
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ executaCancelamentoAutomaticoPedidos ]
		public static bool executaCancelamentoAutomaticoPedidos(out string strMsgInformativa, out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.executaCancelamentoAutomaticoPedidos()";
			int qtdeParcelasDesagioRA;
			int st_tem_desagio_RA;
			int qtdePedidosCancelados = 0;
			Decimal vlTotalFamiliaPrecoVenda;
			Decimal vlTotalFamiliaPrecoNF;
			Decimal vlTotalFamiliaPago;
			Decimal vlTotalFamiliaDevolucaoPrecoVenda;
			Decimal vlTotalFamiliaDevolucaoPrecoNF;
			Decimal vlTotalRA;
			Decimal vlTotalRALiquido;
			String st_pagto;
			String strMsg;
			String strBlocoNotas;
			String strSql;
			String strSqlVlPagoCartao;
			String strWhereBase;
			String strWhereLojasIgnoradas = "";
			String strPedido;
			String strPedidoBase;
			String strCanceladoAutoMotivo;
			String strCanceladoMotivo;
			String strCodigoSubMotivo;
			String strOrigemRegistro;
			String strLog;
			StringBuilder sbPedidoCancelado = new StringBuilder("");
			bool blnSucesso;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
			DateTime dtHrInicio = DateTime.Now;
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

				#region [ Monta o SQL da consulta ]
				// Lembrando que o status 'Pendente Cartão de Crédito' é um status lógico definido pela combinação do status da análise de crédito 'ST_INICIAL' + forma de pagamento usando somente pagamento por cartão.
				// Além disso, nesse caso ainda não foi realizada nenhuma ação de análise de crédito, portanto, os campos analise_credito_data e analise_credito_data_sem_hora estão com o valor NULL, motivo pelo qual se usa a data do cadastramento do pedido.
				strSqlVlPagoCartao = " Coalesce(" +
									"(" +
									"SELECT" +
										" SUM(payment.valor_transacao)" +
									" FROM t_PAGTO_GW_PAG pag INNER JOIN t_PAGTO_GW_PAG_PAYMENT payment ON (pag.id = payment.id_pagto_gw_pag)" +
									" WHERE" +
										" (pag.pedido = t1.pedido_base)" +
										" AND" +
										"(" +
											" (ult_GlobalStatus = '" + Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue() + "')" +
											" OR" +
											" (ult_GlobalStatus = '" + Global.Cte.Braspag.Pagador.GlobalStatus.AUTORIZADA.GetValue() + "')" +
										")" +
									"), 0) AS vl_pago_cartao";

				#region [ Restrição para lojas ignoradas ]
				for (int i = 0; i < Global.Parametros.Geral.CancelamentoAutomaticoPedidosLojasIgnoradas.Count; i++)
				{
					if (Global.Parametros.Geral.CancelamentoAutomaticoPedidosLojasIgnoradas[i] > 0)
					{
						if (strWhereLojasIgnoradas.Length > 0) strWhereLojasIgnoradas += ",";
						strWhereLojasIgnoradas += Global.Parametros.Geral.CancelamentoAutomaticoPedidosLojasIgnoradas[i].ToString();
					}
				}

				if (strWhereLojasIgnoradas.Length > 0) strWhereLojasIgnoradas = " AND (tPedBase.numero_loja NOT IN (" + strWhereLojasIgnoradas + "))";
				#endregion

				strWhereBase = " (t1.st_entrega <> '" + Global.Cte.StEntregaPedido.ST_ENTREGA_ENTREGUE + "')" +
								" AND (t1.st_entrega <> '" + Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO + "')" +
								" AND (t1.st_entrega <> '" + Global.Cte.StEntregaPedido.ST_ENTREGA_A_ENTREGAR + "')" +
								" AND (Coalesce(tPedBase.st_pagto, '') <> '" + Global.Cte.StPagtoPedido.ST_PAGTO_PAGO + "')" +
								" AND (Coalesce(tPedBase.st_pagto, '') <> '" + Global.Cte.StPagtoPedido.ST_PAGTO_PARCIAL + "')" +
								strWhereLojasIgnoradas;

				// A ORDENAÇÃO DO RESULTADO AGRUPA POR:
				//		1) STATUS DA ANÁLISE DE CRÉDITO
				//		2) DATA DA ÚLTIMA ALTERAÇÃO DA ANÁLISE DE CRÉDITO
				//		3) Nº PEDIDO-BASE (PARA AGRUPAR PEDIDOS AUTO-SPLITADOS)
				//		4) TAMANHO DO Nº DO PEDIDO EM ORDEM DECRESCENTE (PARA CANCELAR PRIMEIRO OS PEDIDOS-FILHOTE E POR ÚLTIMO O PEDIDO-BASE)
				//		5) Nº PEDIDO EM ORDEM DECRESCENTE (PARA CANCELAR PRIMEIRO OS PEDIDOS-FILHOTE C/ SUFIXO DE MAIOR VALOR)
				strSql = "SELECT " +
							"*" +
						" FROM (" +
							"SELECT" +
								" t1.id_nfe_emitente," +
								" t1.pedido," +
								" t1.pedido_base," +
								" Coalesce(t1.obs_2, '') AS obs_2," +
								" t1.transportadora_selecao_auto_status," +
								" Coalesce(t1.transportadora_id, '') AS transportadora_id," +
								" t1.st_entrega," +
								" '" + Global.Cte.PrazoCancelAutoPedidoEmDias.PENDENTE_CARTAO_CREDITO.GetName() + "' AS origem_registro," +
								" 'Pendente Cartão de Crédito' AS analise_credito_descricao," +
								" " + Global.Cte.PrazoCancelAutoPedidoEmDias.PENDENTE_CARTAO_CREDITO.ToString() + " AS prazo_cancelamento," +
								" tPedBase.analise_credito," +
								" tPedBase.data_hora AS analise_credito_data," +
								" tPedBase.data AS analise_credito_data_sem_hora," +
								" Coalesce(Datediff(day, tPedBase.data, Convert(datetime, Convert(varchar(10), getdate(), 121), 121)), 0) AS dias_decorridos," +
								" (" +
									"SELECT Count(*) FROM t_PEDIDO t2 WHERE (t2.pedido_base = t1.pedido_base) AND (t2.st_auto_split = 0) AND (t2.tamanho_num_pedido > " + Global.Cte.Etc.TAM_MIN_ID_PEDIDO.ToString() + ")" +
								") AS qtde_pedido_filhote_manual," +
								strSqlVlPagoCartao +
							" FROM t_PEDIDO t1" +
								" INNER JOIN t_PEDIDO AS tPedBase ON (t1.pedido_base=tPedBase.pedido)" +
							" WHERE" +
								strWhereBase +
								" AND (" +
									"(tPedBase.analise_credito = " + Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.ST_INICIAL.ToString() + ") AND (tPedBase.st_forma_pagto_somente_cartao = 1)" +
									" AND (Coalesce(Datediff(day, tPedBase.data, " + Global.sqlMontaGetdateSomenteData() + "), 0) > " + Global.Cte.PrazoCancelAutoPedidoEmDias.PENDENTE_CARTAO_CREDITO.ToString() + ")" +
								")" +
							" UNION " +
							"SELECT" +
								" t1.id_nfe_emitente," +
								" t1.pedido," +
								" t1.pedido_base," +
								" Coalesce(t1.obs_2, '') AS obs_2," +
								" t1.transportadora_selecao_auto_status," +
								" Coalesce(t1.transportadora_id, '') AS transportadora_id," +
								" t1.st_entrega," +
								" '" + Global.Cte.PrazoCancelAutoPedidoEmDias.CREDITO_OK_AGUARDANDO_DEPOSITO.GetName() + "' AS origem_registro," +
								" 'Crédito OK (aguardando depósito)' AS analise_credito_descricao," +
								" " + Global.Cte.PrazoCancelAutoPedidoEmDias.CREDITO_OK_AGUARDANDO_DEPOSITO.ToString() + " AS prazo_cancelamento," +
								" tPedBase.analise_credito," +
								" tPedBase.analise_credito_data," +
								" tPedBase.analise_credito_data_sem_hora," +
								" Coalesce(Datediff(day, tPedBase.analise_credito_data_sem_hora, Convert(datetime, Convert(varchar(10), getdate(), 121), 121)), 0) AS dias_decorridos," +
								" (" +
									"SELECT Count(*) FROM t_PEDIDO t2 WHERE (t2.pedido_base = t1.pedido_base) AND (t2.st_auto_split = 0) AND (t2.tamanho_num_pedido > " + Global.Cte.Etc.TAM_MIN_ID_PEDIDO.ToString() + ")" +
								") AS qtde_pedido_filhote_manual," +
								strSqlVlPagoCartao +
							" FROM t_PEDIDO t1" +
								" INNER JOIN t_PEDIDO AS tPedBase ON (t1.pedido_base=tPedBase.pedido)" +
							" WHERE" +
								strWhereBase +
								" AND (" +
									"(tPedBase.analise_credito = " + Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK_AGUARDANDO_DEPOSITO.ToString() + ")" +
									" AND (tPedBase.analise_credito_data_sem_hora IS NOT NULL)" +
									" AND (Coalesce(Datediff(day, tPedBase.analise_credito_data_sem_hora, " + Global.sqlMontaGetdateSomenteData() + "), 0) > " + Global.Cte.PrazoCancelAutoPedidoEmDias.CREDITO_OK_AGUARDANDO_DEPOSITO.ToString() + ")" +
								")" +
							" UNION " +
							"SELECT" +
								" t1.id_nfe_emitente," +
								" t1.pedido," +
								" t1.pedido_base," +
								" Coalesce(t1.obs_2, '') AS obs_2," +
								" t1.transportadora_selecao_auto_status," +
								" Coalesce(t1.transportadora_id, '') AS transportadora_id," +
								" t1.st_entrega," +
								" '" + Global.Cte.PrazoCancelAutoPedidoEmDias.PENDENTE_VENDAS.GetName() + "' AS origem_registro," +
								" 'Pendente Vendas' AS analise_credito_descricao," +
								" " + Global.Cte.PrazoCancelAutoPedidoEmDias.PENDENTE_VENDAS.ToString() + " AS prazo_cancelamento," +
								" tPedBase.analise_credito," +
								" tPedBase.analise_credito_data," +
								" tPedBase.analise_credito_data_sem_hora," +
								" Coalesce(Datediff(day, tPedBase.analise_credito_data_sem_hora, Convert(datetime, Convert(varchar(10), getdate(), 121), 121)), 0) AS dias_decorridos," +
								" (" +
									"SELECT Count(*) FROM t_PEDIDO t2 WHERE (t2.pedido_base = t1.pedido_base) AND (t2.st_auto_split = 0) AND (t2.tamanho_num_pedido > " + Global.Cte.Etc.TAM_MIN_ID_PEDIDO.ToString() + ")" +
								") AS qtde_pedido_filhote_manual," +
								strSqlVlPagoCartao +
							" FROM t_PEDIDO t1" +
								" INNER JOIN t_PEDIDO AS tPedBase ON (t1.pedido_base=tPedBase.pedido)" +
							" WHERE" +
								strWhereBase +
								" AND (" +
									"(tPedBase.analise_credito = " + Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS.ToString() + ")" +
									" AND (tPedBase.analise_credito_data_sem_hora IS NOT NULL)" +
									" AND (Coalesce(Datediff(day, tPedBase.analise_credito_data_sem_hora, " + Global.sqlMontaGetdateSomenteData() + "), 0) > " + Global.Cte.PrazoCancelAutoPedidoEmDias.PENDENTE_VENDAS.ToString() + ")" +
								")" +
							") t" +
						" WHERE" +
							" (qtde_pedido_filhote_manual = 0)" +
							" AND (LEN(obs_2) = 0)" +
							" AND (vl_pago_cartao = 0)" +
							" AND ((transportadora_selecao_auto_status = 1) OR (LEN(Coalesce(transportadora_id,'')) = 0))" +
						" ORDER BY" +
							" analise_credito," +
							" analise_credito_data_sem_hora," +
							" pedido_base," +
							" LEN(pedido) DESC," +
							" pedido DESC";
				#endregion

				#region [ Log informativo da consulta realizada ]
				strMsg = NOME_DESTA_ROTINA + ":\r\n" + strSql;
				Global.gravaLogAtividade(strMsg);
				#endregion

				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				#endregion

				blnSucesso = false;
				try
				{
					BD.iniciaTransacao();

					for (int i = 0; i < dtbConsulta.Rows.Count; i++)
					{
						qtdePedidosCancelados++;

						rowConsulta = dtbConsulta.Rows[i];
						strPedido = BD.readToString(rowConsulta["pedido"]);

						strMsg = "Processando cancelamento automático do pedido: " + strPedido;
						Global.gravaLogAtividade(strMsg);

						#region [ Determina o código do sub-motivo do cancelamento ]
						// O código do motivo é fixo: 001 - Cancelamento Automático
						// O código do sub-motivo indica a situação do status da análise de crédito
						strOrigemRegistro = BD.readToString(rowConsulta["origem_registro"]);
						if (strOrigemRegistro.Equals(Global.Cte.PrazoCancelAutoPedidoEmDias.PENDENTE_CARTAO_CREDITO.GetName()))
						{
							strCodigoSubMotivo = "001";
						}
						else if (strOrigemRegistro.Equals(Global.Cte.PrazoCancelAutoPedidoEmDias.CREDITO_OK_AGUARDANDO_DEPOSITO.GetName()))
						{
							strCodigoSubMotivo = "002";
						}
						else if (strOrigemRegistro.Equals(Global.Cte.PrazoCancelAutoPedidoEmDias.PENDENTE_VENDAS.GetName()))
						{
							strCodigoSubMotivo = "003";
						}
						else
						{
							strCodigoSubMotivo = "";
						}
						#endregion

						// Campo geral para todos os pedidos cancelados, seja de forma automática ou manual (800 caracteres)
						strCanceladoMotivo = "Pedido cancelado automaticamente por exceder o prazo máximo de " +
										BD.readToInt(rowConsulta["prazo_cancelamento"]).ToString() +
										" dias consecutivos na situação '" + BD.readToString(rowConsulta["analise_credito_descricao"]) +
										"' (desde " + Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(rowConsulta["analise_credito_data_sem_hora"])) + ")";

						// Campo exclusivo para pedidos cancelados de forma automática (160 caracteres)
						strCanceladoAutoMotivo = "Excedeu o prazo máximo de " +
									BD.readToInt(rowConsulta["prazo_cancelamento"]).ToString() +
									" dias consecutivos na situação '" + BD.readToString(rowConsulta["analise_credito_descricao"]) +
									"' (desde " + Global.formataDataDdMmYyyyComSeparador(BD.readToDateTime(rowConsulta["analise_credito_data_sem_hora"])) + ")";
						if (!PedidoDAO.atualizaPedidoCancela(Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA, strPedido, strCanceladoAutoMotivo, strCodigoSubMotivo, strCanceladoMotivo, out strMsgErro))
						{
							throw new Exception("Falha ao atualizar o registro do pedido " + strPedido + " para o status 'cancelado'!!\n" + strMsgErro);
						}
						if (sbPedidoCancelado.Length > 0) sbPedidoCancelado.Append(", ");
						sbPedidoCancelado.Append(strPedido);

						if (!EstoqueDAO.estoquePedidoCancela(Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA, strPedido, out strLog, out strMsgErro))
						{
							throw new Exception("Falha ao processar o cancelamento do pedido " + strPedido + " no estoque!!\n" + strMsgErro);
						}

						if (!EstoqueDAO.estoqueProcessaProdutosVendidosSemPresenca(BD.readToInt(rowConsulta["id_nfe_emitente"]), Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA, out strMsgErro))
						{
							throw new Exception("Falha ao processar os produtos vendidos sem presença no estoque devido ao cancelamento do pedido " + strPedido + "!!\n" + strMsgErro);
						}

						if (!calculaPagamentos(strPedido, out vlTotalFamiliaPrecoVenda, out vlTotalFamiliaPrecoNF, out vlTotalFamiliaPago, out vlTotalFamiliaDevolucaoPrecoVenda, out vlTotalFamiliaDevolucaoPrecoNF, out st_pagto, out strMsgErro))
						{
							throw new Exception("Falha ao calcular os pagamentos devido ao cancelamento do pedido " + strPedido + "!!\n" + strMsgErro);
						}

						vlTotalRA = vlTotalFamiliaPrecoNF - vlTotalFamiliaPrecoVenda;

						if (!calculaTotalRALiquidoBD(strPedido, out vlTotalRALiquido, out strMsgErro))
						{
							throw new Exception("Falha ao calcular o valor total do RA líquido devido ao cancelamento do pedido " + strPedido + "!!\n" + strMsgErro);
						}

						strPedidoBase = Global.retornaNumeroPedidoBase(strPedido);

						qtdeParcelasDesagioRA = 0;
						st_tem_desagio_RA = (vlTotalRA != 0) ? 1 : 0;
						if (!atualizaPedidoBaseCamposValoresTotaisDevidoCancelamento(strPedidoBase, vlTotalFamiliaPrecoVenda, vlTotalFamiliaPrecoNF, vlTotalRA, vlTotalRALiquido, qtdeParcelasDesagioRA, st_tem_desagio_RA, out strMsgErro))
						{
							throw new Exception("Falha ao atualizar os valores totais do pedido-base devido ao cancelamento do pedido " + strPedido + "!!\n" + strMsgErro);
						}

						strBlocoNotas = strCanceladoMotivo;
						if (!gravaPedidoBlocoNotas(strPedido, Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA, Global.Cte.BlocoNotasPedidoNivelAcesso.COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__PUBLICO, strBlocoNotas, out strMsgErro))
						{
							throw new Exception("Falha ao gravar o bloco de notas do pedido " + strPedido + "!!\n" + strMsgErro);
						}
					} // for (int i = 0; i < dtbConsulta.Rows.Count; i++)

					blnSucesso = true;
				}
				catch (Exception ex)
				{
					strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
					Global.gravaLogAtividade(strMsgErro);
					blnSucesso = false;
				}

				if (blnSucesso)
				{
					BD.commitTransacao();
					strMsgInformativa = qtdePedidosCancelados.ToString() + " pedidos cancelados";
					strMsg = "Rotina " + NOME_DESTA_ROTINA + " concluída com sucesso (duração: " + Global.formataDuracaoHMS(DateTime.Now - dtHrInicio) + ")";
					Global.gravaLogAtividade(strMsg);
					return true;
				}
				else
				{
					BD.rollbackTransacao();
					strMsgErro = "Rotina " + NOME_DESTA_ROTINA + ": falha no processamento e reversão da transação!!";
					return false;
				}
			}
			catch (Exception ex)
			{
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.ToString();
				Global.gravaLogAtividade(strMsgErro);
				return false;
			}
		}
		#endregion

		#region [ calculaPagamentos ]
		public static bool calculaPagamentos(String pedido,
											 out Decimal vlTotalFamiliaPrecoVenda,
											 out Decimal vlTotalFamiliaPrecoNF,
											 out Decimal vlTotalFamiliaPago,
											 out Decimal vlTotalFamiliaDevolucaoPrecoVenda,
											 out Decimal vlTotalFamiliaDevolucaoPrecoNF,
											 out String st_pagto,
											 out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.calculaPagamentos()";
			String strSql;
			String strPedidoBase;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
			#endregion

			#region [ Inicialização dos parâmetros ]
			vlTotalFamiliaPrecoVenda = 0;
			vlTotalFamiliaPrecoNF = 0;
			vlTotalFamiliaPago = 0;
			vlTotalFamiliaDevolucaoPrecoVenda = 0;
			vlTotalFamiliaDevolucaoPrecoNF = 0;
			st_pagto = "";
			strMsgErro = "";
			#endregion

			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				strPedidoBase = Global.retornaNumeroPedidoBase(pedido);

				#region [ Obtém o status de pagamento ]
				strSql = "SELECT" +
							" pedido," +
							" st_pagto" +
						" FROM t_PEDIDO" +
						" WHERE" +
							" (pedido = '" + strPedidoBase + "')";
				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				#endregion

				if (dtbConsulta.Rows.Count == 0)
				{
					strMsgErro = "Pedido-base " + strPedidoBase + " não foi encontrado.";
					return false;
				}

				rowConsulta = dtbConsulta.Rows[0];

				st_pagto = BD.readToString(rowConsulta["st_pagto"]);
				#endregion

				#region [ Obtém valor total já pago ]
				strSql = "SELECT" +
							" Coalesce(SUM(valor), 0) AS total" +
						" FROM t_PEDIDO_PAGAMENTO" +
						" WHERE" +
							" (pedido LIKE '" + strPedidoBase + BD.CARACTER_CURINGA_TODOS + "')";
				dtbConsulta.Reset();
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);

				if (dtbConsulta.Rows.Count > 0)
				{
					rowConsulta = dtbConsulta.Rows[0];
					vlTotalFamiliaPago = BD.readToDecimal(rowConsulta["total"]);
				}
				#endregion

				#region [ Obtém total preço venda e total NF ]
				strSql = "SELECT" +
							" Coalesce(SUM(qtde*preco_venda), 0) AS total," +
							" Coalesce(SUM(qtde*preco_NF), 0) AS total_NF" +
						" FROM t_PEDIDO_ITEM" +
							" INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" +
						" WHERE" +
							" (st_entrega <> '" + Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO + "')" +
							" AND (t_PEDIDO.pedido LIKE '" + strPedidoBase + BD.CARACTER_CURINGA_TODOS + "')";
				dtbConsulta.Reset();
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				if (dtbConsulta.Rows.Count > 0)
				{
					rowConsulta = dtbConsulta.Rows[0];
					vlTotalFamiliaPrecoVenda = BD.readToDecimal(rowConsulta["total"]);
					vlTotalFamiliaPrecoNF = BD.readToDecimal(rowConsulta["total_NF"]);
				}
				#endregion

				#region [ Obtém total preço venda e total NF das devoluções ]
				strSql = "SELECT" +
							" Coalesce(SUM(qtde*preco_venda), 0) AS total," +
							" Coalesce(SUM(qtde*preco_NF), 0) AS total_NF" +
						" FROM t_PEDIDO_ITEM_DEVOLVIDO" +
						" WHERE" +
							" (pedido LIKE '" + strPedidoBase + BD.CARACTER_CURINGA_TODOS + "')";
				dtbConsulta.Reset();
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				if (dtbConsulta.Rows.Count > 0)
				{
					rowConsulta = dtbConsulta.Rows[0];
					vlTotalFamiliaDevolucaoPrecoVenda = BD.readToDecimal(rowConsulta["total"]);
					vlTotalFamiliaDevolucaoPrecoNF = BD.readToDecimal(rowConsulta["total_NF"]);
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

		#region [ calculaTotalRALiquidoBD ]
		public static bool calculaTotalRALiquidoBD(String pedido, out Decimal vlTotalRALiquido, out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.calculaTotalRALiquidoBD()";
			double percentualDesagioRALiquido;
			Decimal vlTotalRA;
			String strSql;
			String strPedidoBase;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbConsulta = new DataTable();
			DataRow rowConsulta;
			#endregion

			#region [ Inicialização dos parâmetros ]
			vlTotalRALiquido = 0;
			strMsgErro = "";
			#endregion

			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				strPedidoBase = Global.retornaNumeroPedidoBase(pedido);

				#region [ Obtém perc_desagio_RA_liquida ]
				strSql = "SELECT " +
							"*" +
						" FROM t_PEDIDO" +
						" WHERE" +
							" (pedido = '" + strPedidoBase + "')";
				#region [ Executa a consulta no BD ]
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				#endregion

				if (dtbConsulta.Rows.Count == 0)
				{
					strMsgErro = "Pedido-base " + strPedidoBase + " não foi encontrado.";
					return false;
				}

				rowConsulta = dtbConsulta.Rows[0];

				percentualDesagioRALiquido = BD.readToSingle(rowConsulta["perc_desagio_RA_liquida"]);
				#endregion

				#region [ Obtém os valores totais de NF, RA e venda ]
				vlTotalRA = 0;
				strSql = "SELECT" +
							" Coalesce(SUM(qtde*(preco_NF-preco_venda)), 0) AS total_RA" +
						" FROM t_PEDIDO_ITEM" +
							" INNER JOIN t_PEDIDO ON (t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido)" +
						" WHERE" +
							" (st_entrega<>'" + Global.Cte.StEntregaPedido.ST_ENTREGA_CANCELADO + "')" +
							" AND (t_PEDIDO.pedido LIKE '" + strPedidoBase + BD.CARACTER_CURINGA_TODOS + "')";
				dtbConsulta.Reset();
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbConsulta);
				if (dtbConsulta.Rows.Count > 0)
				{
					rowConsulta = dtbConsulta.Rows[0];
					vlTotalRA = BD.readToDecimal(rowConsulta["total_RA"]);
				}
				#endregion

				vlTotalRALiquido = vlTotalRA - vlTotalRA * (Decimal)(percentualDesagioRALiquido / 100);
				vlTotalRALiquido = Global.arredondaParaMonetario(vlTotalRALiquido);

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

		#region [ gravaPedidoBlocoNotas ]
		public static bool gravaPedidoBlocoNotas(String pedido,
												 String usuario,
												 int nivel_acesso,
												 String mensagem,
												 out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.gravaPedidoBlocoNotas()";
			bool blnSucesso = false;
			int intNsu;
			int intQtdeTentativas = 0;
			int intRetorno;
			StringBuilder sbLog = new StringBuilder("");
			#endregion

			try
			{
				#region [ Laço de tentativas de inserção no banco de dados ]
				do
				{
					intQtdeTentativas++;
					strMsgErro = "";

					#region [ Obtém NSU ]
					if (!BD.geraNsuUsandoTabelaFinControle(Global.Cte.Nsu.T_PEDIDO_BLOCO_NOTAS, out intNsu, out strMsgErro)) intNsu = 0;
					#endregion

					if (intNsu > 0)
					{
						#region [ Preenche o valor dos parâmetros ]
						cmInsertPedidoBlocoNotas.Parameters["@id"].Value = intNsu;
						cmInsertPedidoBlocoNotas.Parameters["@pedido"].Value = pedido;
						cmInsertPedidoBlocoNotas.Parameters["@usuario"].Value = usuario;
						cmInsertPedidoBlocoNotas.Parameters["@nivel_acesso"].Value = nivel_acesso;
						cmInsertPedidoBlocoNotas.Parameters["@mensagem"].Value = ((mensagem == null) ? "" : Texto.leftStr(mensagem, 4000));
						#endregion

						#region [ Monta texto para o log em arquivo ]
						// Se houver conteúdo de alguma tentativa anterior, descarta
						sbLog = new StringBuilder("");
						foreach (SqlParameter item in cmInsertPedidoBlocoNotas.Parameters)
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
						}
						#endregion

						#region [ Tenta inserir o registro ]
						try
						{
							intRetorno = BD.executaNonQuery(ref cmInsertPedidoBlocoNotas);
						}
						catch (Exception ex)
						{
							intRetorno = 0;
							Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: Dados do registro do bloco de notas do pedido = " + sbLog.ToString() + "\n" + ex.ToString());
						}
						#endregion

						#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
						if (intRetorno == 1)
						{
							blnSucesso = true;
						}
						else
						{
							Thread.Sleep(100);
						}
						#endregion
					}
				} while ((!blnSucesso) && (intQtdeTentativas < BD.MAX_TENTATIVAS_INSERT_BD));
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao gravar o bloco de notas do pedido " + pedido + " após " + intQtdeTentativas.ToString() + " tentativas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Falha: Dados do registro do bloco de notas do pedido = " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ gravaPedidoBlocoNotasComDataHora ]
		public static bool gravaPedidoBlocoNotasComDataHora(DateTime dataHoraCadastro,
												 String pedido,
												 String usuario,
												 int nivel_acesso,
												 String mensagem,
												 out String strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.gravaPedidoBlocoNotasComDataHora()";
			bool blnSucesso = false;
			int intNsu;
			int intQtdeTentativas = 0;
			int intRetorno;
			DateTime dtHrCadastroBlocoNotas;
			StringBuilder sbLog = new StringBuilder("");
			#endregion

			try
			{
				dtHrCadastroBlocoNotas = dataHoraCadastro;
				if (dtHrCadastroBlocoNotas == DateTime.MinValue) dtHrCadastroBlocoNotas = DateTime.Now;

				#region [ Laço de tentativas de inserção no banco de dados ]
				do
				{
					intQtdeTentativas++;
					strMsgErro = "";

					#region [ Obtém NSU ]
					if (!BD.geraNsuUsandoTabelaFinControle(Global.Cte.Nsu.T_PEDIDO_BLOCO_NOTAS, out intNsu, out strMsgErro)) intNsu = 0;
					#endregion

					if (intNsu > 0)
					{
						#region [ Preenche o valor dos parâmetros ]
						cmInsertComDataHoraPedidoBlocoNotas.Parameters["@id"].Value = intNsu;
						cmInsertComDataHoraPedidoBlocoNotas.Parameters["@dt_cadastro"].Value = Global.formataDataYyyyMmDdComSeparador(dtHrCadastroBlocoNotas);
						cmInsertComDataHoraPedidoBlocoNotas.Parameters["@dt_hr_cadastro"].Value = Global.formataDataYyyyMmDdHhMmSsComSeparador(dtHrCadastroBlocoNotas);
						cmInsertComDataHoraPedidoBlocoNotas.Parameters["@pedido"].Value = pedido;
						cmInsertComDataHoraPedidoBlocoNotas.Parameters["@usuario"].Value = usuario;
						cmInsertComDataHoraPedidoBlocoNotas.Parameters["@nivel_acesso"].Value = nivel_acesso;
						cmInsertComDataHoraPedidoBlocoNotas.Parameters["@mensagem"].Value = ((mensagem == null) ? "" : Texto.leftStr(mensagem, 4000));
						#endregion

						#region [ Monta texto para o log em arquivo ]
						// Se houver conteúdo de alguma tentativa anterior, descarta
						sbLog = new StringBuilder("");
						foreach (SqlParameter item in cmInsertComDataHoraPedidoBlocoNotas.Parameters)
						{
							if (sbLog.Length > 0) sbLog.Append("; ");
							sbLog.Append(item.ParameterName + "=" + (item.Value != null ? item.Value.ToString() : ""));
						}
						#endregion

						#region [ Tenta inserir o registro ]
						try
						{
							intRetorno = BD.executaNonQuery(ref cmInsertComDataHoraPedidoBlocoNotas);
						}
						catch (Exception ex)
						{
							intRetorno = 0;
							Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Tentativa " + intQtdeTentativas.ToString() + " resultou em exception: Dados do registro do bloco de notas do pedido = " + sbLog.ToString() + "\n" + ex.ToString());
						}
						#endregion

						#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
						if (intRetorno == 1)
						{
							blnSucesso = true;
						}
						else
						{
							Thread.Sleep(100);
						}
						#endregion
					}
				} while ((!blnSucesso) && (intQtdeTentativas < BD.MAX_TENTATIVAS_INSERT_BD));
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao gravar o bloco de notas do pedido " + pedido + " após " + intQtdeTentativas.ToString() + " tentativas!!";
					return false;
				}
				#endregion
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(NOME_DESTA_ROTINA + " - Falha: Dados do registro do bloco de notas do pedido = " + sbLog.ToString() + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ inserePedidoPagamento ]
		public static bool inserePedidoPagamento(PedidoPagamento pedidoPagto, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.inserePedidoPagamento()";
			int id_emailsndsvc_mensagem;
			int intRetorno;
			bool blnGerarNsu = false;
			bool blnGerouNsu;
			string msg_erro_aux;
			string nsuPedidoPagto = "";
			string strMsg;
			string strSubject;
			string strBody;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Gera NSU ]
				if (pedidoPagto.id == null)
				{
					blnGerarNsu = true;
				}
				else if (pedidoPagto.id.Trim().Length == 0)
				{
					blnGerarNsu = true;
				}

				if (blnGerarNsu)
				{
					blnGerouNsu = GeralDAO.geraNsuUsandoTabelaControle(Global.Cte.Nsu.NSU_PEDIDO_PAGAMENTO, out nsuPedidoPagto, out msg_erro_aux);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro de pagamento (t_PEDIDO_PAGAMENTO)!!\n" + msg_erro_aux;
						return false;
					}
					pedidoPagto.id = nsuPedidoPagto;
				}
				else
				{
					// O NSU já foi gerado anteriormente na rotina chamadora
					nsuPedidoPagto = pedidoPagto.id;
				}
				#endregion

				#region [ Campos opcionais já vieram preenchidos da rotina chamadora? ]
				if (pedidoPagto.data == null)
				{
					pedidoPagto.data = DateTime.Now;
					pedidoPagto.hora = Global.digitos(Global.formataHoraHhMmSsComSeparador(DateTime.Now));
				}

				if (pedidoPagto.data == DateTime.MinValue)
				{
					pedidoPagto.data = DateTime.Now;
					pedidoPagto.hora = Global.digitos(Global.formataHoraHhMmSsComSeparador(DateTime.Now));
				}

				if (pedidoPagto.usuario == null) pedidoPagto.usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				if (pedidoPagto.usuario.Trim().Length == 0) pedidoPagto.usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmInsertPedidoPagamento.Parameters["@id"].Value = pedidoPagto.id;
				cmInsertPedidoPagamento.Parameters["@pedido"].Value = pedidoPagto.pedido;
				cmInsertPedidoPagamento.Parameters["@data"].Value = Global.formataDataYyyyMmDdComSeparador(pedidoPagto.data);
				cmInsertPedidoPagamento.Parameters["@hora"].Value = pedidoPagto.hora;
				cmInsertPedidoPagamento.Parameters["@valor"].Value = pedidoPagto.valor;
				cmInsertPedidoPagamento.Parameters["@tipo_pagto"].Value = pedidoPagto.tipo_pagto;
				cmInsertPedidoPagamento.Parameters["@usuario"].Value = pedidoPagto.usuario;
				cmInsertPedidoPagamento.Parameters["@id_pagto_gw_pag_payment"].Value = pedidoPagto.id_pagto_gw_pag_payment;
				#endregion

				#region [ Tenta inserir o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmInsertPedidoPagamento);
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
					svcLog.complemento_1 = Global.serializaObjectToXml(pedidoPagto);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": falha ao tentar gravar o pagamento no pedido [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\r\n" + NOME_DESTA_ROTINA + ": falha ao tentar gravar o pagamento no pedido " + pedidoPagto.pedido + " devido a processamento da transação t_PAGTO_GW_PAG_PAYMENT.id=" + pedidoPagto.id_pagto_gw_pag_payment.ToString() + " (tipo de pagamento: " + pedidoPagto.tipo_pagto + ")\r\n" + ex.ToString();
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
					msg_erro = "Falha ao tentar gravar o pagamento em t_PEDIDO_PAGAMENTO!!\n" + msg_erro;
					return false;
				}
				#endregion

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Sucesso na gravação dos dados (t_PEDIDO_PAGAMENTO.id=" + pedidoPagto.id + ")");

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(pedidoPagto);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": falha ao tentar gravar o pagamento no pedido [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
				strBody = "Mensagem de Financeiro Service\r\n" + NOME_DESTA_ROTINA + ": falha ao tentar gravar o pagamento no pedido " + pedidoPagto.pedido + " devido a processamento da transação t_PAGTO_GW_PAG_PAYMENT.id=" + pedidoPagto.id_pagto_gw_pag_payment.ToString() + " (tipo de pagamento: " + pedidoPagto.tipo_pagto + ")\r\n" + ex.ToString();
				if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
				{
					strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
					Global.gravaLogAtividade(strMsg);
				}

				return false;
			}
		}
		#endregion

		#region [ inserePedidoPagamentoBoletoEC ]
		public static bool inserePedidoPagamentoBoletoEC(PedidoPagamento pedidoPagto, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.inserePedidoPagamentoBoletoEC()";
			int id_emailsndsvc_mensagem;
			int intRetorno;
			bool blnGerarNsu = false;
			bool blnGerouNsu;
			string msg_erro_aux;
			string nsuPedidoPagto = "";
			string strMsg;
			string strSubject;
			string strBody;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Gera NSU ]
				if (pedidoPagto.id == null)
				{
					blnGerarNsu = true;
				}
				else if (pedidoPagto.id.Trim().Length == 0)
				{
					blnGerarNsu = true;
				}

				if (blnGerarNsu)
				{
					blnGerouNsu = GeralDAO.geraNsuUsandoTabelaControle(Global.Cte.Nsu.NSU_PEDIDO_PAGAMENTO, out nsuPedidoPagto, out msg_erro_aux);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro de pagamento (t_PEDIDO_PAGAMENTO)!!\n" + msg_erro_aux;
						return false;
					}
					pedidoPagto.id = nsuPedidoPagto;
				}
				else
				{
					// O NSU já foi gerado anteriormente na rotina chamadora
					nsuPedidoPagto = pedidoPagto.id;
				}
				#endregion

				#region [ Campos opcionais já vieram preenchidos da rotina chamadora? ]
				if (pedidoPagto.data == null)
				{
					pedidoPagto.data = DateTime.Now;
					pedidoPagto.hora = Global.digitos(Global.formataHoraHhMmSsComSeparador(DateTime.Now));
				}

				if (pedidoPagto.data == DateTime.MinValue)
				{
					pedidoPagto.data = DateTime.Now;
					pedidoPagto.hora = Global.digitos(Global.formataHoraHhMmSsComSeparador(DateTime.Now));
				}

				if (pedidoPagto.usuario == null) pedidoPagto.usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				if (pedidoPagto.usuario.Trim().Length == 0) pedidoPagto.usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmInsertPedidoPagamentoBoletoEC.Parameters["@id"].Value = pedidoPagto.id;
				cmInsertPedidoPagamentoBoletoEC.Parameters["@pedido"].Value = pedidoPagto.pedido;
				cmInsertPedidoPagamentoBoletoEC.Parameters["@data"].Value = Global.formataDataYyyyMmDdComSeparador(pedidoPagto.data);
				cmInsertPedidoPagamentoBoletoEC.Parameters["@hora"].Value = pedidoPagto.hora;
				cmInsertPedidoPagamentoBoletoEC.Parameters["@valor"].Value = pedidoPagto.valor;
				cmInsertPedidoPagamentoBoletoEC.Parameters["@tipo_pagto"].Value = pedidoPagto.tipo_pagto;
				cmInsertPedidoPagamentoBoletoEC.Parameters["@usuario"].Value = pedidoPagto.usuario;
				cmInsertPedidoPagamentoBoletoEC.Parameters["@id_braspag_webhook_complementar"].Value = pedidoPagto.id_braspag_webhook_complementar;
				#endregion

				#region [ Tenta inserir o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmInsertPedidoPagamentoBoletoEC);
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
					svcLog.complemento_1 = Global.serializaObjectToXml(pedidoPagto);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": falha ao tentar gravar o pagamento no pedido [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\r\n" + NOME_DESTA_ROTINA + ": falha ao tentar gravar o pagamento no pedido " + pedidoPagto.pedido + " devido a processamento da transação t_BRASPAG_WEBHOOK_COMPLEMENTAR.Id=" + pedidoPagto.id_braspag_webhook_complementar.ToString() + " (tipo de pagamento: " + pedidoPagto.tipo_pagto + ")\r\n" + ex.ToString();
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
					msg_erro = "Falha ao tentar gravar o pagamento em t_PEDIDO_PAGAMENTO!!\n" + msg_erro;
					return false;
				}
				#endregion

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Sucesso na gravação dos dados (t_PEDIDO_PAGAMENTO.id=" + pedidoPagto.id + ")");

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(pedidoPagto);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": falha ao tentar gravar o pagamento no pedido [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
				strBody = "Mensagem de Financeiro Service\r\n" + NOME_DESTA_ROTINA + ": falha ao tentar gravar o pagamento no pedido " + pedidoPagto.pedido + " devido a processamento da transação t_BRASPAG_WEBHOOK_COMPLEMENTAR.Id=" + pedidoPagto.id_braspag_webhook_complementar.ToString() + " (tipo de pagamento: " + pedidoPagto.tipo_pagto + ")\r\n" + ex.ToString();
				if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
				{
					strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
					Global.gravaLogAtividade(strMsg);
				}

				return false;
			}
		}
		#endregion

		#region [ insereFinPedidoHistPagto ]
		public static bool insereFinPedidoHistPagto(PedidoHistPagto histPagto, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.insereFinPedidoHistPagto()";
			int id_emailsndsvc_mensagem;
			int intRetorno;
			bool blnGerarNsu = false;
			bool blnGerouNsu;
			string msg_erro_aux;
			int idNsu;
			string strMsg;
			string strSubject;
			string strBody;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Gera NSU ]
				if (histPagto.id == 0)
				{
					blnGerarNsu = true;
				}

				if (blnGerarNsu)
				{
					blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.Nsu.T_FIN_PEDIDO_HIST_PAGTO, out idNsu, out msg_erro_aux);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro histórico de pagamento (" + Global.Cte.Nsu.T_FIN_PEDIDO_HIST_PAGTO + ")!!\n" + msg_erro_aux;
						return false;
					}
					histPagto.id = idNsu;
				}
				else
				{
					// O NSU já foi gerado anteriormente na rotina chamadora
					idNsu = histPagto.id;
				}
				#endregion

				#region [ Campos opcionais já vieram preenchidos da rotina chamadora? ]
				if (histPagto.usuario_cadastro == null) histPagto.usuario_cadastro = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				if (histPagto.usuario_cadastro.Trim().Length == 0) histPagto.usuario_cadastro = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				if (histPagto.usuario_ult_atualizacao == null) histPagto.usuario_ult_atualizacao = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				if (histPagto.usuario_ult_atualizacao.Trim().Length == 0) histPagto.usuario_ult_atualizacao = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmInsertFinPedidoHistPagto.Parameters["@id"].Value = histPagto.id;
				cmInsertFinPedidoHistPagto.Parameters["@pedido"].Value = histPagto.pedido;
				cmInsertFinPedidoHistPagto.Parameters["@status"].Value = histPagto.status;
				cmInsertFinPedidoHistPagto.Parameters["@ctrl_pagto_id_parcela"].Value = histPagto.ctrl_pagto_id_parcela;
				cmInsertFinPedidoHistPagto.Parameters["@ctrl_pagto_modulo"].Value = histPagto.ctrl_pagto_modulo;
				cmInsertFinPedidoHistPagto.Parameters["@valor_total"].Value = histPagto.valor_total;
				cmInsertFinPedidoHistPagto.Parameters["@valor_rateado"].Value = histPagto.valor_rateado;
				cmInsertFinPedidoHistPagto.Parameters["@descricao"].Value = histPagto.descricao;
				cmInsertFinPedidoHistPagto.Parameters["@usuario_cadastro"].Value = histPagto.usuario_cadastro;
				cmInsertFinPedidoHistPagto.Parameters["@usuario_ult_atualizacao"].Value = histPagto.usuario_ult_atualizacao;
				#endregion

				#region [ Tenta inserir o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmInsertFinPedidoHistPagto);
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
					svcLog.complemento_1 = Global.serializaObjectToXml(histPagto);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": falha ao tentar gravar o registro de histórico de pagamento no pedido " + histPagto.pedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\r\n" + NOME_DESTA_ROTINA + ": falha ao tentar gravar o registro de histórico de pagamento no pedido " + histPagto.pedido + " devido a processamento da transação t_PAGTO_GW_PAG_PAYMENT.id=" + histPagto.ctrl_pagto_id_parcela.ToString() + "\r\n" + ex.ToString();
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
					msg_erro = "Falha ao tentar gravar o registro de histórico de pagamento em t_FIN_PEDIDO_HIST_PAGTO!!\n" + msg_erro;
					return false;
				}
				#endregion

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Sucesso na gravação dos dados (t_FIN_PEDIDO_HIST_PAGTO.id=" + histPagto.id + ")");

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(histPagto);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": falha ao tentar gravar o registro de histórico de pagamento no pedido " + histPagto.pedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
				strBody = "Mensagem de Financeiro Service\r\n" + NOME_DESTA_ROTINA + ": falha ao tentar gravar o registro de histórico de pagamento no pedido " + histPagto.pedido + " devido a processamento da transação t_PAGTO_GW_PAG_PAYMENT.id=" + histPagto.ctrl_pagto_id_parcela.ToString() + "\r\n" + ex.ToString();
				if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
				{
					strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
					Global.gravaLogAtividade(strMsg);
				}

				return false;
			}
		}
		#endregion

		#region [ insereFinPedidoHistPagtoBoletoEC ]
		public static bool insereFinPedidoHistPagtoBoletoEC(PedidoHistPagto histPagto, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.insereFinPedidoHistPagtoBoletoEC()";
			int id_emailsndsvc_mensagem;
			int intRetorno;
			bool blnGerarNsu = false;
			bool blnGerouNsu;
			string msg_erro_aux;
			int idNsu;
			string strMsg;
			string strSubject;
			string strBody;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Gera NSU ]
				if (histPagto.id == 0)
				{
					blnGerarNsu = true;
				}

				if (blnGerarNsu)
				{
					blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.Nsu.T_FIN_PEDIDO_HIST_PAGTO, out idNsu, out msg_erro_aux);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro histórico de pagamento (" + Global.Cte.Nsu.T_FIN_PEDIDO_HIST_PAGTO + ")!!\n" + msg_erro_aux;
						return false;
					}
					histPagto.id = idNsu;
				}
				else
				{
					// O NSU já foi gerado anteriormente na rotina chamadora
					idNsu = histPagto.id;
				}
				#endregion

				#region [ Campos opcionais já vieram preenchidos da rotina chamadora? ]
				if (histPagto.usuario_cadastro == null) histPagto.usuario_cadastro = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				if (histPagto.usuario_cadastro.Trim().Length == 0) histPagto.usuario_cadastro = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				if (histPagto.usuario_ult_atualizacao == null) histPagto.usuario_ult_atualizacao = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				if (histPagto.usuario_ult_atualizacao.Trim().Length == 0) histPagto.usuario_ult_atualizacao = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmInsertFinPedidoHistPagtoBoletoEC.Parameters["@id"].Value = histPagto.id;
				cmInsertFinPedidoHistPagtoBoletoEC.Parameters["@pedido"].Value = histPagto.pedido;
				cmInsertFinPedidoHistPagtoBoletoEC.Parameters["@status"].Value = histPagto.status;
				cmInsertFinPedidoHistPagtoBoletoEC.Parameters["@ctrl_pagto_id_parcela"].Value = histPagto.ctrl_pagto_id_parcela;
				cmInsertFinPedidoHistPagtoBoletoEC.Parameters["@ctrl_pagto_modulo"].Value = histPagto.ctrl_pagto_modulo;
				cmInsertFinPedidoHistPagtoBoletoEC.Parameters["@dt_vencto"].Value = Global.formataDataYyyyMmDdComSeparador(histPagto.dt_vencto);
				cmInsertFinPedidoHistPagtoBoletoEC.Parameters["@dt_credito"].Value = Global.formataDataYyyyMmDdComSeparador(histPagto.dt_credito);
				cmInsertFinPedidoHistPagtoBoletoEC.Parameters["@valor_total"].Value = histPagto.valor_total;
				cmInsertFinPedidoHistPagtoBoletoEC.Parameters["@valor_rateado"].Value = histPagto.valor_rateado;
				cmInsertFinPedidoHistPagtoBoletoEC.Parameters["@valor_pago"].Value = histPagto.valor_pago;
				cmInsertFinPedidoHistPagtoBoletoEC.Parameters["@descricao"].Value = histPagto.descricao;
				cmInsertFinPedidoHistPagtoBoletoEC.Parameters["@usuario_cadastro"].Value = histPagto.usuario_cadastro;
				cmInsertFinPedidoHistPagtoBoletoEC.Parameters["@usuario_ult_atualizacao"].Value = histPagto.usuario_ult_atualizacao;
				#endregion

				#region [ Tenta inserir o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmInsertFinPedidoHistPagtoBoletoEC);
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
					svcLog.complemento_1 = Global.serializaObjectToXml(histPagto);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": falha ao tentar gravar o registro de histórico de pagamento no pedido " + histPagto.pedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\r\n" + NOME_DESTA_ROTINA + ": falha ao tentar gravar o registro de histórico de pagamento no pedido " + histPagto.pedido + " devido a processamento da transação t_BRASPAG_WEBHOOK_COMPLEMENTAR.Id=" + histPagto.ctrl_pagto_id_parcela.ToString() + "\r\n" + ex.ToString();
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
					msg_erro = "Falha ao tentar gravar o registro de histórico de pagamento em t_FIN_PEDIDO_HIST_PAGTO!!\n" + msg_erro;
					return false;
				}
				#endregion

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Sucesso na gravação dos dados (t_FIN_PEDIDO_HIST_PAGTO.id=" + histPagto.id + ")");

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(histPagto);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": falha ao tentar gravar o registro de histórico de pagamento no pedido " + histPagto.pedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
				strBody = "Mensagem de Financeiro Service\r\n" + NOME_DESTA_ROTINA + ": falha ao tentar gravar o registro de histórico de pagamento no pedido " + histPagto.pedido + " devido a processamento da transação t_BRASPAG_WEBHOOK_COMPLEMENTAR.Id=" + histPagto.ctrl_pagto_id_parcela.ToString() + "\r\n" + ex.ToString();
				if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
				{
					strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
					Global.gravaLogAtividade(strMsg);
				}

				return false;
			}
		}
		#endregion

		#region [ updateFinPedidoHistPagtoCampoStatus ]
		public static bool updateFinPedidoHistPagtoCampoStatus(PedidoHistPagto histPagto, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.updateFinPedidoHistPagtoCampoStatus()";
			int id_emailsndsvc_mensagem;
			int intRetorno;
			string msg_erro_aux;
			string strMsg;
			string strSubject;
			string strBody;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Consistências ]
				if (histPagto.id <= 0)
				{
					msg_erro = "ID do registro da tabela " + Global.Cte.Nsu.T_FIN_PEDIDO_HIST_PAGTO + " não foi informado!";
					return false;
				}
				#endregion

				#region [ Campos opcionais já vieram preenchidos da rotina chamadora? ]
				if (histPagto.usuario_ult_atualizacao == null) histPagto.usuario_ult_atualizacao = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				if (histPagto.usuario_ult_atualizacao.Trim().Length == 0) histPagto.usuario_ult_atualizacao = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdateFinPedidoHistPagtoCampoStatus.Parameters["@id"].Value = histPagto.id;
				cmUpdateFinPedidoHistPagtoCampoStatus.Parameters["@status"].Value = histPagto.status;
				cmUpdateFinPedidoHistPagtoCampoStatus.Parameters["@usuario_ult_atualizacao"].Value = histPagto.usuario_ult_atualizacao;
				#endregion

				#region [ Tenta atualizar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateFinPedidoHistPagtoCampoStatus);
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
					svcLog.complemento_1 = Global.serializaObjectToXml(histPagto);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": falha ao tentar atualizar o registro de histórico de pagamento (" + Global.Cte.Nsu.T_FIN_PEDIDO_HIST_PAGTO + ".id=" + histPagto.id.ToString() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\r\n" + NOME_DESTA_ROTINA + ": falha ao tentar atualizar o registro de histórico de pagamento (" + Global.Cte.Nsu.T_FIN_PEDIDO_HIST_PAGTO + ".id=" + histPagto.id.ToString() + ")\r\n" + ex.ToString();
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
					msg_erro = "Falha ao tentar atualizar o registro de histórico de pagamento em t_FIN_PEDIDO_HIST_PAGTO!!\n" + msg_erro;
					return false;
				}
				#endregion

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Sucesso na atualização dos dados (t_FIN_PEDIDO_HIST_PAGTO.id=" + histPagto.id + ")");

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(histPagto);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": falha ao tentar atualizar o registro de histórico de pagamento no pedido (" + Global.Cte.Nsu.T_FIN_PEDIDO_HIST_PAGTO + ".id=" + histPagto.id.ToString() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
				strBody = "Mensagem de Financeiro Service\r\n" + NOME_DESTA_ROTINA + ": falha ao tentar atualizar o registro de histórico de pagamento no pedido (" + Global.Cte.Nsu.T_FIN_PEDIDO_HIST_PAGTO + ".id=" + histPagto.id.ToString() + ")\r\n" + ex.ToString();
				if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
				{
					strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
					Global.gravaLogAtividade(strMsg);
				}

				return false;
			}
		}
		#endregion

		#region [ updatePedidoStPedidoNovoAnaliseCreditoMsgAlertaFlagAtivo ]
		public static bool updatePedidoStPedidoNovoAnaliseCreditoMsgAlertaFlagAtivo(string numeroPedido, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.updatePedidoStPedidoNovoAnaliseCreditoMsgAlertaFlagAtivo()";
			int id_emailsndsvc_mensagem;
			int intRetorno;
			string msg_erro_aux;
			string strMsg;
			string strSubject;
			string strBody;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Consistências ]
				if ((numeroPedido ?? "").Length == 0)
				{
					msg_erro = "O número do pedido não foi informado!";
					return false;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmUpdateStPedidoNovoAnaliseCreditoMsgAlerta.Parameters["@pedido"].Value = numeroPedido;
				cmUpdateStPedidoNovoAnaliseCreditoMsgAlerta.Parameters["@st_pedido_novo_analise_credito_msg_alerta"].Value = 1;
				#endregion

				#region [ Tenta atualizar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateStPedidoNovoAnaliseCreditoMsgAlerta);
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
					svcLog.complemento_1 = numeroPedido;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": falha ao tentar atualizar o campo t_PEDIDO.st_pedido_novo_analise_credito_msg_alerta do pedido " + numeroPedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\r\n" + NOME_DESTA_ROTINA + ": falha ao tentar atualizar o campo t_PEDIDO.st_pedido_novo_analise_credito_msg_alerta do pedido " + numeroPedido + "\r\n" + ex.ToString();
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
					msg_erro = "Falha ao tentar atualizar o campo t_PEDIDO.st_pedido_novo_analise_credito_msg_alerta do pedido " + numeroPedido + "\n" + msg_erro;
					return false;
				}
				#endregion

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Sucesso na atualização do campo t_PEDIDO.st_pedido_novo_analise_credito_msg_alerta do pedido " + numeroPedido);

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = numeroPedido;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				strSubject = Global.montaIdInstanciaServicoEmailSubject() + " " + NOME_DESTA_ROTINA + ": falha ao tentar atualizar o campo t_PEDIDO.st_pedido_novo_analise_credito_msg_alerta do pedido " + numeroPedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
				strBody = "Mensagem de Financeiro Service\r\n" + NOME_DESTA_ROTINA + ": falha ao tentar atualizar o campo t_PEDIDO.st_pedido_novo_analise_credito_msg_alerta do pedido " + numeroPedido + "\r\n" + ex.ToString();
				if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
				{
					strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
					Global.gravaLogAtividade(strMsg);
				}

				return false;
			}
		}
		#endregion

		#region [ getPedidoHistPagtoByCtrlPagtoIdParcela ]
		public static PedidoHistPagto getPedidoHistPagtoByCtrlPagtoIdParcela(int ctrl_pagto_modulo, int ctrl_pagto_id_parcela, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.getPedidoHistPagtoByCtrlPagtoIdParcela()";
			string msg_erro_aux;
			string strSql = "";
			PedidoHistPagto histPagto = new PedidoHistPagto();
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				strSql = "SELECT " +
							"*" +
						" FROM t_FIN_PEDIDO_HIST_PAGTO" +
						" WHERE" +
							" (ctrl_pagto_modulo = " + ctrl_pagto_modulo.ToString() + ")" +
							" AND (ctrl_pagto_id_parcela = " + ctrl_pagto_id_parcela.ToString() + ")" +
						" ORDER BY" +
							" id DESC";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count == 0) return null;

				row = dtbResultado.Rows[0];

				return getPedidoHistPagtoById(BD.readToInt(row["id"]), out msg_erro);
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = "ctrl_pagto_modulo=" + ctrl_pagto_modulo.ToString() + ", ctrl_pagto_id_parcela=" + ctrl_pagto_id_parcela.ToString();
				svcLog.complemento_2 = strSql;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ getPedidoHistPagtoById ]
		public static PedidoHistPagto getPedidoHistPagtoById(int id, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.getPedidoHistPagtoById()";
			string msg_erro_aux;
			string strSql = "";
			PedidoHistPagto histPagto = new PedidoHistPagto();
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				strSql = "SELECT " +
							"*" +
						" FROM t_FIN_PEDIDO_HIST_PAGTO" +
						" WHERE" +
							" (id = " + id.ToString() + ")";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count == 0) return null;

				row = dtbResultado.Rows[0];

				histPagto.id = BD.readToInt(row["id"]);
				histPagto.pedido = BD.readToString(row["pedido"]);
				histPagto.status = BD.readToByte(row["status"]);
				histPagto.id_fluxo_caixa = BD.readToInt(row["id_fluxo_caixa"]);
				histPagto.ctrl_pagto_id_parcela = BD.readToInt(row["ctrl_pagto_id_parcela"]);
				histPagto.ctrl_pagto_modulo = BD.readToByte(row["ctrl_pagto_modulo"]);
				histPagto.dt_vencto = BD.readToDateTime(row["dt_vencto"]);
				histPagto.valor_total = BD.readToDecimal(row["valor_total"]);
				histPagto.valor_rateado = BD.readToDecimal(row["valor_rateado"]);
				histPagto.descricao = BD.readToString(row["descricao"]);
				histPagto.dt_credito = BD.readToDateTime(row["dt_credito"]);
				histPagto.dt_cadastro = BD.readToDateTime(row["dt_cadastro"]);
				histPagto.usuario_cadastro = BD.readToString(row["usuario_cadastro"]);
				histPagto.dt_ult_atualizacao = BD.readToDateTime(row["dt_ult_atualizacao"]);
				histPagto.usuario_ult_atualizacao = BD.readToString(row["usuario_ult_atualizacao"]);
				histPagto.vl_abatimento_concedido = BD.readToDecimal(row["vl_abatimento_concedido"]);
				histPagto.st_boleto_pago_cheque = BD.readToByte(row["st_boleto_pago_cheque"]);
				histPagto.dt_ocorrencia_banco_boleto_pago_cheque = BD.readToDateTime(row["dt_ocorrencia_banco_boleto_pago_cheque"]);
				histPagto.st_boleto_ocorrencia_17 = BD.readToByte(row["st_boleto_ocorrencia_17"]);
				histPagto.dt_ocorrencia_banco_boleto_ocorrencia_17 = BD.readToDateTime(row["dt_ocorrencia_banco_boleto_ocorrencia_17"]);
				histPagto.st_boleto_ocorrencia_15 = BD.readToByte(row["st_boleto_ocorrencia_15"]);
				histPagto.dt_ocorrencia_banco_boleto_ocorrencia_15 = BD.readToDateTime(row["dt_ocorrencia_banco_boleto_ocorrencia_15"]);
				histPagto.st_boleto_ocorrencia_23 = BD.readToByte(row["st_boleto_ocorrencia_23"]);
				histPagto.dt_ocorrencia_banco_boleto_ocorrencia_23 = BD.readToDateTime(row["dt_ocorrencia_banco_boleto_ocorrencia_23"]);
				histPagto.st_boleto_ocorrencia_34 = BD.readToByte(row["st_boleto_ocorrencia_34"]);
				histPagto.dt_ocorrencia_banco_boleto_ocorrencia_34 = BD.readToDateTime(row["dt_ocorrencia_banco_boleto_ocorrencia_34"]);
				histPagto.st_boleto_baixado = BD.readToByte(row["st_boleto_baixado"]);
				histPagto.dt_ocorrencia_banco_boleto_baixado = BD.readToDateTime(row["dt_ocorrencia_banco_boleto_baixado"]);
				histPagto.dt_operacao = BD.readToDateTime(row["dt_operacao"]);
				histPagto.valor_pago = BD.readToDecimal(row["valor_pago"]);

				return histPagto;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = "id=" + id.ToString();
				svcLog.complemento_2 = strSql;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ getPedidoPagamentoByPedido ]
		public static List<PedidoPagamento> getPedidoPagamentoByPedido(string pedido, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.getPedidoPagamentoByPedido()";
			string msg_erro_aux;
			string strSql = "";
			PedidoPagamento pagto = new PedidoPagamento();
			List<PedidoPagamento> listaPagto = new List<PedidoPagamento>();
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				if ((pedido ?? "").Length == 0)
				{
					msg_erro = "Não foi informado o número do pedido!";
					return null;
				}

				strSql = "SELECT " +
							"id" +
						" FROM t_PEDIDO_PAGAMENTO" +
						" WHERE" +
							" (pedido = '" + pedido + "')" +
						" ORDER BY" +
							" id";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count == 0) return null;

				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					row = dtbResultado.Rows[i];
					pagto = getPedidoPagamentoById(BD.readToString(row["id"]), out msg_erro_aux);
					if (pagto != null) listaPagto.Add(pagto);
				}

				return listaPagto;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = "pedido=" + pedido;
				svcLog.complemento_2 = strSql;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ getPedidoPagamentoById ]
		public static PedidoPagamento getPedidoPagamentoById(string id, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "PedidoDAO.getPedidoPagamentoById()";
			string msg_erro_aux;
			string strSql = "";
			PedidoPagamento pagto = new PedidoPagamento();
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				if ((id??"").Length==0)
				{
					msg_erro = "Não foi informado o identificador do registro do pagamento!";
					return null;
				}

				strSql = "SELECT " +
							"*" +
						" FROM t_PEDIDO_PAGAMENTO" +
						" WHERE" +
							" (id = '" + id + "')";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count == 0) return null;

				row = dtbResultado.Rows[0];

				pagto.id = BD.readToString(row["id"]);
				pagto.pedido = BD.readToString(row["pedido"]);
				pagto.data = BD.readToDateTime(row["data"]);
				pagto.hora = BD.readToString(row["hora"]);
				pagto.valor = BD.readToDecimal(row["valor"]);
				pagto.tipo_pagto = BD.readToString(row["tipo_pagto"]);
				pagto.usuario = BD.readToString(row["usuario"]);
				pagto.id_pedido_pagto_cielo = BD.readToInt(row["id_pedido_pagto_cielo"]);
				pagto.id_pedido_pagto_braspag = BD.readToInt(row["id_pedido_pagto_braspag"]);
				pagto.id_pagto_gw_pag_payment = BD.readToInt(row["id_pagto_gw_pag_payment"]);
				pagto.id_braspag_webhook_complementar = BD.readToInt(row["id_braspag_webhook_complementar"]);
				
				return pagto;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = "t_PEDIDO_PAGAMENTO.id=" + id;
				svcLog.complemento_2 = strSql;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion
	}
}
