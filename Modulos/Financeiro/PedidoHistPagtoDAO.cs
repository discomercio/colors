#region [ using ]
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
#endregion

namespace Financeiro
{
	class PedidoHistPagtoDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02;
		private static SqlCommand cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02;
		private static SqlCommand cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06;
		private static SqlCommand cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09;
		private static SqlCommand cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10;
		private static SqlCommand cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia12;
		private static SqlCommand cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia13;
		private static SqlCommand cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia14;
		private static SqlCommand cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15;
		private static SqlCommand cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16;
		private static SqlCommand cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17;
		private static SqlCommand cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22;
		private static SqlCommand cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23;
		private static SqlCommand cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34;
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
		static PedidoHistPagtoDAO()
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

			#region [ cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02 ]
			strSql = "INSERT INTO t_FIN_PEDIDO_HIST_PAGTO (" +
						"id, " +
						"pedido, " +
						"status, " +
						"id_fluxo_caixa, " +
						"ctrl_pagto_id_parcela, " +
						"ctrl_pagto_modulo, " +
						"dt_vencto, " +
						"valor_total, " +
						"valor_rateado, " +
						"descricao, " +
						"usuario_cadastro, " +
						"usuario_ult_atualizacao" +
					") VALUES (" +
						"@id, " +
						"@pedido, " +
						"@status, " +
						"@id_fluxo_caixa, " +
						"@ctrl_pagto_id_parcela, " +
						"@ctrl_pagto_modulo, " +
						"@dt_vencto, " +
						"@valor_total, " +
						"@valor_rateado, " +
						"@descricao, " +
						"@usuario_cadastro, " +
						"@usuario_ult_atualizacao" +
					")";
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02 = BD.criaSqlCommand();
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.CommandText = strSql;
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters.Add("@id", SqlDbType.Int);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters.Add("@status", SqlDbType.TinyInt);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters.Add("@id_fluxo_caixa", SqlDbType.Int);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters.Add("@dt_vencto", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters.Add("@valor_total", SqlDbType.Money);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters.Add("@valor_rateado", SqlDbType.Money);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters.Add("@descricao", SqlDbType.VarChar, 60);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters.Add("@usuario_cadastro", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Prepare();
			#endregion

			#region [ cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02 ]
			strSql = "UPDATE t_FIN_PEDIDO_HIST_PAGTO SET " +
						"status = @status, " +
						"dt_vencto = @dt_vencto, " +
						"valor_total = @valor_total, " +
						"valor_rateado = @valor_rateado, " +
						"descricao = @descricao, " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE (id = @id)";
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02 = BD.criaSqlCommand();
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.CommandText = strSql;
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.Parameters.Add("@id", SqlDbType.Int);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.Parameters.Add("@status", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.Parameters.Add("@dt_vencto", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.Parameters.Add("@valor_total", SqlDbType.Money);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.Parameters.Add("@valor_rateado", SqlDbType.Money);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.Parameters.Add("@descricao", SqlDbType.VarChar, 60);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.Prepare();
			#endregion

			#region [ cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06 ]
			strSql = "UPDATE t_FIN_PEDIDO_HIST_PAGTO SET " +
						"status = @status, " +
						"dt_credito = @dt_credito, " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE" +
						" (ctrl_pagto_modulo = @ctrl_pagto_modulo)" +
						" AND (ctrl_pagto_id_parcela = @ctrl_pagto_id_parcela)";
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06 = BD.criaSqlCommand();
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06.CommandText = strSql;
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06.Parameters.Add("@status", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06.Parameters.Add("@dt_credito", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06.Prepare();
			#endregion

			#region [ cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09 ]
			strSql = "UPDATE t_FIN_PEDIDO_HIST_PAGTO SET " +
						"status = " + Global.Cte.FIN.ST_T_FIN_PEDIDO_HIST_PAGTO.CANCELADO.ToString() + ", " +
						"st_boleto_baixado = @st_boleto_baixado, " +
						"dt_ocorrencia_banco_boleto_baixado = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_baixado") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE" +
						" (ctrl_pagto_modulo = @ctrl_pagto_modulo)" +
						" AND (ctrl_pagto_id_parcela = @ctrl_pagto_id_parcela)";
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09 = BD.criaSqlCommand();
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09.CommandText = strSql;
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09.Parameters.Add("@st_boleto_baixado", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09.Parameters.Add("@dt_ocorrencia_banco_boleto_baixado", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09.Prepare();
			#endregion

			#region [ cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10 ]
			strSql = "UPDATE t_FIN_PEDIDO_HIST_PAGTO SET " +
						"status = " + Global.Cte.FIN.ST_T_FIN_PEDIDO_HIST_PAGTO.CANCELADO.ToString() + ", " +
						"st_boleto_baixado = @st_boleto_baixado, " +
						"dt_ocorrencia_banco_boleto_baixado = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_baixado") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE" +
						" (ctrl_pagto_modulo = @ctrl_pagto_modulo)" +
						" AND (ctrl_pagto_id_parcela = @ctrl_pagto_id_parcela)";
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10 = BD.criaSqlCommand();
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10.CommandText = strSql;
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10.Parameters.Add("@st_boleto_baixado", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10.Parameters.Add("@dt_ocorrencia_banco_boleto_baixado", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10.Prepare();
			#endregion

			#region [ cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia12 ]
			strSql = "UPDATE t_FIN_PEDIDO_HIST_PAGTO SET " +
						"vl_abatimento_concedido = @vl_abatimento_concedido, " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE" +
						" (ctrl_pagto_modulo = @ctrl_pagto_modulo)" +
						" AND (ctrl_pagto_id_parcela = @ctrl_pagto_id_parcela)";
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia12 = BD.criaSqlCommand();
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia12.CommandText = strSql;
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia12.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia12.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia12.Parameters.Add("@vl_abatimento_concedido", SqlDbType.Money);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia12.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia12.Prepare();
			#endregion

			#region [ cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia13 ]
			strSql = "UPDATE t_FIN_PEDIDO_HIST_PAGTO SET " +
						"vl_abatimento_concedido = @vl_abatimento_concedido, " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE" +
						" (ctrl_pagto_modulo = @ctrl_pagto_modulo)" +
						" AND (ctrl_pagto_id_parcela = @ctrl_pagto_id_parcela)";
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia13 = BD.criaSqlCommand();
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia13.CommandText = strSql;
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia13.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia13.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia13.Parameters.Add("@vl_abatimento_concedido", SqlDbType.Money);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia13.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia13.Prepare();
			#endregion

			#region [ cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia14 ]
			strSql = "UPDATE t_FIN_PEDIDO_HIST_PAGTO SET " +
						"dt_vencto = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_vencto") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE" +
						" (ctrl_pagto_modulo = @ctrl_pagto_modulo)" +
						" AND (ctrl_pagto_id_parcela = @ctrl_pagto_id_parcela)";
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia14 = BD.criaSqlCommand();
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia14.CommandText = strSql;
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia14.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia14.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia14.Parameters.Add("@dt_vencto", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia14.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia14.Prepare();
			#endregion

			#region [ cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15 ]
			strSql = "UPDATE t_FIN_PEDIDO_HIST_PAGTO SET " +
						"status = @status, " +
						"dt_credito = @dt_credito, " +
						"st_boleto_ocorrencia_15 = @st_boleto_ocorrencia_15, " +
						"dt_ocorrencia_banco_boleto_ocorrencia_15 = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_ocorrencia_15") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE" +
						" (ctrl_pagto_modulo = @ctrl_pagto_modulo)" +
						" AND (ctrl_pagto_id_parcela = @ctrl_pagto_id_parcela)";
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15 = BD.criaSqlCommand();
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.CommandText = strSql;
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@status", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@dt_credito", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@st_boleto_ocorrencia_15", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@dt_ocorrencia_banco_boleto_ocorrencia_15", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.Prepare();
			#endregion

			#region [ cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16 ]
			strSql = "UPDATE t_FIN_PEDIDO_HIST_PAGTO SET " +
						"st_boleto_pago_cheque = @st_boleto_pago_cheque, " +
						"dt_ocorrencia_banco_boleto_pago_cheque = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_pago_cheque") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE" +
						" (ctrl_pagto_modulo = @ctrl_pagto_modulo)" +
						" AND (ctrl_pagto_id_parcela = @ctrl_pagto_id_parcela)";
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16 = BD.criaSqlCommand();
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16.CommandText = strSql;
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16.Parameters.Add("@st_boleto_pago_cheque", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16.Parameters.Add("@dt_ocorrencia_banco_boleto_pago_cheque", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16.Prepare();
			#endregion

			#region [ cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17 ]
			strSql = "INSERT INTO t_FIN_PEDIDO_HIST_PAGTO (" +
						"id, " +
						"pedido, " +
						"status, " +
						"id_fluxo_caixa, " +
						"ctrl_pagto_id_parcela, " +
						"ctrl_pagto_modulo, " +
						"dt_vencto, " +
						"valor_total, " +
						"valor_rateado, " +
						"descricao, " +
						"usuario_cadastro, " +
						"usuario_ult_atualizacao, " +
						"st_boleto_ocorrencia_17, " +
						"dt_ocorrencia_banco_boleto_ocorrencia_17" +
					") VALUES (" +
						"@id, " +
						"@pedido, " +
						"@status, " +
						"@id_fluxo_caixa, " +
						"@ctrl_pagto_id_parcela, " +
						"@ctrl_pagto_modulo, " +
						"@dt_vencto, " +
						"@valor_total, " +
						"@valor_rateado, " +
						"@descricao, " +
						"@usuario_cadastro, " +
						"@usuario_ult_atualizacao, " +
						"@st_boleto_ocorrencia_17, " +
						"@dt_ocorrencia_banco_boleto_ocorrencia_17" +
					")";
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17 = BD.criaSqlCommand();
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.CommandText = strSql;
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters.Add("@id", SqlDbType.Int);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters.Add("@status", SqlDbType.TinyInt);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters.Add("@id_fluxo_caixa", SqlDbType.Int);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters.Add("@dt_vencto", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters.Add("@valor_total", SqlDbType.Money);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters.Add("@valor_rateado", SqlDbType.Money);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters.Add("@descricao", SqlDbType.VarChar, 60);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters.Add("@usuario_cadastro", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters.Add("@st_boleto_ocorrencia_17", SqlDbType.TinyInt);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters.Add("@dt_ocorrencia_banco_boleto_ocorrencia_17", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Prepare();
			#endregion

			#region [ cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22 ]
			strSql = "UPDATE t_FIN_PEDIDO_HIST_PAGTO SET " +
						"st_boleto_pago_cheque = @st_boleto_pago_cheque, " +
						"dt_ocorrencia_banco_boleto_pago_cheque = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_pago_cheque") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE" +
						" (ctrl_pagto_modulo = @ctrl_pagto_modulo)" +
						" AND (ctrl_pagto_id_parcela = @ctrl_pagto_id_parcela)";
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22 = BD.criaSqlCommand();
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22.CommandText = strSql;
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22.Parameters.Add("@st_boleto_pago_cheque", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22.Parameters.Add("@dt_ocorrencia_banco_boleto_pago_cheque", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22.Prepare();
			#endregion

			#region [ cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23 ]
			strSql = "UPDATE t_FIN_PEDIDO_HIST_PAGTO SET " +
						"st_boleto_ocorrencia_23 = @st_boleto_ocorrencia_23, " +
						"dt_ocorrencia_banco_boleto_ocorrencia_23 = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_ocorrencia_23") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE" +
						" (ctrl_pagto_modulo = @ctrl_pagto_modulo)" +
						" AND (ctrl_pagto_id_parcela = @ctrl_pagto_id_parcela)";
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23 = BD.criaSqlCommand();
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23.CommandText = strSql;
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23.Parameters.Add("@st_boleto_ocorrencia_23", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23.Parameters.Add("@dt_ocorrencia_banco_boleto_ocorrencia_23", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23.Prepare();
			#endregion

			#region [ cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34 ]
			strSql = "UPDATE t_FIN_PEDIDO_HIST_PAGTO SET " +
						"st_boleto_ocorrencia_34 = @st_boleto_ocorrencia_34, " +
						"dt_ocorrencia_banco_boleto_ocorrencia_34 = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@dt_ocorrencia_banco_boleto_ocorrencia_34") + ", " +
						"dt_ult_atualizacao = " + Global.sqlMontaGetdateSomenteData() + ", " +
						"usuario_ult_atualizacao = @usuario_ult_atualizacao " +
					 "WHERE" +
						" (ctrl_pagto_modulo = @ctrl_pagto_modulo)" +
						" AND (ctrl_pagto_id_parcela = @ctrl_pagto_id_parcela)";
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34 = BD.criaSqlCommand();
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34.CommandText = strSql;
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34.Parameters.Add("@ctrl_pagto_modulo", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34.Parameters.Add("@ctrl_pagto_id_parcela", SqlDbType.Int);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34.Parameters.Add("@st_boleto_ocorrencia_34", SqlDbType.TinyInt);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34.Parameters.Add("@dt_ocorrencia_banco_boleto_ocorrencia_34", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34.Parameters.Add("@usuario_ult_atualizacao", SqlDbType.VarChar, 10);
			cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34.Prepare();
			#endregion
		}
		#endregion

		#region [ obtemRegistroPedidoHistPagtoByCtrlPagtoIdParcela ]
		/// <summary>
		/// Retorna o DataTable contendo o registro do histórico de pagamentos do pedido,
		/// sendo que a pesquisa é feita através do campo ctrl_pagto_id_parcela + ctrl_pagto_modulo, 
		/// ou seja, está sendo localizado o registro associado a um dos boletos, cheques 
		/// ou parcelas do Visa.
		/// </summary>
		/// <param name="pedido">
		/// Número do pedido.
		/// </param>
		/// <param name="ctrlPagtoIdParcela">
		/// Nº identificação do registro no módulo de controle que gerou o lançamento no fluxo de caixa (módulos
		/// de controle: boleto, cheque, Visa)
		/// </param>
		/// <param name="ctrPagtoModulo">
		/// Código de identificação do módulo de controle
		///		1 = Boleto
		///		2 = Cheque
		///		3 = Visa
		/// </param>
		/// <returns>
		/// Retorna o DataTable contendo o registro especificado
		/// </returns>
		private static DsDataSource.DtbFinPedidoHistPagtoDataTable obtemRegistroPedidoHistPagtoByCtrlPagtoIdParcela(String pedido, int ctrlPagtoIdParcela, byte ctrPagtoModulo)
		{
			String strSql;
			String strWhere = "";
			SqlCommand cmCommand;
			SqlDataAdapter daDataAdapter;
			DsDataSource.DtbFinPedidoHistPagtoDataTable dtbFinPedidoHistPagto = new DsDataSource.DtbFinPedidoHistPagtoDataTable();

			cmCommand = BD.criaSqlCommand();
			daDataAdapter = BD.criaSqlDataAdapter();

			strWhere = " (pedido = '" + pedido + "')" +
					   " AND (ctrl_pagto_id_parcela = " + ctrlPagtoIdParcela.ToString() + ")" +
					   " AND (ctrl_pagto_modulo = " + ctrPagtoModulo.ToString() + ")";
			if (strWhere.Length > 0) strWhere = " WHERE" + strWhere;

			strSql = "SELECT " +
						"*" +
					" FROM t_FIN_PEDIDO_HIST_PAGTO" +
					strWhere;
			cmCommand.CommandText = strSql;
			daDataAdapter.SelectCommand = cmCommand;
			daDataAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			daDataAdapter.Fill(dtbFinPedidoHistPagto);

			return dtbFinPedidoHistPagto;
		}
		#endregion

		#region [ getPedidoHistPagtoByCtrlPagtoIdParcela ]
		/// <summary>
		/// Retorna um objeto representando uma parcela de pagamento no histórico de pagamentos 
		/// do pedido contendo os dados lidos do BD, sendo que a pesquisa é feita através do campo 
		/// ctrl_pagto_id_parcela + ctrl_pagto_modulo, ou seja, está sendo localizado o registro 
		/// associado a um dos boletos, cheques ou parcelas do Visa.
		/// </summary>
		/// <param name="pedido">
		/// Número do pedido.
		/// </param>
		/// <param name="ctrlPagtoIdParcela">
		/// Nº identificação do registro no módulo de controle que gerou o lançamento no fluxo de caixa (módulos
		/// de controle: boleto, cheque, Visa)
		/// </param>
		/// <param name="ctrPagtoModulo">
		/// Código de identificação do módulo de controle
		///		1 = Boleto
		///		2 = Cheque
		///		3 = Visa
		/// </param>
		/// <returns>
		/// Retorna um objeto PedidoHistPagto com os dados da parcela de pagamento no histórico 
		/// de pagamentos do pedido.
		/// </returns>
		public static PedidoHistPagto getPedidoHistPagtoByCtrlPagtoIdParcela(String pedido, int ctrlPagtoIdParcela, byte ctrPagtoModulo)
		{
			PedidoHistPagto pedidoHistPagto = new PedidoHistPagto();
			DsDataSource.DtbFinPedidoHistPagtoDataTable dtbFinPedidoHistPagto;
			DsDataSource.DtbFinPedidoHistPagtoRow rowFinPedidoHistPagto;

			if (ctrlPagtoIdParcela == 0) throw new FinanceiroException("O identificador do registro não foi informado!!");
			if (ctrPagtoModulo == 0) throw new FinanceiroException("Não foi informado o código do módulo de controle!!");

			dtbFinPedidoHistPagto = obtemRegistroPedidoHistPagtoByCtrlPagtoIdParcela(pedido, ctrlPagtoIdParcela, ctrPagtoModulo);

			if (dtbFinPedidoHistPagto.Rows.Count == 0) return null;

			rowFinPedidoHistPagto = (DsDataSource.DtbFinPedidoHistPagtoRow)dtbFinPedidoHistPagto.Rows[0];

			pedidoHistPagto.id = BD.readToInt(rowFinPedidoHistPagto.id);
			pedidoHistPagto.pedido = BD.readToString(rowFinPedidoHistPagto.pedido);
			pedidoHistPagto.status = BD.readToByte(rowFinPedidoHistPagto.status);
			pedidoHistPagto.id_fluxo_caixa = BD.readToInt(rowFinPedidoHistPagto.id_fluxo_caixa);
			pedidoHistPagto.ctrl_pagto_id_parcela = BD.readToInt(rowFinPedidoHistPagto.ctrl_pagto_id_parcela);
			pedidoHistPagto.ctrl_pagto_modulo = BD.readToByte(rowFinPedidoHistPagto.ctrl_pagto_modulo);
			pedidoHistPagto.dt_vencto = BD.readToDateTime(rowFinPedidoHistPagto.dt_vencto);
			pedidoHistPagto.valor_total = BD.readToDecimal(rowFinPedidoHistPagto.valor_total);
			pedidoHistPagto.valor_rateado = BD.readToDecimal(rowFinPedidoHistPagto.valor_rateado);
			pedidoHistPagto.descricao = BD.readToString(rowFinPedidoHistPagto.descricao);
			pedidoHistPagto.dt_credito = (!rowFinPedidoHistPagto.Isdt_creditoNull() ? (DateTime)rowFinPedidoHistPagto.dt_credito : DateTime.MinValue);
			pedidoHistPagto.dt_cadastro = BD.readToDateTime(rowFinPedidoHistPagto.dt_cadastro);
			pedidoHistPagto.usuario_cadastro = BD.readToString(rowFinPedidoHistPagto.usuario_cadastro);
			pedidoHistPagto.dt_ult_atualizacao = BD.readToDateTime(rowFinPedidoHistPagto.dt_ult_atualizacao);
			pedidoHistPagto.usuario_ult_atualizacao = BD.readToString(rowFinPedidoHistPagto.usuario_ult_atualizacao);
			pedidoHistPagto.vl_abatimento_concedido = BD.readToDecimal(rowFinPedidoHistPagto.vl_abatimento_concedido);
			pedidoHistPagto.st_boleto_pago_cheque = rowFinPedidoHistPagto.st_boleto_pago_cheque;
			pedidoHistPagto.dt_ocorrencia_banco_boleto_pago_cheque = (!rowFinPedidoHistPagto.Isdt_ocorrencia_banco_boleto_pago_chequeNull() ? (DateTime)rowFinPedidoHistPagto.dt_ocorrencia_banco_boleto_pago_cheque : DateTime.MinValue);
			pedidoHistPagto.st_boleto_ocorrencia_17 = rowFinPedidoHistPagto.st_boleto_ocorrencia_17;
			pedidoHistPagto.dt_ocorrencia_banco_boleto_ocorrencia_17 = (!rowFinPedidoHistPagto.Isdt_ocorrencia_banco_boleto_ocorrencia_17Null() ? (DateTime)rowFinPedidoHistPagto.dt_ocorrencia_banco_boleto_ocorrencia_17 : DateTime.MinValue);
			pedidoHistPagto.st_boleto_ocorrencia_15 = rowFinPedidoHistPagto.st_boleto_ocorrencia_15;
			pedidoHistPagto.dt_ocorrencia_banco_boleto_ocorrencia_15 = (!rowFinPedidoHistPagto.Isdt_ocorrencia_banco_boleto_ocorrencia_15Null() ? (DateTime)rowFinPedidoHistPagto.dt_ocorrencia_banco_boleto_ocorrencia_15 : DateTime.MinValue);
			pedidoHistPagto.st_boleto_ocorrencia_23 = rowFinPedidoHistPagto.st_boleto_ocorrencia_23;
			pedidoHistPagto.dt_ocorrencia_banco_boleto_ocorrencia_23 = (!rowFinPedidoHistPagto.Isdt_ocorrencia_banco_boleto_ocorrencia_23Null() ? (DateTime)rowFinPedidoHistPagto.dt_ocorrencia_banco_boleto_ocorrencia_23 : DateTime.MinValue);
			pedidoHistPagto.st_boleto_ocorrencia_34 = rowFinPedidoHistPagto.st_boleto_ocorrencia_34;
			pedidoHistPagto.dt_ocorrencia_banco_boleto_ocorrencia_34 = (!rowFinPedidoHistPagto.Isdt_ocorrencia_banco_boleto_ocorrencia_34Null() ? (DateTime)rowFinPedidoHistPagto.dt_ocorrencia_banco_boleto_ocorrencia_34 : DateTime.MinValue);
			pedidoHistPagto.st_boleto_baixado = rowFinPedidoHistPagto.st_boleto_baixado;
			pedidoHistPagto.dt_ocorrencia_banco_boleto_baixado = (!rowFinPedidoHistPagto.Isdt_ocorrencia_banco_boleto_baixadoNull() ? (DateTime)rowFinPedidoHistPagto.dt_ocorrencia_banco_boleto_baixado : DateTime.MinValue);

			return pedidoHistPagto;
		}
		#endregion

		#region [ inserePagtoDevidoBoletoOcorrencia02 ]
		public static bool inserePagtoDevidoBoletoOcorrencia02(String usuario,
														  int idLancamentoFluxoCaixa,
														  int idBoletoItem,
														  DateTime dataVencto,
														  decimal valorTitulo,
														  String descricao,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Inserção no histórico de pagamento do pedido devido a boleto com ocorrência 02 (entrada confirmada)";
			bool blnExisteRegistro;
			bool blnGerouNsu;
			int intNsuPedidoHistPagto = 0;
			int intRetorno;
			int intCounter;
			decimal valorRateado;
			String strPedido;
			PedidoHistPagto pedidoHistPagtoJaGravado;
			DsDataSource.DtbFinBoletoItemRateioDataTable dtbFinBoletoItemRateio;
			DsDataSource.DtbFinBoletoItemRateioRow rowFinBoletoItemRateio;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Obtém os pedidos envolvidos no rateio ]
				dtbFinBoletoItemRateio = BoletoDAO.obtemBoletoItemRateio(idBoletoItem);
				#endregion

				for (intCounter = 0; intCounter < dtbFinBoletoItemRateio.Rows.Count; intCounter++)
				{
					rowFinBoletoItemRateio = (DsDataSource.DtbFinBoletoItemRateioRow)dtbFinBoletoItemRateio.Rows[intCounter];
					strPedido = rowFinBoletoItemRateio.pedido;
					if (Global.isNumeroPedido(strPedido))
					{
						valorRateado = rowFinBoletoItemRateio.valor;

						#region [ Pesquisa BD p/ verificar se já existe registro associado a este boleto ]
						pedidoHistPagtoJaGravado = getPedidoHistPagtoByCtrlPagtoIdParcela(strPedido, idBoletoItem, Global.Cte.FIN.CtrlPagtoModulo.BOLETO);
						blnExisteRegistro = true;
						if (pedidoHistPagtoJaGravado == null)
							blnExisteRegistro = false;
						else if (pedidoHistPagtoJaGravado.id == 0)
							blnExisteRegistro = false;
						#endregion

						if (blnExisteRegistro)
						{
							#region [ Já existe um registro associado a este boleto, então atualiza-o ]

							#region [ Preenche o valor dos parâmetros ]
							cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.Parameters["@id"].Value = pedidoHistPagtoJaGravado.id;
							cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.Parameters["@status"].Value = Global.Cte.FIN.ST_T_FIN_PEDIDO_HIST_PAGTO.PREVISAO;
							cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.Parameters["@dt_vencto"].Value = Global.formataDataYyyyMmDdComSeparador(dataVencto);
							cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.Parameters["@valor_total"].Value = valorTitulo;
							cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.Parameters["@valor_rateado"].Value = valorRateado;
							cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.Parameters["@descricao"].Value = descricao;
							cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02.Parameters["@usuario_ult_atualizacao"].Value = usuario;
							#endregion

							#region [ Tenta atualizar o registro ]
							try
							{
								intRetorno = BD.executaNonQuery(ref cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia02);
							}
							catch (Exception ex)
							{
								intRetorno = 0;
								Global.gravaLogAtividade(strOperacao + "\nAlteração sobre registro já existente!!\n" + ex.ToString());
							}
							#endregion

							#region [ Processamento para sucesso ou falha desta tentativa de atualização ]
							if (intRetorno != 1)
							{
								strMsgErro = "Falha ao tentar atualizar o registro associado do histórico de pagamentos do pedido devido a boleto com ocorrência 02 (entrada confirmada)!!";
								return false;
							}
							#endregion

							#endregion
						}
						else
						{
							#region [ Não há nenhum registro associado a este boleto, então cria um novo ]

							#region [ Gera NSU ]
							blnGerouNsu = BD.geraNsu(Global.Cte.FIN.NSU.T_FIN_PEDIDO_HIST_PAGTO, ref intNsuPedidoHistPagto, ref strMsgErro);
							if (!blnGerouNsu)
							{
								strMsgErro = "Falha ao gerar o NSU para o registro do histórico de pagamentos do pedido!!\n" + strMsgErro;
								return false;
							}
							#endregion

							#region [ Preenche o valor dos parâmetros ]
							cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters["@id"].Value = intNsuPedidoHistPagto;
							cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters["@pedido"].Value = strPedido;
							cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters["@status"].Value = Global.Cte.FIN.ST_T_FIN_PEDIDO_HIST_PAGTO.PREVISAO;
							cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters["@id_fluxo_caixa"].Value = idLancamentoFluxoCaixa;
							cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters["@ctrl_pagto_id_parcela"].Value = idBoletoItem;
							cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters["@ctrl_pagto_modulo"].Value = Global.Cte.FIN.CtrlPagtoModulo.BOLETO;
							cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters["@dt_vencto"].Value = Global.formataDataYyyyMmDdComSeparador(dataVencto);
							cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters["@valor_total"].Value = valorTitulo;
							cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters["@valor_rateado"].Value = valorRateado;
							cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters["@descricao"].Value = descricao;
							cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters["@usuario_cadastro"].Value = usuario;
							cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02.Parameters["@usuario_ult_atualizacao"].Value = usuario;
							#endregion

							#region [ Tenta inserir o registro ]
							try
							{
								intRetorno = BD.executaNonQuery(ref cmPedidoHistPagtoInsertDevidoBoletoOcorrencia02);
							}
							catch (Exception ex)
							{
								intRetorno = 0;
								Global.gravaLogAtividade(strOperacao + "\n" + ex.ToString());
							}
							#endregion

							#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
							if (intRetorno != 1)
							{
								strMsgErro = "Falha ao tentar gerar o registro automático no histórico de pagamentos do pedido devido a boleto com ocorrência 02 (entrada confirmada)!!";
								return false;
							}
							#endregion

							#endregion
						}
					}
				}

				return true;
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaPagtoDevidoBoletoOcorrencia06 ]
		public static bool atualizaPagtoDevidoBoletoOcorrencia06(String usuario,
														  int idBoletoItem,
														  DateTime dtCredito,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o histórico de pagamentos do pedido devido a ocorrência 06 (liquidação normal)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06.Parameters["@ctrl_pagto_modulo"].Value = Global.Cte.FIN.CtrlPagtoModulo.BOLETO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06.Parameters["@ctrl_pagto_id_parcela"].Value = idBoletoItem;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06.Parameters["@status"].Value = Global.Cte.FIN.ST_T_FIN_PEDIDO_HIST_PAGTO.QUITADO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06.Parameters["@dt_credito"].Value = Global.formataDataYyyyMmDdComSeparador(dtCredito);
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia06);
					blnSucesso = true;
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do histórico de pagamentos do pedido durante tratamento da ocorrência 06 (liquidação normal)!!";
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

		#region [ atualizaPagtoDevidoBoletoOcorrencia09 ]
		public static bool atualizaPagtoDevidoBoletoOcorrencia09(String usuario,
														  int idBoletoItem,
														  DateTime dataOcorrenciaBanco,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o histórico de pagamentos do pedido devido a ocorrência 09 (baixado automaticamente via arquivo)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09.Parameters["@ctrl_pagto_modulo"].Value = Global.Cte.FIN.CtrlPagtoModulo.BOLETO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09.Parameters["@ctrl_pagto_id_parcela"].Value = idBoletoItem;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09.Parameters["@st_boleto_baixado"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09.Parameters["@dt_ocorrencia_banco_boleto_baixado"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia09);
					blnSucesso = true;
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do histórico de pagamentos do pedido durante tratamento da ocorrência 09 (baixado automaticamente via arquivo)!!";
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

		#region [ atualizaPagtoDevidoBoletoOcorrencia10 ]
		public static bool atualizaPagtoDevidoBoletoOcorrencia10(String usuario,
														  int idBoletoItem,
														  DateTime dataOcorrenciaBanco,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o histórico de pagamentos do pedido devido a ocorrência 10 (baixado conforme instruções da agência)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10.Parameters["@ctrl_pagto_modulo"].Value = Global.Cte.FIN.CtrlPagtoModulo.BOLETO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10.Parameters["@ctrl_pagto_id_parcela"].Value = idBoletoItem;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10.Parameters["@st_boleto_baixado"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10.Parameters["@dt_ocorrencia_banco_boleto_baixado"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia10);
					blnSucesso = true;
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do histórico de pagamentos do pedido durante tratamento da ocorrência 10 (baixado conforme instruções da agência)!!";
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

		#region [ atualizaPagtoDevidoBoletoOcorrencia12 ]
		public static bool atualizaPagtoDevidoBoletoOcorrencia12(String usuario,
														  int idBoletoItem,
														  decimal vlAbatimentoConcedido,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o histórico de pagamentos do pedido devido a ocorrência 12 (abatimento concedido)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia12.Parameters["@ctrl_pagto_modulo"].Value = Global.Cte.FIN.CtrlPagtoModulo.BOLETO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia12.Parameters["@ctrl_pagto_id_parcela"].Value = idBoletoItem;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia12.Parameters["@vl_abatimento_concedido"].Value = vlAbatimentoConcedido;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia12.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia12);
					blnSucesso = true;
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do histórico de pagamentos do pedido durante tratamento da ocorrência 12 (abatimento concedido)!!";
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

		#region [ atualizaPagtoDevidoBoletoOcorrencia13 ]
		public static bool atualizaPagtoDevidoBoletoOcorrencia13(String usuario,
														  int idBoletoItem,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o histórico de pagamentos do pedido devido a ocorrência 13 (abatimento cancelado)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia13.Parameters["@ctrl_pagto_modulo"].Value = Global.Cte.FIN.CtrlPagtoModulo.BOLETO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia13.Parameters["@ctrl_pagto_id_parcela"].Value = idBoletoItem;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia13.Parameters["@vl_abatimento_concedido"].Value = 0m;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia13.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia13);
					blnSucesso = true;
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do histórico de pagamentos do pedido durante tratamento da ocorrência 13 (abatimento cancelado)!!";
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

		#region [ atualizaPagtoDevidoBoletoOcorrencia14 ]
		public static bool atualizaPagtoDevidoBoletoOcorrencia14(String usuario,
														  int idBoletoItem,
														  DateTime dtNovoVencto,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o histórico de pagamentos do pedido devido a ocorrência 14 (vencimento alterado)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia14.Parameters["@ctrl_pagto_modulo"].Value = Global.Cte.FIN.CtrlPagtoModulo.BOLETO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia14.Parameters["@ctrl_pagto_id_parcela"].Value = idBoletoItem;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia14.Parameters["@dt_vencto"].Value = Global.formataDataYyyyMmDdComSeparador(dtNovoVencto);
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia14.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia14);
					blnSucesso = true;
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do histórico de pagamentos do pedido durante tratamento da ocorrência 14 (vencimento alterado)!!";
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

		#region [ atualizaPagtoDevidoBoletoOcorrencia15 ]
		public static bool atualizaPagtoDevidoBoletoOcorrencia15(String usuario,
														  int idBoletoItem,
														  DateTime dtCredito,
														  DateTime dtOcorrenciaBanco,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o histórico de pagamentos do pedido devido a ocorrência 15 (liquidação em cartório)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.Parameters["@ctrl_pagto_modulo"].Value = Global.Cte.FIN.CtrlPagtoModulo.BOLETO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.Parameters["@ctrl_pagto_id_parcela"].Value = idBoletoItem;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.Parameters["@status"].Value = Global.Cte.FIN.ST_T_FIN_PEDIDO_HIST_PAGTO.QUITADO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.Parameters["@dt_credito"].Value = Global.formataDataYyyyMmDdComSeparador(dtCredito);
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.Parameters["@st_boleto_ocorrencia_15"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.Parameters["@dt_ocorrencia_banco_boleto_ocorrencia_15"].Value = Global.formataDataYyyyMmDdComSeparador(dtOcorrenciaBanco);
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia15);
					blnSucesso = true;
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do histórico de pagamentos do pedido durante tratamento da ocorrência 15 (liquidação em cartório)!!";
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

		#region [ atualizaPagtoDevidoBoletoOcorrencia16 ]
		public static bool atualizaPagtoDevidoBoletoOcorrencia16(String usuario,
														  int idBoletoItem,
														  DateTime dataOcorrenciaBanco,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o histórico de pagamentos do pedido devido a ocorrência 16 (título pago em cheque)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16.Parameters["@ctrl_pagto_modulo"].Value = Global.Cte.FIN.CtrlPagtoModulo.BOLETO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16.Parameters["@ctrl_pagto_id_parcela"].Value = idBoletoItem;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16.Parameters["@st_boleto_pago_cheque"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16.Parameters["@dt_ocorrencia_banco_boleto_pago_cheque"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia16);
					blnSucesso = true;
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do histórico de pagamentos do pedido durante tratamento da ocorrência 16 (título pago em cheque)!!";
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

		#region [ inserePagtoDevidoBoletoOcorrencia17 ]
		public static bool inserePagtoDevidoBoletoOcorrencia17(String usuario,
														  int idLancamentoFluxoCaixa,
														  int idBoletoItem,
														  DateTime dataCredito,
														  decimal valorPago,
														  String descricao,
														  DateTime dataOcorrenciaBanco,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Inserção no histórico de pagamento do pedido devido a boleto com ocorrência 17 (liquidação após baixa ou título não registrado)";
			bool blnGerouNsu;
			int intNsuPedidoHistPagto = 0;
			int intRetorno;
			int intCounter;
			decimal valorRateado;
			String strPedido;
			DsDataSource.DtbFinBoletoItemRateioDataTable dtbFinBoletoItemRateio;
			DsDataSource.DtbFinBoletoItemRateioRow rowFinBoletoItemRateio;
			DsDataSource.DtbFinBoletoItemRow rowFinBoletoItem;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Consistência ]
				if (idBoletoItem == 0)
				{
					strMsgErro = "Não foi informado o nº identificação do registro do boleto!!\nOperação: " + strOperacao;
					return false;
				}
				#endregion

				#region [ Obtém os dados do boleto ]
				rowFinBoletoItem = BoletoDAO.obtemRegistroBoletoItem(idBoletoItem);
				if (rowFinBoletoItem == null)
				{
					strMsgErro = "Falha ao localizar o registro do boleto (t_FIN_BOLETO_ITEM.id=" + idBoletoItem.ToString() + ")!!\nOperação: " + strOperacao;
					return false;
				}
				#endregion

				#region [ Obtém os pedidos envolvidos no rateio ]
				dtbFinBoletoItemRateio = BoletoDAO.obtemBoletoItemRateio(idBoletoItem);
				#endregion

				for (intCounter = 0; intCounter < dtbFinBoletoItemRateio.Rows.Count; intCounter++)
				{
					rowFinBoletoItemRateio = (DsDataSource.DtbFinBoletoItemRateioRow)dtbFinBoletoItemRateio.Rows[intCounter];
					strPedido = rowFinBoletoItemRateio.pedido;
					if (Global.isNumeroPedido(strPedido))
					{
						valorRateado = rowFinBoletoItemRateio.valor;

						#region [ Insere novo registro ]

						#region [ Gera NSU ]
						blnGerouNsu = BD.geraNsu(Global.Cte.FIN.NSU.T_FIN_PEDIDO_HIST_PAGTO, ref intNsuPedidoHistPagto, ref strMsgErro);
						if (!blnGerouNsu)
						{
							strMsgErro = "Falha ao gerar o NSU para o registro do histórico de pagamentos do pedido!!\n" + strMsgErro;
							return false;
						}
						#endregion

						#region [ Preenche o valor dos parâmetros ]
						cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters["@id"].Value = intNsuPedidoHistPagto;
						cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters["@pedido"].Value = strPedido;
						cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters["@status"].Value = Global.Cte.FIN.ST_T_FIN_PEDIDO_HIST_PAGTO.QUITADO;
						cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters["@id_fluxo_caixa"].Value = idLancamentoFluxoCaixa;
						cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters["@ctrl_pagto_id_parcela"].Value = idBoletoItem;
						cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters["@ctrl_pagto_modulo"].Value = Global.Cte.FIN.CtrlPagtoModulo.BOLETO;
						cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters["@dt_vencto"].Value = Global.formataDataYyyyMmDdComSeparador(dataCredito);
						cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters["@valor_total"].Value = rowFinBoletoItem.valor;
						cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters["@valor_rateado"].Value = valorRateado;
						cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters["@descricao"].Value = descricao;
						cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters["@usuario_cadastro"].Value = usuario;
						cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters["@usuario_ult_atualizacao"].Value = usuario;
						cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters["@st_boleto_ocorrencia_17"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
						cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17.Parameters["@dt_ocorrencia_banco_boleto_ocorrencia_17"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
						#endregion

						#region [ Tenta inserir o registro ]
						try
						{
							intRetorno = BD.executaNonQuery(ref cmPedidoHistPagtoInsertDevidoBoletoOcorrencia17);
						}
						catch (Exception ex)
						{
							intRetorno = 0;
							Global.gravaLogAtividade(strOperacao + "\n" + ex.ToString());
						}
						#endregion

						#region [ Processamento para sucesso ou falha desta tentativa de inserção ]
						if (intRetorno != 1)
						{
							strMsgErro = "Falha ao tentar gerar o registro automático no histórico de pagamentos do pedido devido a boleto com ocorrência 17 (liquidação após baixa ou título não registrado)!!";
							return false;
						}
						#endregion

						#endregion
					}
				}

				return true;
			}
			catch (Exception ex)
			{
				// Para o usuário, exibe uma mensagem mais sucinta
				strMsgErro = ex.Message;
				// No log em arquivo, grava o stack de erro completo
				Global.gravaLogAtividade(strOperacao + "\n" + ex.ToString());
				return false;
			}
		}
		#endregion

		#region [ atualizaPagtoDevidoBoletoOcorrencia22 ]
		public static bool atualizaPagtoDevidoBoletoOcorrencia22(String usuario,
														  int idBoletoItem,
														  DateTime dataOcorrenciaBanco,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o histórico de pagamentos do pedido devido a ocorrência 22 (título com pagamento cancelado)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22.Parameters["@ctrl_pagto_modulo"].Value = Global.Cte.FIN.CtrlPagtoModulo.BOLETO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22.Parameters["@ctrl_pagto_id_parcela"].Value = idBoletoItem;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22.Parameters["@st_boleto_pago_cheque"].Value = Global.Cte.FIN.StCampoFlag.FLAG_DESLIGADO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22.Parameters["@dt_ocorrencia_banco_boleto_pago_cheque"].Value = "";
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia22);
					blnSucesso = true;
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do histórico de pagamentos do pedido durante tratamento da ocorrência 22 (título com pagamento cancelado)!!";
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

		#region [ atualizaPagtoDevidoBoletoOcorrencia23 ]
		public static bool atualizaPagtoDevidoBoletoOcorrencia23(String usuario,
														  int idBoletoItem,
														  DateTime dataOcorrenciaBanco,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o histórico de pagamentos do pedido devido a ocorrência 23 (entrada do título em cartório)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23.Parameters["@ctrl_pagto_modulo"].Value = Global.Cte.FIN.CtrlPagtoModulo.BOLETO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23.Parameters["@ctrl_pagto_id_parcela"].Value = idBoletoItem;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23.Parameters["@st_boleto_ocorrencia_23"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23.Parameters["@dt_ocorrencia_banco_boleto_ocorrencia_23"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia23);
					blnSucesso = true;
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do histórico de pagamentos do pedido durante tratamento da ocorrência 23 (entrada do título em cartório)!!";
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

		#region [ atualizaPagtoDevidoBoletoOcorrencia34 ]
		public static bool atualizaPagtoDevidoBoletoOcorrencia34(String usuario,
														  int idBoletoItem,
														  DateTime dataOcorrenciaBanco,
														  ref String strMsgErro)
		{
			#region [ Declarações ]
			String strOperacao = "Atualiza o histórico de pagamentos do pedido devido a ocorrência 34 (retirado de cartório e manutenção carteira)";
			bool blnSucesso = false;
			int intRetorno;
			#endregion

			strMsgErro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34.Parameters["@ctrl_pagto_modulo"].Value = Global.Cte.FIN.CtrlPagtoModulo.BOLETO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34.Parameters["@ctrl_pagto_id_parcela"].Value = idBoletoItem;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34.Parameters["@st_boleto_ocorrencia_34"].Value = Global.Cte.FIN.StCampoFlag.FLAG_LIGADO;
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34.Parameters["@dt_ocorrencia_banco_boleto_ocorrencia_34"].Value = Global.formataDataYyyyMmDdComSeparador(dataOcorrenciaBanco);
				cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34.Parameters["@usuario_ult_atualizacao"].Value = usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmPedidoHistPagtoUpdateDevidoBoletoOcorrencia34);
					blnSucesso = true;
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					Global.gravaLogAtividade(strOperacao + " - Tentativa resultou em exception!!\n" + ex.ToString());
				}
				#endregion

				#region [ Processamento final de sucesso ou falha ]
				if (blnSucesso)
				{
					return true;
				}
				else
				{
					strMsgErro = "Falha ao tentar atualizar os dados do histórico de pagamentos do pedido durante tratamento da ocorrência 34 (retirado de cartório e manutenção carteira)!!";
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
