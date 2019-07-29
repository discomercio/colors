using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace FinanceiroService
{
	class BraspagDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsertPagOpComplementar;
		private static SqlCommand cmInsertPagOpComplementarXml;
		private static SqlCommand cmUpdatePagOpComplementarCaptureCreditCardResp;
		private static SqlCommand cmUpdatePagOpComplementarVoidCreditCardResp;
		private static SqlCommand cmUpdatePagOpComplementarRefundCreditCardResp;
		private static SqlCommand cmUpdatePagOpComplementarGetOrderIdDataResp;
		private static SqlCommand cmUpdatePagOpComplementarGetOrderDataResp;
		private static SqlCommand cmUpdatePagOpComplementarGetTransactionDataResp;
		private static SqlCommand cmUpdatePagPaymentGetTransactionDataResp;
		private static SqlCommand cmUpdatePagPaymentCaptureCreditCardRespSucesso;
		private static SqlCommand cmUpdatePagPaymentCaptureCreditCardRespFalha;
		private static SqlCommand cmUpdatePagPaymentVoidCreditCardRespSucesso;
		private static SqlCommand cmUpdatePagPaymentVoidCreditCardRespFalha;
		private static SqlCommand cmUpdatePagPaymentRefundCreditCardRespSucesso;
		private static SqlCommand cmUpdatePagPaymentRefundCreditCardRespRefundAccepted;
		private static SqlCommand cmUpdatePagPaymentRefundCreditCardRespFalha;
		private static SqlCommand cmUpdatePagPaymentBraspagTransactionId;
		private static SqlCommand cmUpdatePagPaymentPagtoRegPedido;
		private static SqlCommand cmUpdatePagPaymentEstornoRegPedido;
		private static SqlCommand cmUpdatePagPaymentAFFinalizado;
		private static SqlCommand cmUpdatePagPaymentRefundPendingConfirmado;
		private static SqlCommand cmUpdatePagPaymentRefundPendingFalha;
		private static SqlCommand cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva;
		private static SqlCommand cmUpdateWebhookQueryDadosComplementaresFalhaTemporaria;
		private static SqlCommand cmUpdateWebhookQueryDadosComplementaresQtdeTentativas;
		private static SqlCommand cmUpdateWebhookQueryDadosComplementaresSucesso;
		private static SqlCommand cmInsertWebhookQueryDadosComplementares;
		private static SqlCommand cmUpdateWebhookComplementarPagtoRegPedido;
		private static SqlCommand cmUpdateWebhookEmailEnviadoStatusSucesso;
		private static SqlCommand cmUpdateWebhookEmailEnviadoStatusFalha;
		private static SqlCommand cmUpdateWebhookProcessamentoErpStatusSucesso;
		private static SqlCommand cmUpdateWebhookProcessamentoErpStatusFalha;
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
		static BraspagDAO()
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

			#region [ cmInsertPagOpComplementar ]
			strSql = "INSERT INTO t_PAGTO_GW_PAG_OP_COMPLEMENTAR (" +
						"id, " +
						"id_pagto_gw_pag, " +
						"id_pagto_gw_pag_payment, " +
						"usuario, " +
						"operacao, " +
						"trx_TX_data, " +
						"trx_TX_data_hora, " +
						"req_RequestId, " +
						"req_Version, " +
						"req_MerchantId, " +
						"req_BraspagTransactionId, " +
						"req_OrderId, " +
						"req_Amount, " +
						"req_ServiceTaxAmount" +
					") VALUES (" +
						"@id, " +
						"@id_pagto_gw_pag, " +
						"@id_pagto_gw_pag_payment, " +
						"@usuario, " +
						"@operacao, " +
						Global.sqlMontaGetdateSomenteData() + ", " +
						"getdate(), " +
						"@req_RequestId, " +
						"@req_Version, " +
						"@req_MerchantId, " +
						"@req_BraspagTransactionId, " +
						"@req_OrderId, " +
						"@req_Amount, " +
						"@req_ServiceTaxAmount" +
					")";
			cmInsertPagOpComplementar = BD.criaSqlCommand();
			cmInsertPagOpComplementar.CommandText = strSql;
			cmInsertPagOpComplementar.Parameters.Add("@id", SqlDbType.Int);
			cmInsertPagOpComplementar.Parameters.Add("@id_pagto_gw_pag", SqlDbType.Int);
			cmInsertPagOpComplementar.Parameters.Add("@id_pagto_gw_pag_payment", SqlDbType.Int);
			cmInsertPagOpComplementar.Parameters.Add("@usuario", SqlDbType.VarChar, 10);
			cmInsertPagOpComplementar.Parameters.Add("@operacao", SqlDbType.VarChar, 40);
			cmInsertPagOpComplementar.Parameters.Add("@req_RequestId", SqlDbType.VarChar, 64);
			cmInsertPagOpComplementar.Parameters.Add("@req_Version", SqlDbType.VarChar, 7);
			cmInsertPagOpComplementar.Parameters.Add("@req_MerchantId", SqlDbType.VarChar, 40);
			cmInsertPagOpComplementar.Parameters.Add("@req_BraspagTransactionId", SqlDbType.VarChar, 36);
			cmInsertPagOpComplementar.Parameters.Add("@req_OrderId", SqlDbType.VarChar, 20);
			cmInsertPagOpComplementar.Parameters.Add("@req_Amount", SqlDbType.VarChar, 18);
			cmInsertPagOpComplementar.Parameters.Add("@req_ServiceTaxAmount", SqlDbType.VarChar, 18);
			cmInsertPagOpComplementar.Prepare();
			#endregion

			#region [ cmInsertPagOpComplementarXml ]
			strSql = "INSERT INTO t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML (" +
						"id, " +
						"id_pagto_gw_pag_op_complementar, " +
						"tipo_transacao, " +
						"fluxo_xml, " +
						"xml " +
					") VALUES (" +
						"@id, " +
						"@id_pagto_gw_pag_op_complementar, " +
						"@tipo_transacao, " +
						"@fluxo_xml, " +
						"@xml" +
					")";
			cmInsertPagOpComplementarXml = BD.criaSqlCommand();
			cmInsertPagOpComplementarXml.CommandText = strSql;
			cmInsertPagOpComplementarXml.Parameters.Add("@id", SqlDbType.Int);
			cmInsertPagOpComplementarXml.Parameters.Add("@id_pagto_gw_pag_op_complementar", SqlDbType.Int);
			cmInsertPagOpComplementarXml.Parameters.Add("@tipo_transacao", SqlDbType.VarChar, 20);
			cmInsertPagOpComplementarXml.Parameters.Add("@fluxo_xml", SqlDbType.VarChar, 2);
			cmInsertPagOpComplementarXml.Parameters.Add("@xml", SqlDbType.VarChar, -1); // varchar(max)
			cmInsertPagOpComplementarXml.Prepare();
			#endregion

			#region [ cmUpdatePagOpComplementarCaptureCreditCardResp ]
			strSql = "UPDATE t_PAGTO_GW_PAG_OP_COMPLEMENTAR SET" +
						" trx_RX_data = " + Global.sqlMontaGetdateSomenteData() + "," +
						" trx_RX_data_hora = getdate()," +
						" trx_RX_status = @trx_RX_status," +
						" trx_RX_vazio_status = @trx_RX_vazio_status," +
						" st_sucesso = @st_sucesso," +
						" resp_BraspagTransactionId = @resp_BraspagTransactionId," +
						" resp_AuthorizationCode = @resp_AuthorizationCode," +
						" resp_ProofOfSale = @resp_ProofOfSale," +
						" resp_Status = @resp_Status" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagOpComplementarCaptureCreditCardResp = BD.criaSqlCommand();
			cmUpdatePagOpComplementarCaptureCreditCardResp.CommandText = strSql;
			cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters.Add("@trx_RX_status", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters.Add("@trx_RX_vazio_status", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters.Add("@st_sucesso", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters.Add("@resp_BraspagTransactionId", SqlDbType.VarChar, 36);
			cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters.Add("@resp_AuthorizationCode", SqlDbType.VarChar, 32);
			cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters.Add("@resp_ProofOfSale", SqlDbType.VarChar, 32);
			cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters.Add("@resp_Status", SqlDbType.VarChar, 2);
			cmUpdatePagOpComplementarCaptureCreditCardResp.Prepare();
			#endregion

			#region [ cmUpdatePagOpComplementarGetOrderIdDataResp ]
			strSql = "UPDATE t_PAGTO_GW_PAG_OP_COMPLEMENTAR SET" +
						" trx_RX_data = " + Global.sqlMontaGetdateSomenteData() + "," +
						" trx_RX_data_hora = getdate()," +
						" trx_RX_status = @trx_RX_status," +
						" trx_RX_vazio_status = @trx_RX_vazio_status," +
						" st_sucesso = @st_sucesso," +
						" resp_BraspagTransactionId = @resp_BraspagTransactionId" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagOpComplementarGetOrderIdDataResp = BD.criaSqlCommand();
			cmUpdatePagOpComplementarGetOrderIdDataResp.CommandText = strSql;
			cmUpdatePagOpComplementarGetOrderIdDataResp.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagOpComplementarGetOrderIdDataResp.Parameters.Add("@trx_RX_status", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarGetOrderIdDataResp.Parameters.Add("@trx_RX_vazio_status", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarGetOrderIdDataResp.Parameters.Add("@st_sucesso", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarGetOrderIdDataResp.Parameters.Add("@resp_BraspagTransactionId", SqlDbType.VarChar, 36);
			cmUpdatePagOpComplementarGetOrderIdDataResp.Prepare();
			#endregion

			#region [ cmUpdatePagOpComplementarGetOrderDataResp ]
			strSql = "UPDATE t_PAGTO_GW_PAG_OP_COMPLEMENTAR SET" +
						" trx_RX_data = " + Global.sqlMontaGetdateSomenteData() + "," +
						" trx_RX_data_hora = getdate()," +
						" trx_RX_status = @trx_RX_status," +
						" trx_RX_vazio_status = @trx_RX_vazio_status," +
						" st_sucesso = @st_sucesso," +
						" resp_BraspagTransactionId = @resp_BraspagTransactionId" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagOpComplementarGetOrderDataResp = BD.criaSqlCommand();
			cmUpdatePagOpComplementarGetOrderDataResp.CommandText = strSql;
			cmUpdatePagOpComplementarGetOrderDataResp.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagOpComplementarGetOrderDataResp.Parameters.Add("@trx_RX_status", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarGetOrderDataResp.Parameters.Add("@trx_RX_vazio_status", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarGetOrderDataResp.Parameters.Add("@st_sucesso", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarGetOrderDataResp.Parameters.Add("@resp_BraspagTransactionId", SqlDbType.VarChar, 36);
			cmUpdatePagOpComplementarGetOrderDataResp.Prepare();
			#endregion

			#region [ cmUpdatePagOpComplementarGetTransactionDataResp ]
			strSql = "UPDATE t_PAGTO_GW_PAG_OP_COMPLEMENTAR SET" +
						" trx_RX_data = " + Global.sqlMontaGetdateSomenteData() + "," +
						" trx_RX_data_hora = getdate()," +
						" trx_RX_status = @trx_RX_status," +
						" trx_RX_vazio_status = @trx_RX_vazio_status," +
						" st_sucesso = @st_sucesso," +
						" resp_BraspagTransactionId = @resp_BraspagTransactionId," +
						" resp_AuthorizationCode = @resp_AuthorizationCode," +
						" resp_ProofOfSale = @resp_ProofOfSale," +
						" resp_Status = @resp_Status" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagOpComplementarGetTransactionDataResp = BD.criaSqlCommand();
			cmUpdatePagOpComplementarGetTransactionDataResp.CommandText = strSql;
			cmUpdatePagOpComplementarGetTransactionDataResp.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagOpComplementarGetTransactionDataResp.Parameters.Add("@trx_RX_status", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarGetTransactionDataResp.Parameters.Add("@trx_RX_vazio_status", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarGetTransactionDataResp.Parameters.Add("@st_sucesso", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarGetTransactionDataResp.Parameters.Add("@resp_BraspagTransactionId", SqlDbType.VarChar, 36);
			cmUpdatePagOpComplementarGetTransactionDataResp.Parameters.Add("@resp_AuthorizationCode", SqlDbType.VarChar, 32);
			cmUpdatePagOpComplementarGetTransactionDataResp.Parameters.Add("@resp_ProofOfSale", SqlDbType.VarChar, 32);
			cmUpdatePagOpComplementarGetTransactionDataResp.Parameters.Add("@resp_Status", SqlDbType.VarChar, 2);
			cmUpdatePagOpComplementarGetTransactionDataResp.Prepare();
			#endregion

			#region [ cmUpdatePagPaymentGetTransactionDataResp ]
			strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" +
						" ult_GlobalStatus = @ult_GlobalStatus," +
						" ult_atualizacao_data_hora = getdate()," +
						" ult_atualizacao_usuario = @ult_atualizacao_usuario," +
						" ult_id_pagto_gw_pag_payment_op_complementar = @ult_id_pagto_gw_pag_payment_op_complementar," +
						" resp_CapturedDate = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@resp_CapturedDate") + "," +
						" resp_VoidedDate = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@resp_VoidedDate") +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagPaymentGetTransactionDataResp = BD.criaSqlCommand();
			cmUpdatePagPaymentGetTransactionDataResp.CommandText = strSql;
			cmUpdatePagPaymentGetTransactionDataResp.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagPaymentGetTransactionDataResp.Parameters.Add("@ult_GlobalStatus", SqlDbType.VarChar, 5);
			cmUpdatePagPaymentGetTransactionDataResp.Parameters.Add("@ult_atualizacao_usuario", SqlDbType.VarChar, 10);
			cmUpdatePagPaymentGetTransactionDataResp.Parameters.Add("@ult_id_pagto_gw_pag_payment_op_complementar", SqlDbType.Int);
			cmUpdatePagPaymentGetTransactionDataResp.Parameters.Add("@resp_CapturedDate", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentGetTransactionDataResp.Parameters.Add("@resp_VoidedDate", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentGetTransactionDataResp.Prepare();
			#endregion

			#region [ cmUpdatePagPaymentBraspagTransactionId ]
			strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" +
						" resp_PaymentDataResponse_BraspagTransactionId = @BraspagTransactionId" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagPaymentBraspagTransactionId = BD.criaSqlCommand();
			cmUpdatePagPaymentBraspagTransactionId.CommandText = strSql;
			cmUpdatePagPaymentBraspagTransactionId.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagPaymentBraspagTransactionId.Parameters.Add("@BraspagTransactionId", SqlDbType.VarChar, 36);
			cmUpdatePagPaymentBraspagTransactionId.Prepare();
			#endregion

			#region [ cmUpdatePagPaymentCaptureCreditCardRespSucesso ]
			strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" +
						" st_processamento_PAG_finalizado = 1," +
						" dt_hr_processamento_PAG_finalizado = getdate()," +
						" ult_GlobalStatus = @ult_GlobalStatus," +
						" resp_CapturedDate = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@resp_CapturedDate") + "," +
						" captura_confirmada_status = @captura_confirmada_status," +
						" captura_confirmada_data = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@captura_confirmada_data") + "," +
						" captura_confirmada_data_hora = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@captura_confirmada_data_hora") + "," +
						" captura_confirmada_usuario = @captura_confirmada_usuario," +
						" ult_atualizacao_data_hora = getdate()," +
						" ult_atualizacao_usuario = @ult_atualizacao_usuario," +
						" ult_id_pagto_gw_pag_payment_op_complementar = @ult_id_pagto_gw_pag_payment_op_complementar" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagPaymentCaptureCreditCardRespSucesso = BD.criaSqlCommand();
			cmUpdatePagPaymentCaptureCreditCardRespSucesso.CommandText = strSql;
			cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters.Add("@ult_GlobalStatus", SqlDbType.VarChar, 5);
			cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters.Add("@resp_CapturedDate", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters.Add("@captura_confirmada_status", SqlDbType.TinyInt);
			cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters.Add("@captura_confirmada_data", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters.Add("@captura_confirmada_data_hora", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters.Add("@captura_confirmada_usuario", SqlDbType.VarChar, 10);
			cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters.Add("@ult_atualizacao_usuario", SqlDbType.VarChar, 10);
			cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters.Add("@ult_id_pagto_gw_pag_payment_op_complementar", SqlDbType.Int);
			cmUpdatePagPaymentCaptureCreditCardRespSucesso.Prepare();
			#endregion

			#region [ cmUpdatePagPaymentCaptureCreditCardRespFalha ]
			strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" +
						" captura_confirmada_erro_status = @captura_confirmada_erro_status," +
						" captura_confirmada_erro_data = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@captura_confirmada_erro_data") + "," +
						" captura_confirmada_erro_data_hora = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@captura_confirmada_erro_data_hora") + "," +
						" captura_confirmada_erro_mensagem = @captura_confirmada_erro_mensagem," +
						" ult_atualizacao_data_hora = getdate()," +
						" ult_atualizacao_usuario = @ult_atualizacao_usuario," +
						" ult_id_pagto_gw_pag_payment_op_complementar = @ult_id_pagto_gw_pag_payment_op_complementar" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagPaymentCaptureCreditCardRespFalha = BD.criaSqlCommand();
			cmUpdatePagPaymentCaptureCreditCardRespFalha.CommandText = strSql;
			cmUpdatePagPaymentCaptureCreditCardRespFalha.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagPaymentCaptureCreditCardRespFalha.Parameters.Add("@captura_confirmada_erro_status", SqlDbType.TinyInt);
			cmUpdatePagPaymentCaptureCreditCardRespFalha.Parameters.Add("@captura_confirmada_erro_data", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentCaptureCreditCardRespFalha.Parameters.Add("@captura_confirmada_erro_data_hora", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentCaptureCreditCardRespFalha.Parameters.Add("@captura_confirmada_erro_mensagem", SqlDbType.VarChar, -1); // varchar(max)
			cmUpdatePagPaymentCaptureCreditCardRespFalha.Parameters.Add("@ult_atualizacao_usuario", SqlDbType.VarChar, 10);
			cmUpdatePagPaymentCaptureCreditCardRespFalha.Parameters.Add("@ult_id_pagto_gw_pag_payment_op_complementar", SqlDbType.Int);
			cmUpdatePagPaymentCaptureCreditCardRespFalha.Prepare();
			#endregion

			#region [ cmUpdatePagPaymentVoidCreditCardRespSucesso ]
			strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" +
						" st_processamento_PAG_finalizado = 1," +
						" dt_hr_processamento_PAG_finalizado = getdate()," +
						" ult_GlobalStatus = @ult_GlobalStatus," +
						" resp_VoidedDate = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@resp_VoidedDate") + "," +
						" voided_status = @voided_status," +
						" voided_data = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@voided_data") + "," +
						" voided_data_hora = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@voided_data_hora") + "," +
						" voided_usuario = @voided_usuario," +
						" ult_atualizacao_data_hora = getdate()," +
						" ult_atualizacao_usuario = @ult_atualizacao_usuario," +
						" ult_id_pagto_gw_pag_payment_op_complementar = @ult_id_pagto_gw_pag_payment_op_complementar" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagPaymentVoidCreditCardRespSucesso = BD.criaSqlCommand();
			cmUpdatePagPaymentVoidCreditCardRespSucesso.CommandText = strSql;
			cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters.Add("@ult_GlobalStatus", SqlDbType.VarChar, 5);
			cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters.Add("@resp_VoidedDate", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters.Add("@voided_status", SqlDbType.TinyInt);
			cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters.Add("@voided_data", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters.Add("@voided_data_hora", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters.Add("@voided_usuario", SqlDbType.VarChar, 10);
			cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters.Add("@ult_atualizacao_usuario", SqlDbType.VarChar, 10);
			cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters.Add("@ult_id_pagto_gw_pag_payment_op_complementar", SqlDbType.Int);
			cmUpdatePagPaymentVoidCreditCardRespSucesso.Prepare();
			#endregion

			#region [ cmUpdatePagPaymentVoidCreditCardRespFalha ]
			strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" +
						" voided_erro_status = @voided_erro_status," +
						" voided_erro_data = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@voided_erro_data") + "," +
						" voided_erro_data_hora = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@voided_erro_data_hora") + "," +
						" voided_erro_mensagem = @voided_erro_mensagem," +
						" ult_atualizacao_data_hora = getdate()," +
						" ult_atualizacao_usuario = @ult_atualizacao_usuario," +
						" ult_id_pagto_gw_pag_payment_op_complementar = @ult_id_pagto_gw_pag_payment_op_complementar" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagPaymentVoidCreditCardRespFalha = BD.criaSqlCommand();
			cmUpdatePagPaymentVoidCreditCardRespFalha.CommandText = strSql;
			cmUpdatePagPaymentVoidCreditCardRespFalha.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagPaymentVoidCreditCardRespFalha.Parameters.Add("@voided_erro_status", SqlDbType.TinyInt);
			cmUpdatePagPaymentVoidCreditCardRespFalha.Parameters.Add("@voided_erro_data", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentVoidCreditCardRespFalha.Parameters.Add("@voided_erro_data_hora", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentVoidCreditCardRespFalha.Parameters.Add("@voided_erro_mensagem", SqlDbType.VarChar, -1); // varchar(max)
			cmUpdatePagPaymentVoidCreditCardRespFalha.Parameters.Add("@ult_atualizacao_usuario", SqlDbType.VarChar, 10);
			cmUpdatePagPaymentVoidCreditCardRespFalha.Parameters.Add("@ult_id_pagto_gw_pag_payment_op_complementar", SqlDbType.Int);
			cmUpdatePagPaymentVoidCreditCardRespFalha.Prepare();
			#endregion

			#region [ cmUpdatePagPaymentRefundCreditCardRespSucesso ]
			strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" +
						" st_processamento_PAG_finalizado = 1," +
						" dt_hr_processamento_PAG_finalizado = getdate()," +
						" ult_GlobalStatus = @ult_GlobalStatus," +
						" resp_VoidedDate = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@resp_VoidedDate") + "," +
						" refunded_status = @refunded_status," +
						" refunded_data = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@refunded_data") + "," +
						" refunded_data_hora = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@refunded_data_hora") + "," +
						" refunded_usuario = @refunded_usuario," +
						" ult_atualizacao_data_hora = getdate()," +
						" ult_atualizacao_usuario = @ult_atualizacao_usuario," +
						" ult_id_pagto_gw_pag_payment_op_complementar = @ult_id_pagto_gw_pag_payment_op_complementar" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagPaymentRefundCreditCardRespSucesso = BD.criaSqlCommand();
			cmUpdatePagPaymentRefundCreditCardRespSucesso.CommandText = strSql;
			cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters.Add("@ult_GlobalStatus", SqlDbType.VarChar, 5);
			cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters.Add("@resp_VoidedDate", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters.Add("@refunded_status", SqlDbType.TinyInt);
			cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters.Add("@refunded_data", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters.Add("@refunded_data_hora", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters.Add("@refunded_usuario", SqlDbType.VarChar, 10);
			cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters.Add("@ult_atualizacao_usuario", SqlDbType.VarChar, 10);
			cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters.Add("@ult_id_pagto_gw_pag_payment_op_complementar", SqlDbType.Int);
			cmUpdatePagPaymentRefundCreditCardRespSucesso.Prepare();
			#endregion

			#region [ cmUpdatePagPaymentRefundCreditCardRespRefundAccepted ]
			strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" +
						" refund_pending_status = 1,"+
						" refund_pending_data = " + Global.sqlMontaGetdateSomenteData() + "," +
						" refund_pending_data_hora = getdate(),"+
						" refund_pending_usuario = @refund_pending_usuario,"+
						" ult_GlobalStatus = @ult_GlobalStatus," +
						" ult_atualizacao_data_hora = getdate()," +
						" ult_atualizacao_usuario = @ult_atualizacao_usuario," +
						" ult_id_pagto_gw_pag_payment_op_complementar = @ult_id_pagto_gw_pag_payment_op_complementar" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagPaymentRefundCreditCardRespRefundAccepted = BD.criaSqlCommand();
			cmUpdatePagPaymentRefundCreditCardRespRefundAccepted.CommandText = strSql;
			cmUpdatePagPaymentRefundCreditCardRespRefundAccepted.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagPaymentRefundCreditCardRespRefundAccepted.Parameters.Add("@refund_pending_usuario", SqlDbType.VarChar, 10);
			cmUpdatePagPaymentRefundCreditCardRespRefundAccepted.Parameters.Add("@ult_GlobalStatus", SqlDbType.VarChar, 5);
			cmUpdatePagPaymentRefundCreditCardRespRefundAccepted.Parameters.Add("@ult_atualizacao_usuario", SqlDbType.VarChar, 10);
			cmUpdatePagPaymentRefundCreditCardRespRefundAccepted.Parameters.Add("@ult_id_pagto_gw_pag_payment_op_complementar", SqlDbType.Int);
			cmUpdatePagPaymentRefundCreditCardRespRefundAccepted.Prepare();
			#endregion

			#region [ cmUpdatePagPaymentRefundCreditCardRespFalha ]
			strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" +
						" refunded_erro_status = @refunded_erro_status," +
						" refunded_erro_data = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@refunded_erro_data") + "," +
						" refunded_erro_data_hora = " + Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@refunded_erro_data_hora") + "," +
						" refunded_erro_mensagem = @refunded_erro_mensagem," +
						" ult_atualizacao_data_hora = getdate()," +
						" ult_atualizacao_usuario = @ult_atualizacao_usuario," +
						" ult_id_pagto_gw_pag_payment_op_complementar = @ult_id_pagto_gw_pag_payment_op_complementar" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagPaymentRefundCreditCardRespFalha = BD.criaSqlCommand();
			cmUpdatePagPaymentRefundCreditCardRespFalha.CommandText = strSql;
			cmUpdatePagPaymentRefundCreditCardRespFalha.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagPaymentRefundCreditCardRespFalha.Parameters.Add("@refunded_erro_status", SqlDbType.TinyInt);
			cmUpdatePagPaymentRefundCreditCardRespFalha.Parameters.Add("@refunded_erro_data", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentRefundCreditCardRespFalha.Parameters.Add("@refunded_erro_data_hora", SqlDbType.VarChar, 19);
			cmUpdatePagPaymentRefundCreditCardRespFalha.Parameters.Add("@refunded_erro_mensagem", SqlDbType.VarChar, -1); // varchar(max)
			cmUpdatePagPaymentRefundCreditCardRespFalha.Parameters.Add("@ult_atualizacao_usuario", SqlDbType.VarChar, 10);
			cmUpdatePagPaymentRefundCreditCardRespFalha.Parameters.Add("@ult_id_pagto_gw_pag_payment_op_complementar", SqlDbType.Int);
			cmUpdatePagPaymentRefundCreditCardRespFalha.Prepare();
			#endregion

			#region [ cmUpdatePagOpComplementarVoidCreditCardResp ]
			strSql = "UPDATE t_PAGTO_GW_PAG_OP_COMPLEMENTAR SET" +
						" trx_RX_data = " + Global.sqlMontaGetdateSomenteData() + "," +
						" trx_RX_data_hora = getdate()," +
						" trx_RX_status = @trx_RX_status," +
						" trx_RX_vazio_status = @trx_RX_vazio_status," +
						" st_sucesso = @st_sucesso," +
						" resp_BraspagTransactionId = @resp_BraspagTransactionId," +
						" resp_AuthorizationCode = @resp_AuthorizationCode," +
						" resp_ProofOfSale = @resp_ProofOfSale," +
						" resp_Status = @resp_Status" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagOpComplementarVoidCreditCardResp = BD.criaSqlCommand();
			cmUpdatePagOpComplementarVoidCreditCardResp.CommandText = strSql;
			cmUpdatePagOpComplementarVoidCreditCardResp.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagOpComplementarVoidCreditCardResp.Parameters.Add("@trx_RX_status", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarVoidCreditCardResp.Parameters.Add("@trx_RX_vazio_status", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarVoidCreditCardResp.Parameters.Add("@st_sucesso", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarVoidCreditCardResp.Parameters.Add("@resp_BraspagTransactionId", SqlDbType.VarChar, 36);
			cmUpdatePagOpComplementarVoidCreditCardResp.Parameters.Add("@resp_AuthorizationCode", SqlDbType.VarChar, 32);
			cmUpdatePagOpComplementarVoidCreditCardResp.Parameters.Add("@resp_ProofOfSale", SqlDbType.VarChar, 32);
			cmUpdatePagOpComplementarVoidCreditCardResp.Parameters.Add("@resp_Status", SqlDbType.VarChar, 2);
			cmUpdatePagOpComplementarVoidCreditCardResp.Prepare();
			#endregion

			#region [ cmUpdatePagOpComplementarRefundCreditCardResp ]
			strSql = "UPDATE t_PAGTO_GW_PAG_OP_COMPLEMENTAR SET" +
						" trx_RX_data = " + Global.sqlMontaGetdateSomenteData() + "," +
						" trx_RX_data_hora = getdate()," +
						" trx_RX_status = @trx_RX_status," +
						" trx_RX_vazio_status = @trx_RX_vazio_status," +
						" st_sucesso = @st_sucesso," +
						" resp_BraspagTransactionId = @resp_BraspagTransactionId," +
						" resp_AuthorizationCode = @resp_AuthorizationCode," +
						" resp_ProofOfSale = @resp_ProofOfSale," +
						" resp_Status = @resp_Status" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagOpComplementarRefundCreditCardResp = BD.criaSqlCommand();
			cmUpdatePagOpComplementarRefundCreditCardResp.CommandText = strSql;
			cmUpdatePagOpComplementarRefundCreditCardResp.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagOpComplementarRefundCreditCardResp.Parameters.Add("@trx_RX_status", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarRefundCreditCardResp.Parameters.Add("@trx_RX_vazio_status", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarRefundCreditCardResp.Parameters.Add("@st_sucesso", SqlDbType.TinyInt);
			cmUpdatePagOpComplementarRefundCreditCardResp.Parameters.Add("@resp_BraspagTransactionId", SqlDbType.VarChar, 36);
			cmUpdatePagOpComplementarRefundCreditCardResp.Parameters.Add("@resp_AuthorizationCode", SqlDbType.VarChar, 32);
			cmUpdatePagOpComplementarRefundCreditCardResp.Parameters.Add("@resp_ProofOfSale", SqlDbType.VarChar, 32);
			cmUpdatePagOpComplementarRefundCreditCardResp.Parameters.Add("@resp_Status", SqlDbType.VarChar, 2);
			cmUpdatePagOpComplementarRefundCreditCardResp.Prepare();
			#endregion

			#region [ cmUpdatePagPaymentPagtoRegPedido ]
			strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" +
						" pagto_registrado_no_pedido_status = @pagto_registrado_no_pedido_status," +
						" pagto_registrado_no_pedido_tipo_operacao = @pagto_registrado_no_pedido_tipo_operacao," +
						" pagto_registrado_no_pedido_data = " + Global.sqlMontaGetdateSomenteData() + "," +
						" pagto_registrado_no_pedido_data_hora = getdate()," +
						" pagto_registrado_no_pedido_usuario = @pagto_registrado_no_pedido_usuario," +
						" pagto_registrado_no_pedido_id_pedido_pagamento = @pagto_registrado_no_pedido_id_pedido_pagamento," +
						" pagto_registrado_no_pedido_st_pagto_anterior = @pagto_registrado_no_pedido_st_pagto_anterior," +
						" pagto_registrado_no_pedido_st_pagto_novo = @pagto_registrado_no_pedido_st_pagto_novo" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagPaymentPagtoRegPedido = BD.criaSqlCommand();
			cmUpdatePagPaymentPagtoRegPedido.CommandText = strSql;
			cmUpdatePagPaymentPagtoRegPedido.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagPaymentPagtoRegPedido.Parameters.Add("@pagto_registrado_no_pedido_status", SqlDbType.TinyInt);
			cmUpdatePagPaymentPagtoRegPedido.Parameters.Add("@pagto_registrado_no_pedido_tipo_operacao", SqlDbType.VarChar, 3);
			cmUpdatePagPaymentPagtoRegPedido.Parameters.Add("@pagto_registrado_no_pedido_usuario", SqlDbType.VarChar, 10);
			cmUpdatePagPaymentPagtoRegPedido.Parameters.Add("@pagto_registrado_no_pedido_id_pedido_pagamento", SqlDbType.VarChar, 12);
			cmUpdatePagPaymentPagtoRegPedido.Parameters.Add("@pagto_registrado_no_pedido_st_pagto_anterior", SqlDbType.VarChar, 1);
			cmUpdatePagPaymentPagtoRegPedido.Parameters.Add("@pagto_registrado_no_pedido_st_pagto_novo", SqlDbType.VarChar, 1);
			cmUpdatePagPaymentPagtoRegPedido.Prepare();
			#endregion

			#region [ cmUpdatePagPaymentEstornoRegPedido ]
			strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" +
						" estorno_registrado_no_pedido_status = @estorno_registrado_no_pedido_status," +
						" estorno_registrado_no_pedido_tipo_operacao = @estorno_registrado_no_pedido_tipo_operacao," +
						" estorno_registrado_no_pedido_data = " + Global.sqlMontaGetdateSomenteData() + "," +
						" estorno_registrado_no_pedido_data_hora = getdate()," +
						" estorno_registrado_no_pedido_usuario = @estorno_registrado_no_pedido_usuario," +
						" estorno_registrado_no_pedido_id_pedido_pagamento = @estorno_registrado_no_pedido_id_pedido_pagamento," +
						" estorno_registrado_no_pedido_st_pagto_anterior = @estorno_registrado_no_pedido_st_pagto_anterior," +
						" estorno_registrado_no_pedido_st_pagto_novo = @estorno_registrado_no_pedido_st_pagto_novo" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagPaymentEstornoRegPedido = BD.criaSqlCommand();
			cmUpdatePagPaymentEstornoRegPedido.CommandText = strSql;
			cmUpdatePagPaymentEstornoRegPedido.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagPaymentEstornoRegPedido.Parameters.Add("@estorno_registrado_no_pedido_status", SqlDbType.TinyInt);
			cmUpdatePagPaymentEstornoRegPedido.Parameters.Add("@estorno_registrado_no_pedido_tipo_operacao", SqlDbType.VarChar, 3);
			cmUpdatePagPaymentEstornoRegPedido.Parameters.Add("@estorno_registrado_no_pedido_usuario", SqlDbType.VarChar, 10);
			cmUpdatePagPaymentEstornoRegPedido.Parameters.Add("@estorno_registrado_no_pedido_id_pedido_pagamento", SqlDbType.VarChar, 12);
			cmUpdatePagPaymentEstornoRegPedido.Parameters.Add("@estorno_registrado_no_pedido_st_pagto_anterior", SqlDbType.VarChar, 1);
			cmUpdatePagPaymentEstornoRegPedido.Parameters.Add("@estorno_registrado_no_pedido_st_pagto_novo", SqlDbType.VarChar, 1);
			cmUpdatePagPaymentEstornoRegPedido.Prepare();
			#endregion

			#region [ cmUpdatePagPaymentAFFinalizado ]
			strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" +
						" st_processamento_AF_finalizado = 1," +
						" dt_hr_processamento_AF_finalizado = getdate()" +
					" WHERE" +
						" (id = @id)" +
						" AND (st_processamento_AF_finalizado = 0)";
			cmUpdatePagPaymentAFFinalizado = BD.criaSqlCommand();
			cmUpdatePagPaymentAFFinalizado.CommandText = strSql;
			cmUpdatePagPaymentAFFinalizado.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagPaymentAFFinalizado.Prepare();
			#endregion

			#region [ cmUpdatePagPaymentRefundPendingConfirmado ]
			strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" +
						" refund_pending_confirmado_status = 1," +
						" refund_pending_confirmado_data = " + Global.sqlMontaGetdateSomenteData() + "," +
						" refund_pending_confirmado_data_hora = getdate()," +
						" refund_pending_confirmado_usuario = @refund_pending_confirmado_usuario" +
					" WHERE" +
						" (id = @id)" +
						" AND (refund_pending_status = 1)" +
						" AND (refund_pending_confirmado_status = 0)" +
						" AND (refund_pending_falha_status = 0)";
			cmUpdatePagPaymentRefundPendingConfirmado = BD.criaSqlCommand();
			cmUpdatePagPaymentRefundPendingConfirmado.CommandText = strSql;
			cmUpdatePagPaymentRefundPendingConfirmado.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagPaymentRefundPendingConfirmado.Parameters.Add("@refund_pending_confirmado_usuario", SqlDbType.VarChar, 10);
			cmUpdatePagPaymentRefundPendingConfirmado.Prepare();
			#endregion

			#region [ cmUpdatePagPaymentRefundPendingFalha ]
			strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET" +
						" refund_pending_falha_status = 1," +
						" refund_pending_falha_data = " + Global.sqlMontaGetdateSomenteData() + "," +
						" refund_pending_falha_data_hora = getdate()," +
						" refund_pending_falha_motivo = @refund_pending_falha_motivo" +
					" WHERE" +
						" (id = @id)" +
						" AND (refund_pending_status = 1)" +
						" AND (refund_pending_confirmado_status = 0)" +
						" AND (refund_pending_falha_status = 0)";
			cmUpdatePagPaymentRefundPendingFalha = BD.criaSqlCommand();
			cmUpdatePagPaymentRefundPendingFalha.CommandText = strSql;
			cmUpdatePagPaymentRefundPendingFalha.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagPaymentRefundPendingFalha.Parameters.Add("@refund_pending_falha_motivo", SqlDbType.VarChar, -1); // varchar(max)
			cmUpdatePagPaymentRefundPendingFalha.Prepare();
			#endregion

			#region [ cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva ]
			strSql = "UPDATE t_BRASPAG_WEBHOOK SET" +
						" EmailEnviadoStatus = @EmailEnviadoStatus," +
						" BraspagDadosComplementaresQueryStatus = @BraspagDadosComplementaresQueryStatus," +
						" BraspagDadosComplementaresQueryDataHora = getdate()," +
						" BraspagDadosComplementaresQueryTentativas = @BraspagDadosComplementaresQueryTentativas," +
						" BraspagDadosComplementaresQueryDtHrUltTentativa = getdate()," +
						" MsgErro = @MsgErro" +
					" WHERE" +
						" (Id = @Id)";
			cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva = BD.criaSqlCommand();
			cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva.CommandText = strSql;
			cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva.Parameters.Add("@Id", SqlDbType.Int);
			cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva.Parameters.Add("@EmailEnviadoStatus", SqlDbType.TinyInt);
			cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva.Parameters.Add("@BraspagDadosComplementaresQueryStatus", SqlDbType.TinyInt);
			cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva.Parameters.Add("@BraspagDadosComplementaresQueryTentativas", SqlDbType.Int);
			cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva.Parameters.Add("@MsgErro", SqlDbType.VarChar, -1); // varchar(max)
			cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva.Prepare();
			#endregion

			#region [ cmUpdateWebhookQueryDadosComplementaresFalhaTemporaria ]
			strSql = "UPDATE t_BRASPAG_WEBHOOK SET" +
						" BraspagDadosComplementaresQueryStatus = @BraspagDadosComplementaresQueryStatus," +
						" BraspagDadosComplementaresQueryDataHora = getdate()," +
						" BraspagDadosComplementaresQueryTentativas = @BraspagDadosComplementaresQueryTentativas," +
						" BraspagDadosComplementaresQueryDtHrUltTentativa = getdate()," +
						" MsgErroTemporario = @MsgErroTemporario" +
					" WHERE" +
						" (Id = @Id)";
			cmUpdateWebhookQueryDadosComplementaresFalhaTemporaria = BD.criaSqlCommand();
			cmUpdateWebhookQueryDadosComplementaresFalhaTemporaria.CommandText = strSql;
			cmUpdateWebhookQueryDadosComplementaresFalhaTemporaria.Parameters.Add("@Id", SqlDbType.Int);
			cmUpdateWebhookQueryDadosComplementaresFalhaTemporaria.Parameters.Add("@BraspagDadosComplementaresQueryStatus", SqlDbType.TinyInt);
			cmUpdateWebhookQueryDadosComplementaresFalhaTemporaria.Parameters.Add("@BraspagDadosComplementaresQueryTentativas", SqlDbType.Int);
			cmUpdateWebhookQueryDadosComplementaresFalhaTemporaria.Parameters.Add("@MsgErroTemporario", SqlDbType.VarChar, -1); // varchar(max)
			cmUpdateWebhookQueryDadosComplementaresFalhaTemporaria.Prepare();
			#endregion

			#region [ cmUpdateWebhookQueryDadosComplementaresQtdeTentativas ]
			strSql = "UPDATE t_BRASPAG_WEBHOOK SET" +
						" BraspagDadosComplementaresQueryTentativas = @BraspagDadosComplementaresQueryTentativas," +
						" BraspagDadosComplementaresQueryDtHrUltTentativa = getdate()" +
					" WHERE" +
						" (Id = @Id)";
			cmUpdateWebhookQueryDadosComplementaresQtdeTentativas = BD.criaSqlCommand();
			cmUpdateWebhookQueryDadosComplementaresQtdeTentativas.CommandText = strSql;
			cmUpdateWebhookQueryDadosComplementaresQtdeTentativas.Parameters.Add("@Id", SqlDbType.Int);
			cmUpdateWebhookQueryDadosComplementaresQtdeTentativas.Parameters.Add("@BraspagDadosComplementaresQueryTentativas", SqlDbType.Int);
			cmUpdateWebhookQueryDadosComplementaresQtdeTentativas.Prepare();
			#endregion

			#region [ cmUpdateWebhookQueryDadosComplementaresSucesso ]
			strSql = "UPDATE t_BRASPAG_WEBHOOK SET" +
						" BraspagDadosComplementaresQueryStatus = @BraspagDadosComplementaresQueryStatus," +
						" BraspagDadosComplementaresQueryDataHora = getdate()," +
						" BraspagDadosComplementaresQueryTentativas = @BraspagDadosComplementaresQueryTentativas," +
						" BraspagDadosComplementaresQueryDtHrUltTentativa = getdate()" +
					" WHERE" +
						" (Id = @Id)";
			cmUpdateWebhookQueryDadosComplementaresSucesso = BD.criaSqlCommand();
			cmUpdateWebhookQueryDadosComplementaresSucesso.CommandText = strSql;
			cmUpdateWebhookQueryDadosComplementaresSucesso.Parameters.Add("@Id", SqlDbType.Int);
			cmUpdateWebhookQueryDadosComplementaresSucesso.Parameters.Add("@BraspagDadosComplementaresQueryStatus", SqlDbType.TinyInt);
			cmUpdateWebhookQueryDadosComplementaresSucesso.Parameters.Add("@BraspagDadosComplementaresQueryTentativas", SqlDbType.Int);
			cmUpdateWebhookQueryDadosComplementaresSucesso.Prepare();
			#endregion

			#region [ cmInsertWebhookQueryDadosComplementares ]
			strSql = "INSERT INTO t_BRASPAG_WEBHOOK_COMPLEMENTAR (" +
						"id_braspag_webhook, " +
						"BraspagTransactionId, " +
						"BraspagOrderId, " +
						"PaymentMethod, " +
						"GlobalStatus, " +
						"ReceivedDate, " +
						"CapturedDate, " +
						"CustomerName, " +
						"BoletoExpirationDate, " +
						"Amount, " +
						"ValorAmount, " +
						"PaidAmount, " +
						"ValorPaidAmount, " +
						"pedido" +
					")" +
					" OUTPUT INSERTED.Id" +
					" VALUES " +
					"(" +
						"@id_braspag_webhook, " +
						"@BraspagTransactionId, " +
						"@BraspagOrderId, " +
						"@PaymentMethod, " +
						"@GlobalStatus, " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@ReceivedDate") + ", " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@CapturedDate") + ", " +
						"@CustomerName, " +
						Global.sqlMontaCaseWhenParametroStringVaziaComoNull("@BoletoExpirationDate") + ", " +
						"@Amount, " +
						"@ValorAmount, " +
						"@PaidAmount, " +
						"@ValorPaidAmount, " +
						"@pedido" +
					")";
			cmInsertWebhookQueryDadosComplementares = BD.criaSqlCommand();
			cmInsertWebhookQueryDadosComplementares.CommandText = strSql;
			cmInsertWebhookQueryDadosComplementares.Parameters.Add("@id_braspag_webhook", SqlDbType.Int);
			cmInsertWebhookQueryDadosComplementares.Parameters.Add("@BraspagTransactionId", SqlDbType.VarChar, 36);
			cmInsertWebhookQueryDadosComplementares.Parameters.Add("@BraspagOrderId", SqlDbType.VarChar, 36);
			cmInsertWebhookQueryDadosComplementares.Parameters.Add("@PaymentMethod", SqlDbType.VarChar, 3);
			cmInsertWebhookQueryDadosComplementares.Parameters.Add("@GlobalStatus", SqlDbType.VarChar, 5);
			cmInsertWebhookQueryDadosComplementares.Parameters.Add("@ReceivedDate", SqlDbType.VarChar, 19);
			cmInsertWebhookQueryDadosComplementares.Parameters.Add("@CapturedDate", SqlDbType.VarChar, 19);
			cmInsertWebhookQueryDadosComplementares.Parameters.Add("@CustomerName", SqlDbType.VarChar, 80);
			cmInsertWebhookQueryDadosComplementares.Parameters.Add("@BoletoExpirationDate", SqlDbType.VarChar, 19);
			cmInsertWebhookQueryDadosComplementares.Parameters.Add("@Amount", SqlDbType.VarChar, 18);
			cmInsertWebhookQueryDadosComplementares.Parameters.Add("@ValorAmount", SqlDbType.Money);
			cmInsertWebhookQueryDadosComplementares.Parameters.Add("@PaidAmount", SqlDbType.VarChar, 18);
			cmInsertWebhookQueryDadosComplementares.Parameters.Add("@ValorPaidAmount", SqlDbType.Money);
			cmInsertWebhookQueryDadosComplementares.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmInsertWebhookQueryDadosComplementares.Prepare();
			#endregion

			#region [ cmUpdateWebhookComplementarPagtoRegPedido ]
			strSql = "UPDATE t_BRASPAG_WEBHOOK_COMPLEMENTAR SET" +
						" PagtoRegistradoNoPedidoStatus = @PagtoRegistradoNoPedidoStatus," +
						" PagtoRegistradoNoPedidoTipoOperacao = @PagtoRegistradoNoPedidoTipoOperacao," +
						" PagtoRegistradoNoPedidoData = " + Global.sqlMontaGetdateSomenteData() + "," +
						" PagtoRegistradoNoPedidoDataHora = getdate()," +
						" PagtoRegistradoNoPedido_id_pedido_pagamento = @PagtoRegistradoNoPedido_id_pedido_pagamento," +
						" PagtoRegistradoNoPedidoStPagtoAnterior = @PagtoRegistradoNoPedidoStPagtoAnterior," +
						" PagtoRegistradoNoPedidoStPagtoNovo = @PagtoRegistradoNoPedidoStPagtoNovo," +
						" AnaliseCreditoStatusAnterior = @AnaliseCreditoStatusAnterior," +
						" AnaliseCreditoStatusNovo = @AnaliseCreditoStatusNovo" +
					" WHERE" +
						" (Id = @Id)";
			cmUpdateWebhookComplementarPagtoRegPedido = BD.criaSqlCommand();
			cmUpdateWebhookComplementarPagtoRegPedido.CommandText = strSql;
			cmUpdateWebhookComplementarPagtoRegPedido.Parameters.Add("@Id", SqlDbType.Int);
			cmUpdateWebhookComplementarPagtoRegPedido.Parameters.Add("@PagtoRegistradoNoPedidoStatus", SqlDbType.TinyInt);
			cmUpdateWebhookComplementarPagtoRegPedido.Parameters.Add("@PagtoRegistradoNoPedidoTipoOperacao", SqlDbType.VarChar, 3);
			cmUpdateWebhookComplementarPagtoRegPedido.Parameters.Add("@PagtoRegistradoNoPedido_id_pedido_pagamento", SqlDbType.VarChar, 12);
			cmUpdateWebhookComplementarPagtoRegPedido.Parameters.Add("@PagtoRegistradoNoPedidoStPagtoAnterior", SqlDbType.VarChar, 1);
			cmUpdateWebhookComplementarPagtoRegPedido.Parameters.Add("@PagtoRegistradoNoPedidoStPagtoNovo", SqlDbType.VarChar, 1);
			cmUpdateWebhookComplementarPagtoRegPedido.Parameters.Add("@AnaliseCreditoStatusAnterior", SqlDbType.SmallInt);
			cmUpdateWebhookComplementarPagtoRegPedido.Parameters.Add("@AnaliseCreditoStatusNovo", SqlDbType.SmallInt);
			cmUpdateWebhookComplementarPagtoRegPedido.Prepare();
			#endregion

			#region [ cmUpdateWebhookEmailEnviadoStatusSucesso ]
			strSql = "UPDATE t_BRASPAG_WEBHOOK SET" +
						" EmailEnviadoStatus = @EmailEnviadoStatus," +
						" EmailEnviadoDataHora = getdate()" +
					" WHERE" +
						" (Id = @Id)";
			cmUpdateWebhookEmailEnviadoStatusSucesso = BD.criaSqlCommand();
			cmUpdateWebhookEmailEnviadoStatusSucesso.CommandText = strSql;
			cmUpdateWebhookEmailEnviadoStatusSucesso.Parameters.Add("@Id", SqlDbType.Int);
			cmUpdateWebhookEmailEnviadoStatusSucesso.Parameters.Add("@EmailEnviadoStatus", SqlDbType.TinyInt);
			cmUpdateWebhookEmailEnviadoStatusSucesso.Prepare();
			#endregion

			#region [ cmUpdateWebhookEmailEnviadoStatusFalha ]
			strSql = "UPDATE t_BRASPAG_WEBHOOK SET" +
						" EmailEnviadoStatus = @EmailEnviadoStatus," +
						" EmailEnviadoDataHora = getdate()," +
						" MsgErro = @MsgErro" +
					" WHERE" +
						" (Id = @Id)";
			cmUpdateWebhookEmailEnviadoStatusFalha = BD.criaSqlCommand();
			cmUpdateWebhookEmailEnviadoStatusFalha.CommandText = strSql;
			cmUpdateWebhookEmailEnviadoStatusFalha.Parameters.Add("@Id", SqlDbType.Int);
			cmUpdateWebhookEmailEnviadoStatusFalha.Parameters.Add("@EmailEnviadoStatus", SqlDbType.TinyInt);
			cmUpdateWebhookEmailEnviadoStatusFalha.Parameters.Add("@MsgErro", SqlDbType.VarChar, -1); // varchar(max)
			cmUpdateWebhookEmailEnviadoStatusFalha.Prepare();
			#endregion

			#region [ cmUpdateWebhookProcessamentoErpStatusSucesso ]
			strSql = "UPDATE t_BRASPAG_WEBHOOK SET" +
						" ProcessamentoErpStatus = @ProcessamentoErpStatus," +
						" ProcessamentoErpDataHora = getdate()" +
					" WHERE" +
						" (Id = @Id)";
			cmUpdateWebhookProcessamentoErpStatusSucesso = BD.criaSqlCommand();
			cmUpdateWebhookProcessamentoErpStatusSucesso.CommandText = strSql;
			cmUpdateWebhookProcessamentoErpStatusSucesso.Parameters.Add("@Id", SqlDbType.Int);
			cmUpdateWebhookProcessamentoErpStatusSucesso.Parameters.Add("@ProcessamentoErpStatus", SqlDbType.Int);
			cmUpdateWebhookProcessamentoErpStatusSucesso.Prepare();
			#endregion

			#region [ cmUpdateWebhookProcessamentoErpStatusFalha ]
			strSql = "UPDATE t_BRASPAG_WEBHOOK SET" +
						" ProcessamentoErpStatus = @ProcessamentoErpStatus," +
						" ProcessamentoErpDataHora = getdate()," +
						" MsgErro = @MsgErro" +
					" WHERE" +
						" (Id = @Id)";
			cmUpdateWebhookProcessamentoErpStatusFalha = BD.criaSqlCommand();
			cmUpdateWebhookProcessamentoErpStatusFalha.CommandText = strSql;
			cmUpdateWebhookProcessamentoErpStatusFalha.Parameters.Add("@Id", SqlDbType.Int);
			cmUpdateWebhookProcessamentoErpStatusFalha.Parameters.Add("@ProcessamentoErpStatus", SqlDbType.Int);
			cmUpdateWebhookProcessamentoErpStatusFalha.Parameters.Add("@MsgErro", SqlDbType.VarChar, -1); // varchar(max)
			cmUpdateWebhookProcessamentoErpStatusFalha.Prepare();
			#endregion
		}
		#endregion

		#region [ contagemRequisicoesPagByCampoOrderId ]
		public static int contagemRequisicoesPagByCampoOrderId(string OrderId, out string msg_erro)
		{
			#region [ Declarações ]
			int intContagem;
			string strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
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

				strSql = "SELECT" +
							" Count(*) AS qtde" +
						" FROM t_PAGTO_GW_PAG" +
							" INNER JOIN t_PAGTO_GW_PAG_PAYMENT ON (t_PAGTO_GW_PAG.id = t_PAGTO_GW_PAG_PAYMENT.id_pagto_gw_pag)" +
						" WHERE" +
							" (req_OrderData_OrderId = '" + OrderId + "')";
				cmCommand.CommandText = strSql;
				intContagem = (int)cmCommand.ExecuteScalar();
				return intContagem;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				return 0;
			}
		}
		#endregion

		#region [ contagemTransacoesPaymentByCampoOrderId ]
		public static int contagemTransacoesPaymentByCampoOrderId(string OrderId, out string msg_erro)
		{
			#region [ Declarações ]
			int intContagem;
			string strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
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

				strSql = "SELECT" +
							" Count(*) AS qtde" +
						" FROM t_PAGTO_GW_PAG_PAYMENT" +
						" WHERE" +
							" (id_pagto_gw_pag IN (SELECT id FROM t_PAGTO_GW_PAG WHERE (req_OrderData_OrderId = '" + OrderId + "')))";
				cmCommand.CommandText = strSql;
				intContagem = (int)cmCommand.ExecuteScalar();
				return intContagem;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();
				return 0;
			}
		}
		#endregion

		#region [ isPedidoERPDesteAmbiente ]
		/// <summary>
		/// Analisa se o número do pedido está no formato usado no sistema (ERP)
		/// Formato usado no sistema: 999999X ou 999999X-A
		/// </summary>
		/// <param name="orderID">Número do pedido a ser analisado</param>
		/// <param name="merchantId">Chave do MerchantId usada nas requisições Braspag</param>
		/// <returns>
		/// true = pedido no formato usado no sistema (ERP)
		/// false = pedido fora do padrão do sistema (ERP)
		/// </returns>
		public static bool isPedidoERPDesteAmbiente(string orderID, string merchantId, out string pedidoERP)
		{
			#region [ Declarações ]
			bool blnTemLetra = false;
			int qtdeDigitosPrefixo = 0;
			string strSql;
			string strPedido = "";
			SqlCommand cmCommand;
			SqlDataReader dr;
			#endregion

			pedidoERP = "";

			// Importante: É importante lembrar que o nº do pedido enviado p/ a Braspag pode conter um sufixo, caso tenham sido enviados mais do que uma requisição de autorização.
			// ===========  Esta situação pode ocorrer no caso de várias tentativas devido a falhas ou autorizações negadas, sendo que a cada tentativa o sufixo é incrementado.
			if (orderID == null) return false;
			if (orderID.Trim().Length == 0) return false;
			for (int i = 0; i < orderID.Length; i++)
			{
				if (Global.isLetra(orderID[i]))
				{
					blnTemLetra = true;
					break;
				}

				if (Global.isDigit(orderID[i]))
				{
					qtdeDigitosPrefixo++;
				}
				else
				{
					break;
				}
			}

			// Numeração não é de um pedido do ERP
			if ((qtdeDigitosPrefixo < Global.Cte.Etc.TAM_MIN_NUM_PEDIDO) || (qtdeDigitosPrefixo > (Global.Cte.Etc.TAM_MIN_NUM_PEDIDO + 1)) || (!blnTemLetra)) return false;

			// Verifica se o pedido está cadastrado neste ambiente
			cmCommand = BD.criaSqlCommand();
			strSql = "SELECT TOP 1 " +
						"pedido" +
					" FROM t_PAGTO_GW_PAG" +
					" WHERE" +
						" (req_OrderData_OrderId = '" + orderID + "')" +
						" AND (req_OrderData_MerchantId = '" + merchantId + "')" +
					" ORDER BY" +
						" id DESC";
			cmCommand.CommandText = strSql;
			dr = cmCommand.ExecuteReader();
			try
			{
				if (dr.Read())
				{
					if (!Convert.IsDBNull(dr["pedido"])) strPedido = dr["pedido"].ToString();
					if (strPedido.Length > 0)
					{
						pedidoERP = strPedido;
						return true;
					}
					else
					{
						return false;
					}
				}
				else
				{
					return false;
				}
			}
			finally
			{
				dr.Close();
			}
		}
		#endregion

		#region [ inserePagOpComplementar ]
		public static bool inserePagOpComplementar(BraspagPagOpComplementar op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.inserePagOpComplementar()";
			bool blnGerouNsu;
			int idPagtoGwPagOpCompl = 0;
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				if (op.id == 0)
				{
					blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_OP_COMPLEMENTAR, out idPagtoGwPagOpCompl, out msg_erro);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_OP_COMPLEMENTAR + "\n" + msg_erro;
						return false;
					}
					op.id = idPagtoGwPagOpCompl;
				}
				else
				{
					// O NSU já foi gerado anteriormente na rotina chamadora
					idPagtoGwPagOpCompl = op.id;
				}

				#region [ Preenche o valor dos parâmetros ]
				cmInsertPagOpComplementar.Parameters["@id"].Value = op.id;
				cmInsertPagOpComplementar.Parameters["@id_pagto_gw_pag"].Value = op.id_pagto_gw_pag;
				cmInsertPagOpComplementar.Parameters["@id_pagto_gw_pag_payment"].Value = op.id_pagto_gw_pag_payment;
				cmInsertPagOpComplementar.Parameters["@usuario"].Value = op.usuario;
				cmInsertPagOpComplementar.Parameters["@operacao"].Value = op.operacao;
				cmInsertPagOpComplementar.Parameters["@req_RequestId"].Value = op.req_RequestId;
				cmInsertPagOpComplementar.Parameters["@req_Version"].Value = op.req_Version;
				cmInsertPagOpComplementar.Parameters["@req_MerchantId"].Value = op.req_MerchantId;
				cmInsertPagOpComplementar.Parameters["@req_BraspagTransactionId"].Value = op.req_BraspagTransactionId;
				cmInsertPagOpComplementar.Parameters["@req_OrderId"].Value = (op.req_OrderId == null ? "" : op.req_OrderId);
				cmInsertPagOpComplementar.Parameters["@req_Amount"].Value = (op.req_Amount == null ? "" : op.req_Amount);
				cmInsertPagOpComplementar.Parameters["@req_ServiceTaxAmount"].Value = (op.req_ServiceTaxAmount == null ? "" : op.req_ServiceTaxAmount);
				#endregion

				#region [ Tenta inserir o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmInsertPagOpComplementar);
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
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion
				}
				#endregion

				#region [ Gravou o registro? ]
				if (intRetorno == 0)
				{
					msg_erro = "Falha ao tentar gravar o registro na tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_OP_COMPLEMENTAR + "\n" + msg_erro;
					return false;
				}
				#endregion

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Sucesso na gravação dos dados (t_PAGTO_GW_PAG_OP_COMPLEMENTAR.id=" + op.id.ToString() + ")");

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.Message;

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ inserePagOpComplementarXml ]
		public static bool inserePagOpComplementarXml(BraspagPagOpComplementarXml op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.inserePagOpComplementarXml()";
			bool blnGerouNsu;
			int idPagtoGwPagOpComplXml = 0;
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				if (op.id == 0)
				{
					blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML, out idPagtoGwPagOpComplXml, out msg_erro);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML + "\n" + msg_erro;
						return false;
					}
					op.id = idPagtoGwPagOpComplXml;
				}
				else
				{
					// O NSU já foi gerado anteriormente na rotina chamadora
					idPagtoGwPagOpComplXml = op.id;
				}

				#region [ Preenche o valor dos parâmetros ]
				cmInsertPagOpComplementarXml.Parameters["@id"].Value = op.id;
				cmInsertPagOpComplementarXml.Parameters["@id_pagto_gw_pag_op_complementar"].Value = op.id_pagto_gw_pag_op_complementar;
				cmInsertPagOpComplementarXml.Parameters["@tipo_transacao"].Value = op.tipo_transacao;
				cmInsertPagOpComplementarXml.Parameters["@fluxo_xml"].Value = op.fluxo_xml;
				cmInsertPagOpComplementarXml.Parameters["@xml"].Value = op.xml;
				#endregion

				#region [ Tenta inserir o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmInsertPagOpComplementarXml);
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
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion
				}
				#endregion

				#region [ Gravou o registro? ]
				if (intRetorno == 0)
				{
					msg_erro = "Falha ao tentar gravar o registro na tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML + "\n" + msg_erro;
					return false;
				}
				#endregion

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Sucesso na gravação dos dados (t_PAGTO_GW_PAG_OP_COMPLEMENTAR_XML.id=" + op.id.ToString() + ")");

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.Message;

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ getBraspagPagById ]
		public static BraspagPag getBraspagPagById(int id, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.getBraspagPagById()";
			string strSql = "";
			string msg_erro_aux;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			BraspagPag pag = new BraspagPag();
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

				strSql = "SELECT * FROM t_PAGTO_GW_PAG WHERE (id = " + id.ToString() + ")";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count == 0) return null;

				row = dtbResultado.Rows[0];

				pag.id = BD.readToInt(row["id"]);
				pag.data = BD.readToDateTime(row["data"]);
				pag.data_hora = BD.readToDateTime(row["data_hora"]);
				pag.owner = BD.readToInt(row["owner"]);
				pag.usuario = BD.readToString(row["usuario"]);
				pag.loja = BD.readToString(row["loja"]);
				pag.id_cliente = BD.readToString(row["id_cliente"]);
				pag.pedido = BD.readToString(row["pedido"]);
				pag.pedido_com_sufixo_nsu = BD.readToString(row["pedido_com_sufixo_nsu"]);
				pag.valor_pedido = BD.readToDecimal(row["valor_pedido"]);
				pag.operacao = BD.readToString(row["operacao"]);
				pag.executado_pelo_cliente_status = BD.readToByte(row["executado_pelo_cliente_status"]);
				pag.origem_endereco_IP = BD.readToString(row["origem_endereco_IP"]);
				pag.FingerPrint_SessionID = BD.readToString(row["FingerPrint_SessionID"]);
				pag.trx_TX_data = BD.readToDateTime(row["trx_TX_data"]);
				pag.trx_TX_data_hora = BD.readToDateTime(row["trx_TX_data_hora"]);
				pag.trx_RX_status = BD.readToByte(row["trx_RX_status"]);
				pag.trx_RX_data = BD.readToDateTime(row["trx_RX_data"]);
				pag.trx_RX_data_hora = BD.readToDateTime(row["trx_RX_data_hora"]);
				pag.trx_RX_vazio_status = BD.readToByte(row["trx_RX_vazio_status"]);
				pag.trx_erro_status = BD.readToByte(row["trx_erro_status"]);
				pag.trx_erro_codigo = BD.readToString(row["trx_erro_codigo"]);
				pag.trx_erro_mensagem = BD.readToString(row["trx_erro_mensagem"]);
				pag.trx_TX_id_pagto_gw_pag_xml = BD.readToInt(row["trx_TX_id_pagto_gw_pag_xml"]);
				pag.trx_RX_id_pagto_gw_pag_xml = BD.readToInt(row["trx_RX_id_pagto_gw_pag_xml"]);
				pag.req_RequestId = BD.readToString(row["req_RequestId"]);
				pag.req_Version = BD.readToString(row["req_Version"]);
				pag.req_OrderData_MerchantId = BD.readToString(row["req_OrderData_MerchantId"]);
				pag.req_OrderData_OrderId = BD.readToString(row["req_OrderData_OrderId"]);
				pag.req_CustomerData_CustomerIdentity = BD.readToString(row["req_CustomerData_CustomerIdentity"]);
				pag.req_CustomerData_CustomerName = BD.readToString(row["req_CustomerData_CustomerName"]);
				pag.resp_CorrelationId = BD.readToString(row["resp_CorrelationId"]);
				pag.resp_Success = BD.readToString(row["resp_Success"]);
				pag.resp_OrderData_OrderId = BD.readToString(row["resp_OrderData_OrderId"]);
				pag.resp_OrderData_BraspagOrderId = BD.readToString(row["resp_OrderData_BraspagOrderId"]);
				pag.recibo_url_css = BD.readToString(row["recibo_url_css"]);
				pag.recibo_html = BD.readToString(row["recibo_html"]);
				pag.msg_alerta_tela = BD.readToString(row["msg_alerta_tela"]);
				pag.SessionCtrlInfo = BD.readToString(row["SessionCtrlInfo"]);

				return pag;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(pag);
				svcLog.complemento_2 = strSql;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ getBraspagPagPaymentById ]
		public static BraspagPagPayment getBraspagPagPaymentById(int id, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.getBraspagPagPaymentById()";
			string strSql = "";
			string msg_erro_aux;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			BraspagPagPayment payment = new BraspagPagPayment();
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

				strSql = "SELECT * FROM t_PAGTO_GW_PAG_PAYMENT WHERE (id = " + id.ToString() + ")";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count == 0) return null;

				row = dtbResultado.Rows[0];

				payment.id = BD.readToInt(row["id"]);
				payment.id_pagto_gw_pag = BD.readToInt(row["id_pagto_gw_pag"]);
				payment.ordem = BD.readToInt(row["ordem"]);
				payment.st_enviado_analise_AF = BD.readToByte(row["st_enviado_analise_AF"]);
				payment.id_pagto_gw_af = BD.readToInt(row["id_pagto_gw_af"]);
				payment.bandeira = BD.readToString(row["bandeira"]);
				payment.valor_transacao = BD.readToDecimal(row["valor_transacao"]);
				payment.checkout_opcao_parcelamento = BD.readToString(row["checkout_opcao_parcelamento"]);
				payment.checkout_titular_nome = BD.readToString(row["checkout_titular_nome"]);
				payment.checkout_titular_cpf_cnpj = BD.readToString(row["checkout_titular_cpf_cnpj"]);
				payment.checkout_cartao_numero = BD.readToString(row["checkout_cartao_numero"]);
				payment.checkout_cartao_validade_mes = BD.readToString(row["checkout_cartao_validade_mes"]);
				payment.checkout_cartao_validade_ano = BD.readToString(row["checkout_cartao_validade_ano"]);
				payment.checkout_cartao_codigo_seguranca = BD.readToString(row["checkout_cartao_codigo_seguranca"]);
				payment.checkout_cartao_proprio = BD.readToString(row["checkout_cartao_proprio"]);
				payment.checkout_fatura_end_logradouro = BD.readToString(row["checkout_fatura_end_logradouro"]);
				payment.checkout_fatura_end_numero = BD.readToString(row["checkout_fatura_end_numero"]);
				payment.checkout_fatura_end_complemento = BD.readToString(row["checkout_fatura_end_complemento"]);
				payment.checkout_fatura_end_cidade = BD.readToString(row["checkout_fatura_end_cidade"]);
				payment.checkout_fatura_end_uf = BD.readToString(row["checkout_fatura_end_uf"]);
				payment.checkout_fatura_end_cep = BD.readToString(row["checkout_fatura_end_cep"]);
				payment.checkout_fatura_tel_pais = BD.readToString(row["checkout_fatura_tel_pais"]);
				payment.checkout_fatura_tel_ddd = BD.readToString(row["checkout_fatura_tel_ddd"]);
				payment.checkout_fatura_tel_numero = BD.readToString(row["checkout_fatura_tel_numero"]);
				payment.checkout_email = BD.readToString(row["checkout_email"]);
				payment.prim_GlobalStatus = BD.readToString(row["prim_GlobalStatus"]);
				payment.prim_atualizacao_data_hora = BD.readToDateTime(row["prim_atualizacao_data_hora"]);
				payment.prim_atualizacao_usuario = BD.readToString(row["prim_atualizacao_usuario"]);
				payment.ult_GlobalStatus = BD.readToString(row["ult_GlobalStatus"]);
				payment.ult_atualizacao_data_hora = BD.readToDateTime(row["ult_atualizacao_data_hora"]);
				payment.ult_atualizacao_usuario = BD.readToString(row["ult_atualizacao_usuario"]);
				payment.ult_id_pagto_gw_pag_payment_op_complementar = BD.readToInt(row["ult_id_pagto_gw_pag_payment_op_complementar"]);
				payment.resp_AuthorizedDate = BD.readToDateTime(row["resp_AuthorizedDate"]);
				payment.resp_CapturedDate = BD.readToDateTime(row["resp_CapturedDate"]);
				payment.resp_VoidedDate = BD.readToDateTime(row["resp_VoidedDate"]);
				payment.tratado_manual_status = BD.readToByte(row["tratado_manual_status"]);
				payment.tratado_manual_usuario = BD.readToString(row["tratado_manual_usuario"]);
				payment.tratado_manual_data = BD.readToDateTime(row["tratado_manual_data"]);
				payment.tratado_manual_data_hora = BD.readToDateTime(row["tratado_manual_data_hora"]);
				payment.tratado_manual_obs = BD.readToString(row["tratado_manual_obs"]);
				payment.tratado_manual_ult_atualiz_usuario = BD.readToString(row["tratado_manual_ult_atualiz_usuario"]);
				payment.tratado_manual_ult_atualiz_data = BD.readToDateTime(row["tratado_manual_ult_atualiz_data"]);
				payment.tratado_manual_ult_atualiz_data_hora = BD.readToDateTime(row["tratado_manual_ult_atualiz_data_hora"]);
				payment.req_PaymentDataRequest_PaymentMethod = BD.readToString(row["req_PaymentDataRequest_PaymentMethod"]);
				payment.req_PaymentDataRequest_Amount = BD.readToString(row["req_PaymentDataRequest_Amount"]);
				payment.req_PaymentDataRequest_Currency = BD.readToString(row["req_PaymentDataRequest_Currency"]);
				payment.req_PaymentDataRequest_Country = BD.readToString(row["req_PaymentDataRequest_Country"]);
				payment.req_PaymentDataRequest_ServiceTaxAmount = BD.readToString(row["req_PaymentDataRequest_ServiceTaxAmount"]);
				payment.req_PaymentDataRequest_NumberOfPayments = BD.readToString(row["req_PaymentDataRequest_NumberOfPayments"]);
				payment.req_PaymentDataRequest_PaymentPlan = BD.readToString(row["req_PaymentDataRequest_PaymentPlan"]);
				payment.req_PaymentDataRequest_TransactionType = BD.readToString(row["req_PaymentDataRequest_TransactionType"]);
				payment.req_PaymentDataRequest_CardHolder = BD.readToString(row["req_PaymentDataRequest_CardHolder"]);
				payment.req_PaymentDataRequest_CardNumber = BD.readToString(row["req_PaymentDataRequest_CardNumber"]);
				payment.req_PaymentDataRequest_CardSecurityCode = BD.readToString(row["req_PaymentDataRequest_CardSecurityCode"]);
				payment.req_PaymentDataRequest_CardExpirationDate = BD.readToString(row["req_PaymentDataRequest_CardExpirationDate"]);
				payment.resp_PaymentDataResponse_BraspagTransactionId = BD.readToString(row["resp_PaymentDataResponse_BraspagTransactionId"]);
				payment.resp_PaymentDataResponse_PaymentMethod = BD.readToString(row["resp_PaymentDataResponse_PaymentMethod"]);
				payment.resp_PaymentDataResponse_Amount = BD.readToString(row["resp_PaymentDataResponse_Amount"]);
				payment.resp_PaymentDataResponse_AcquirerTransactionId = BD.readToString(row["resp_PaymentDataResponse_AcquirerTransactionId"]);
				payment.resp_PaymentDataResponse_AuthorizationCode = BD.readToString(row["resp_PaymentDataResponse_AuthorizationCode"]);
				payment.resp_PaymentDataResponse_CreditCardToken = BD.readToString(row["resp_PaymentDataResponse_CreditCardToken"]);
				payment.resp_PaymentDataResponse_ProofOfSale = BD.readToString(row["resp_PaymentDataResponse_ProofOfSale"]);
				payment.resp_PaymentDataResponse_ReturnCode = BD.readToString(row["resp_PaymentDataResponse_ReturnCode"]);
				payment.resp_PaymentDataResponse_ReturnMessage = BD.readToString(row["resp_PaymentDataResponse_ReturnMessage"]);
				payment.resp_PaymentDataResponse_Status = BD.readToString(row["resp_PaymentDataResponse_Status"]);
				payment.captura_confirmada_status = BD.readToByte(row["captura_confirmada_status"]);
				payment.captura_confirmada_data = BD.readToDateTime(row["captura_confirmada_data"]);
				payment.captura_confirmada_data_hora = BD.readToDateTime(row["captura_confirmada_data_hora"]);
				payment.captura_confirmada_usuario = BD.readToString(row["captura_confirmada_usuario"]);
				payment.voided_status = BD.readToByte(row["voided_status"]);
				payment.voided_data = BD.readToDateTime(row["voided_data"]);
				payment.voided_data_hora = BD.readToDateTime(row["voided_data_hora"]);
				payment.voided_usuario = BD.readToString(row["voided_usuario"]);
				payment.refunded_status = BD.readToByte(row["refunded_status"]);
				payment.refunded_data = BD.readToDateTime(row["refunded_data"]);
				payment.refunded_data_hora = BD.readToDateTime(row["refunded_data_hora"]);
				payment.refunded_usuario = BD.readToString(row["refunded_usuario"]);
				payment.refund_pending_status = BD.readToByte(row["refund_pending_status"]);
				payment.refund_pending_data = BD.readToDateTime(row["refund_pending_data"]);
				payment.refund_pending_data_hora = BD.readToDateTime(row["refund_pending_data_hora"]);
				payment.refund_pending_usuario = BD.readToString(row["refund_pending_usuario"]);
				payment.refund_pending_confirmado_status = BD.readToByte(row["refund_pending_confirmado_status"]);
				payment.refund_pending_confirmado_data = BD.readToDateTime(row["refund_pending_confirmado_data"]);
				payment.refund_pending_confirmado_data_hora = BD.readToDateTime(row["refund_pending_confirmado_data_hora"]);
				payment.refund_pending_confirmado_usuario = BD.readToString(row["refund_pending_confirmado_usuario"]);
				payment.refund_pending_falha_status = BD.readToByte(row["refund_pending_falha_status"]);
				payment.refund_pending_falha_data = BD.readToDateTime(row["refund_pending_falha_data"]);
				payment.refund_pending_falha_data_hora = BD.readToDateTime(row["refund_pending_falha_data_hora"]);
				payment.refund_pending_falha_motivo = BD.readToString(row["refund_pending_falha_motivo"]);
				payment.captura_confirmada_erro_status = BD.readToByte(row["captura_confirmada_erro_status"]);
				payment.captura_confirmada_erro_data = BD.readToDateTime(row["captura_confirmada_erro_data"]);
				payment.captura_confirmada_erro_data_hora = BD.readToDateTime(row["captura_confirmada_erro_data_hora"]);
				payment.captura_confirmada_erro_mensagem = BD.readToString(row["captura_confirmada_erro_mensagem"]);
				payment.voided_erro_status = BD.readToByte(row["voided_erro_status"]);
				payment.voided_erro_data = BD.readToDateTime(row["voided_erro_data"]);
				payment.voided_erro_data_hora = BD.readToDateTime(row["voided_erro_data_hora"]);
				payment.voided_erro_mensagem = BD.readToString(row["voided_erro_mensagem"]);
				payment.refunded_erro_status = BD.readToByte(row["refunded_erro_status"]);
				payment.refunded_erro_data = BD.readToDateTime(row["refunded_erro_data"]);
				payment.refunded_erro_data_hora = BD.readToDateTime(row["refunded_erro_data_hora"]);
				payment.refunded_erro_mensagem = BD.readToString(row["refunded_erro_mensagem"]);
				payment.pedido_hist_pagto_gravado_status = BD.readToByte(row["pedido_hist_pagto_gravado_status"]);
				payment.pedido_hist_pagto_gravado_data = BD.readToDateTime(row["pedido_hist_pagto_gravado_data"]);
				payment.pedido_hist_pagto_gravado_data_hora = BD.readToDateTime(row["pedido_hist_pagto_gravado_data_hora"]);
				payment.pagto_registrado_no_pedido_status = BD.readToByte(row["pagto_registrado_no_pedido_status"]);
				payment.pagto_registrado_no_pedido_tipo_operacao = BD.readToString(row["pagto_registrado_no_pedido_tipo_operacao"]);
				payment.pagto_registrado_no_pedido_data = BD.readToDateTime(row["pagto_registrado_no_pedido_data"]);
				payment.pagto_registrado_no_pedido_data_hora = BD.readToDateTime(row["pagto_registrado_no_pedido_data_hora"]);
				payment.pagto_registrado_no_pedido_usuario = BD.readToString(row["pagto_registrado_no_pedido_usuario"]);
				payment.pagto_registrado_no_pedido_id_pedido_pagamento = BD.readToString(row["pagto_registrado_no_pedido_id_pedido_pagamento"]);
				payment.pagto_registrado_no_pedido_st_pagto_anterior = BD.readToString(row["pagto_registrado_no_pedido_st_pagto_anterior"]);
				payment.pagto_registrado_no_pedido_st_pagto_novo = BD.readToString(row["pagto_registrado_no_pedido_st_pagto_novo"]);
				payment.st_cancelado_envio_analise_AF = BD.readToByte(row["st_cancelado_envio_analise_AF"]);
				payment.checkout_fatura_end_bairro = BD.readToString(row["checkout_fatura_end_bairro"]);
				payment.st_processamento_AF_finalizado = BD.readToByte(row["st_processamento_AF_finalizado"]);
				payment.dt_hr_processamento_AF_finalizado = BD.readToDateTime(row["dt_hr_processamento_AF_finalizado"]);
				payment.st_processamento_PAG_finalizado = BD.readToByte(row["st_processamento_PAG_finalizado"]);
				payment.dt_hr_processamento_PAG_finalizado = BD.readToDateTime(row["dt_hr_processamento_PAG_finalizado"]);

				return payment;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(payment);
				svcLog.complemento_2 = strSql;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ getBraspagWebhookById ]
		public static BraspagWebhook getBraspagWebhookById(int id, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.getBraspagWebhookById()";
			string strSql = "";
			string msg_erro_aux;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			BraspagWebhook result = new BraspagWebhook();
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

				strSql = "SELECT * FROM t_BRASPAG_WEBHOOK WHERE (Id = " + id.ToString() + ")";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count == 0) return null;

				row = dtbResultado.Rows[0];
				result.Id = BD.readToInt(row["Id"]);
				result.DataCadastro = BD.readToDateTime(row["DataCadastro"]);
				result.DataHoraCadastro = BD.readToDateTime(row["DataHoraCadastro"]);
				result.Empresa = BD.readToString(row["Empresa"]);
				result.NumPedido = BD.readToString(row["NumPedido"]);
				result.Status = BD.readToString(row["Status"]);
				result.CODPAGAMENTO = BD.readToString(row["CODPAGAMENTO"]);
				result.BraspagDadosComplementaresQueryStatus = BD.readToByte(row["BraspagDadosComplementaresQueryStatus"]);
				result.BraspagDadosComplementaresQueryDataHora = BD.readToDateTime(row["BraspagDadosComplementaresQueryDataHora"]);
				result.EmailEnviadoStatus = BD.readToByte(row["EmailEnviadoStatus"]);
				result.EmailEnviadoDataHora = BD.readToDateTime(row["EmailEnviadoDataHora"]);
				result.ProcessamentoErpStatus = BD.readToInt(row["ProcessamentoErpStatus"]);
				result.ProcessamentoErpDataHora = BD.readToDateTime(row["ProcessamentoErpDataHora"]);
				result.BraspagDadosComplementaresQueryTentativas = BD.readToInt(row["BraspagDadosComplementaresQueryTentativas"]);
				result.BraspagDadosComplementaresQueryDtHrUltTentativa = BD.readToDateTime(row["BraspagDadosComplementaresQueryDtHrUltTentativa"]);
				result.MsgErro = BD.readToString(row["MsgErro"]);
				result.MsgErroTemporario = BD.readToString(row["MsgErroTemporario"]);

				return result;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + id.ToString();
				svcLog.complemento_2 = Global.serializaObjectToXml(result);
				svcLog.complemento_3 = strSql;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ getBraspagWebhookComplementarById ]
		public static BraspagWebhookComplementar getBraspagWebhookComplementarById(int id, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.getBraspagWebhookComplementarById()";
			string strSql = "";
			string msg_erro_aux;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			BraspagWebhookComplementar result = new BraspagWebhookComplementar();
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

				strSql = "SELECT * FROM t_BRASPAG_WEBHOOK_COMPLEMENTAR WHERE (Id = " + id.ToString() + ")";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count == 0) return null;

				row = dtbResultado.Rows[0];
				result.Id = BD.readToInt(row["Id"]);
				result.id_braspag_webhook = BD.readToInt(row["id_braspag_webhook"]);
				result.Data = BD.readToDateTime(row["Data"]);
				result.DataHora = BD.readToDateTime(row["DataHora"]);
				result.BraspagTransactionId = BD.readToString(row["BraspagTransactionId"]);
				result.BraspagOrderId = BD.readToString(row["BraspagOrderId"]);
				result.PaymentMethod = BD.readToString(row["PaymentMethod"]);
				result.GlobalStatus = BD.readToString(row["GlobalStatus"]);
				result.ReceivedDate = BD.readToDateTime(row["ReceivedDate"]);
				result.CapturedDate = BD.readToDateTime(row["CapturedDate"]);
				result.CustomerName = BD.readToString(row["CustomerName"]);
				result.BoletoExpirationDate = BD.readToDateTime(row["BoletoExpirationDate"]);
				result.Amount = BD.readToString(row["Amount"]);
				result.ValorAmount = BD.readToDecimal(row["ValorAmount"]);
				result.PaidAmount = BD.readToString(row["PaidAmount"]);
				result.ValorPaidAmount = BD.readToDecimal(row["ValorPaidAmount"]);
				result.pedido = BD.readToString(row["pedido"]);
				result.PagtoRegistradoNoPedidoStatus = BD.readToByte(row["PagtoRegistradoNoPedidoStatus"]);
				result.PagtoRegistradoNoPedidoTipoOperacao = BD.readToString(row["PagtoRegistradoNoPedidoTipoOperacao"]);
				result.PagtoRegistradoNoPedidoData = BD.readToDateTime(row["PagtoRegistradoNoPedidoData"]);
				result.PagtoRegistradoNoPedidoDataHora = BD.readToDateTime(row["PagtoRegistradoNoPedidoDataHora"]);
				result.PagtoRegistradoNoPedido_id_pedido_pagamento = BD.readToString(row["PagtoRegistradoNoPedido_id_pedido_pagamento"]);
				result.PagtoRegistradoNoPedidoStPagtoAnterior = BD.readToString(row["PagtoRegistradoNoPedidoStPagtoAnterior"]);
				result.PagtoRegistradoNoPedidoStPagtoNovo = BD.readToString(row["PagtoRegistradoNoPedidoStPagtoNovo"]);
				result.AnaliseCreditoStatusAnterior = BD.readToInt(row["AnaliseCreditoStatusAnterior"]);
				result.AnaliseCreditoStatusNovo = BD.readToInt(row["AnaliseCreditoStatusNovo"]);
				result.PedidoHistPagtoGravadoStatus = BD.readToByte(row["PedidoHistPagtoGravadoStatus"]);
				result.PedidoHistPagtoGravadoData = BD.readToDateTime(row["PedidoHistPagtoGravadoData"]);
				result.PedidoHistPagtoGravadoDataHora = BD.readToDateTime(row["PedidoHistPagtoGravadoDataHora"]);
				result.MsgErro = BD.readToString(row["MsgErro"]);

				return result;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + id.ToString();
				svcLog.complemento_2 = Global.serializaObjectToXml(result);
				svcLog.complemento_3 = strSql;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ updatePagOpComplementarCaptureCreditCardResp ]
		public static bool updatePagOpComplementarCaptureCreditCardResp(BraspagPagOpComplementar op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagOpComplementarCaptureCreditCardResp()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters["@id"].Value = op.id;
				cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters["@trx_RX_status"].Value = op.trx_RX_status;
				cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters["@trx_RX_vazio_status"].Value = op.trx_RX_vazio_status;
				cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters["@st_sucesso"].Value = op.st_sucesso;
				cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters["@resp_BraspagTransactionId"].Value = (op.resp_BraspagTransactionId == null ? "" : op.resp_BraspagTransactionId);
				cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters["@resp_AuthorizationCode"].Value = (op.resp_AuthorizationCode == null ? "" : op.resp_AuthorizationCode);
				cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters["@resp_ProofOfSale"].Value = (op.resp_ProofOfSale == null ? "" : op.resp_ProofOfSale);
				cmUpdatePagOpComplementarCaptureCreditCardResp.Parameters["@resp_Status"].Value = (op.resp_Status == null ? "" : op.resp_Status);
				#endregion

				#region [ Tenta alterar o(s) registro(s), se houver algum ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagOpComplementarCaptureCreditCardResp);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.Message;

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagOpComplementarVoidCreditCardResp ]
		public static bool updatePagOpComplementarVoidCreditCardResp(BraspagPagOpComplementar op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagOpComplementarVoidCreditCardResp()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagOpComplementarVoidCreditCardResp.Parameters["@id"].Value = op.id;
				cmUpdatePagOpComplementarVoidCreditCardResp.Parameters["@trx_RX_status"].Value = op.trx_RX_status;
				cmUpdatePagOpComplementarVoidCreditCardResp.Parameters["@trx_RX_vazio_status"].Value = op.trx_RX_vazio_status;
				cmUpdatePagOpComplementarVoidCreditCardResp.Parameters["@st_sucesso"].Value = op.st_sucesso;
				cmUpdatePagOpComplementarVoidCreditCardResp.Parameters["@resp_BraspagTransactionId"].Value = (op.resp_BraspagTransactionId == null ? "" : op.resp_BraspagTransactionId);
				cmUpdatePagOpComplementarVoidCreditCardResp.Parameters["@resp_AuthorizationCode"].Value = (op.resp_AuthorizationCode == null ? "" : op.resp_AuthorizationCode);
				cmUpdatePagOpComplementarVoidCreditCardResp.Parameters["@resp_ProofOfSale"].Value = (op.resp_ProofOfSale == null ? "" : op.resp_ProofOfSale);
				cmUpdatePagOpComplementarVoidCreditCardResp.Parameters["@resp_Status"].Value = (op.resp_Status == null ? "" : op.resp_Status);
				#endregion

				#region [ Tenta alterar o(s) registro(s), se houver algum ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagOpComplementarVoidCreditCardResp);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.Message;

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagOpComplementarRefundCreditCardResp ]
		public static bool updatePagOpComplementarRefundCreditCardResp(BraspagPagOpComplementar op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagOpComplementarRefundCreditCardResp()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagOpComplementarRefundCreditCardResp.Parameters["@id"].Value = op.id;
				cmUpdatePagOpComplementarRefundCreditCardResp.Parameters["@trx_RX_status"].Value = op.trx_RX_status;
				cmUpdatePagOpComplementarRefundCreditCardResp.Parameters["@trx_RX_vazio_status"].Value = op.trx_RX_vazio_status;
				cmUpdatePagOpComplementarRefundCreditCardResp.Parameters["@st_sucesso"].Value = op.st_sucesso;
				cmUpdatePagOpComplementarRefundCreditCardResp.Parameters["@resp_BraspagTransactionId"].Value = (op.resp_BraspagTransactionId == null ? "" : op.resp_BraspagTransactionId);
				cmUpdatePagOpComplementarRefundCreditCardResp.Parameters["@resp_AuthorizationCode"].Value = (op.resp_AuthorizationCode == null ? "" : op.resp_AuthorizationCode);
				cmUpdatePagOpComplementarRefundCreditCardResp.Parameters["@resp_ProofOfSale"].Value = (op.resp_ProofOfSale == null ? "" : op.resp_ProofOfSale);
				cmUpdatePagOpComplementarRefundCreditCardResp.Parameters["@resp_Status"].Value = (op.resp_Status == null ? "" : op.resp_Status);
				#endregion

				#region [ Tenta alterar o(s) registro(s), se houver algum ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagOpComplementarRefundCreditCardResp);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.Message;

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagOpComplementarGetOrderIdDataResp ]
		public static bool updatePagOpComplementarGetOrderIdDataResp(BraspagPagOpComplementar op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagOpComplementarGetOrderIdDataResp()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagOpComplementarGetOrderIdDataResp.Parameters["@id"].Value = op.id;
				cmUpdatePagOpComplementarGetOrderIdDataResp.Parameters["@trx_RX_status"].Value = op.trx_RX_status;
				cmUpdatePagOpComplementarGetOrderIdDataResp.Parameters["@trx_RX_vazio_status"].Value = op.trx_RX_vazio_status;
				cmUpdatePagOpComplementarGetOrderIdDataResp.Parameters["@st_sucesso"].Value = op.st_sucesso;
				cmUpdatePagOpComplementarGetOrderIdDataResp.Parameters["@resp_BraspagTransactionId"].Value = (op.resp_BraspagTransactionId == null ? "" : op.resp_BraspagTransactionId);
				#endregion

				#region [ Tenta alterar o(s) registro(s), se houver algum ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagOpComplementarGetOrderIdDataResp);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.Message;

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagOpComplementarGetOrderDataResp ]
		public static bool updatePagOpComplementarGetOrderDataResp(BraspagPagOpComplementar op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagOpComplementarGetOrderDataResp()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagOpComplementarGetOrderDataResp.Parameters["@id"].Value = op.id;
				cmUpdatePagOpComplementarGetOrderDataResp.Parameters["@trx_RX_status"].Value = op.trx_RX_status;
				cmUpdatePagOpComplementarGetOrderDataResp.Parameters["@trx_RX_vazio_status"].Value = op.trx_RX_vazio_status;
				cmUpdatePagOpComplementarGetOrderDataResp.Parameters["@st_sucesso"].Value = op.st_sucesso;
				cmUpdatePagOpComplementarGetOrderDataResp.Parameters["@resp_BraspagTransactionId"].Value = (op.resp_BraspagTransactionId == null ? "" : op.resp_BraspagTransactionId);
				#endregion

				#region [ Tenta alterar o(s) registro(s), se houver algum ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagOpComplementarGetOrderDataResp);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.Message;

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagOpComplementarGetTransactionDataResp ]
		public static bool updatePagOpComplementarGetTransactionDataResp(BraspagPagOpComplementar op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagOpComplementarGetTransactionDataResp()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagOpComplementarGetTransactionDataResp.Parameters["@id"].Value = op.id;
				cmUpdatePagOpComplementarGetTransactionDataResp.Parameters["@trx_RX_status"].Value = op.trx_RX_status;
				cmUpdatePagOpComplementarGetTransactionDataResp.Parameters["@trx_RX_vazio_status"].Value = op.trx_RX_vazio_status;
				cmUpdatePagOpComplementarGetTransactionDataResp.Parameters["@st_sucesso"].Value = op.st_sucesso;
				cmUpdatePagOpComplementarGetTransactionDataResp.Parameters["@resp_BraspagTransactionId"].Value = (op.resp_BraspagTransactionId == null ? "" : op.resp_BraspagTransactionId);
				cmUpdatePagOpComplementarGetTransactionDataResp.Parameters["@resp_AuthorizationCode"].Value = (op.resp_AuthorizationCode == null ? "" : op.resp_AuthorizationCode);
				cmUpdatePagOpComplementarGetTransactionDataResp.Parameters["@resp_ProofOfSale"].Value = (op.resp_ProofOfSale == null ? "" : op.resp_ProofOfSale);
				cmUpdatePagOpComplementarGetTransactionDataResp.Parameters["@resp_Status"].Value = (op.resp_Status == null ? "" : op.resp_Status);
				#endregion

				#region [ Tenta alterar o(s) registro(s), se houver algum ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagOpComplementarGetTransactionDataResp);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.Message;

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagPaymentGetTransactionDataResp ]
		public static bool updatePagPaymentGetTransactionDataResp(BraspagUpdatePagPaymentGetTransactionDataResponse op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagPaymentGetTransactionDataResp()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagPaymentGetTransactionDataResp.Parameters["@id"].Value = op.id_pagto_gw_pag_payment;
				cmUpdatePagPaymentGetTransactionDataResp.Parameters["@ult_GlobalStatus"].Value = op.ult_GlobalStatus;
				cmUpdatePagPaymentGetTransactionDataResp.Parameters["@ult_atualizacao_usuario"].Value = op.ult_atualizacao_usuario;
				cmUpdatePagPaymentGetTransactionDataResp.Parameters["@ult_id_pagto_gw_pag_payment_op_complementar"].Value = op.ult_id_pagto_gw_pag_payment_op_complementar;
				cmUpdatePagPaymentGetTransactionDataResp.Parameters["@resp_CapturedDate"].Value = Global.formataDataYyyyMmDdHhMmSsComSeparador(op.resp_CapturedDate);
				cmUpdatePagPaymentGetTransactionDataResp.Parameters["@resp_VoidedDate"].Value = Global.formataDataYyyyMmDdHhMmSsComSeparador(op.resp_VoidedDate);
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagPaymentGetTransactionDataResp);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + " com os dados de resposta do método Braspag GetTransactionData(): " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id = " + op.id_pagto_gw_pag_payment.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagPaymentBraspagTransactionId ]
		public static bool updatePagPaymentBraspagTransactionId(int id, string BraspagTransactionId, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagPaymentBraspagTransactionId()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagPaymentBraspagTransactionId.Parameters["@id"].Value = id;
				cmUpdatePagPaymentBraspagTransactionId.Parameters["@BraspagTransactionId"].Value = (BraspagTransactionId == null ? "" : BraspagTransactionId);
				#endregion

				#region [ Tenta alterar o(s) registro(s), se houver algum ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagPaymentBraspagTransactionId);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = "t_PAGTO_GW_PAG_PAYMENT.id = " + id.ToString() + "; BraspagTransactionId = " + BraspagTransactionId;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.Message;

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = "t_PAGTO_GW_PAG_PAYMENT.id = " + id.ToString() + "; BraspagTransactionId = " + BraspagTransactionId;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagPaymentCaptureCreditCardRespSucesso ]
		public static bool updatePagPaymentCaptureCreditCardRespSucesso(BraspagUpdatePagPaymentCaptureCreditCardTransactionResponseSucesso op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagPaymentCaptureCreditCardRespSucesso()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters["@id"].Value = op.id_pagto_gw_pag_payment;
				cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters["@ult_GlobalStatus"].Value = (op.ult_GlobalStatus == null ? "" : op.ult_GlobalStatus);
				cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters["@ult_atualizacao_usuario"].Value = op.ult_atualizacao_usuario;
				cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters["@ult_id_pagto_gw_pag_payment_op_complementar"].Value = op.ult_id_pagto_gw_pag_payment_op_complementar;
				cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters["@resp_CapturedDate"].Value = (op.resp_CapturedDate == DateTime.MinValue ? "" : Global.formataDataYyyyMmDdHhMmSsComSeparador(op.resp_CapturedDate));
				cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters["@captura_confirmada_status"].Value = op.captura_confirmada_status;
				cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters["@captura_confirmada_data"].Value = (op.captura_confirmada_data == DateTime.MinValue ? "" : Global.formataDataYyyyMmDdComSeparador(op.captura_confirmada_data));
				cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters["@captura_confirmada_data_hora"].Value = (op.captura_confirmada_data_hora == DateTime.MinValue ? "" : Global.formataDataYyyyMmDdHhMmSsComSeparador(op.captura_confirmada_data_hora));
				cmUpdatePagPaymentCaptureCreditCardRespSucesso.Parameters["@captura_confirmada_usuario"].Value = op.captura_confirmada_usuario ?? "";
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagPaymentCaptureCreditCardRespSucesso);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + " com os dados de resposta do método Braspag CaptureCreditCardTransaction(): " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id = " + op.id_pagto_gw_pag_payment.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagPaymentCaptureCreditCardRespFalha ]
		public static bool updatePagPaymentCaptureCreditCardRespFalha(BraspagUpdatePagPaymentCaptureCreditCardTransactionResponseFalha op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagPaymentCaptureCreditCardRespFalha()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagPaymentCaptureCreditCardRespFalha.Parameters["@id"].Value = op.id_pagto_gw_pag_payment;
				cmUpdatePagPaymentCaptureCreditCardRespFalha.Parameters["@ult_atualizacao_usuario"].Value = op.ult_atualizacao_usuario;
				cmUpdatePagPaymentCaptureCreditCardRespFalha.Parameters["@ult_id_pagto_gw_pag_payment_op_complementar"].Value = op.ult_id_pagto_gw_pag_payment_op_complementar;
				cmUpdatePagPaymentCaptureCreditCardRespFalha.Parameters["@captura_confirmada_erro_status"].Value = op.captura_confirmada_erro_status;
				cmUpdatePagPaymentCaptureCreditCardRespFalha.Parameters["@captura_confirmada_erro_data"].Value = (op.captura_confirmada_erro_data == DateTime.MinValue ? "" : Global.formataDataYyyyMmDdComSeparador(op.captura_confirmada_erro_data));
				cmUpdatePagPaymentCaptureCreditCardRespFalha.Parameters["@captura_confirmada_erro_data_hora"].Value = (op.captura_confirmada_erro_data_hora == DateTime.MinValue ? "" : Global.formataDataYyyyMmDdHhMmSsComSeparador(op.captura_confirmada_erro_data_hora));
				cmUpdatePagPaymentCaptureCreditCardRespFalha.Parameters["@captura_confirmada_erro_mensagem"].Value = (op.captura_confirmada_erro_mensagem == null ? "" : op.captura_confirmada_erro_mensagem);
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagPaymentCaptureCreditCardRespFalha);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + " com os dados de resposta do método Braspag CaptureCreditCardTransaction(): " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id = " + op.id_pagto_gw_pag_payment.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagPaymentVoidCreditCardRespSucesso ]
		public static bool updatePagPaymentVoidCreditCardRespSucesso(BraspagUpdatePagPaymentVoidCreditCardTransactionResponseSucesso op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagPaymentVoidCreditCardRespSucesso()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters["@id"].Value = op.id_pagto_gw_pag_payment;
				cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters["@ult_GlobalStatus"].Value = (op.ult_GlobalStatus == null ? "" : op.ult_GlobalStatus);
				cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters["@ult_atualizacao_usuario"].Value = op.ult_atualizacao_usuario;
				cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters["@ult_id_pagto_gw_pag_payment_op_complementar"].Value = op.ult_id_pagto_gw_pag_payment_op_complementar;
				cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters["@resp_VoidedDate"].Value = (op.resp_VoidedDate == DateTime.MinValue ? "" : Global.formataDataYyyyMmDdHhMmSsComSeparador(op.resp_VoidedDate));
				cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters["@voided_status"].Value = op.voided_status;
				cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters["@voided_data"].Value = (op.voided_data == DateTime.MinValue ? "" : Global.formataDataYyyyMmDdComSeparador(op.voided_data));
				cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters["@voided_data_hora"].Value = (op.voided_data_hora == DateTime.MinValue ? "" : Global.formataDataYyyyMmDdHhMmSsComSeparador(op.voided_data_hora));
				cmUpdatePagPaymentVoidCreditCardRespSucesso.Parameters["@voided_usuario"].Value = op.voided_usuario ?? "";
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagPaymentVoidCreditCardRespSucesso);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + " com os dados de resposta do método Braspag VoidCreditCardTransaction(): " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id = " + op.id_pagto_gw_pag_payment.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagPaymentVoidCreditCardRespFalha ]
		public static bool updatePagPaymentVoidCreditCardRespFalha(BraspagUpdatePagPaymentVoidCreditCardTransactionResponseFalha op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagPaymentVoidCreditCardRespFalha()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagPaymentVoidCreditCardRespFalha.Parameters["@id"].Value = op.id_pagto_gw_pag_payment;
				cmUpdatePagPaymentVoidCreditCardRespFalha.Parameters["@ult_atualizacao_usuario"].Value = op.ult_atualizacao_usuario;
				cmUpdatePagPaymentVoidCreditCardRespFalha.Parameters["@ult_id_pagto_gw_pag_payment_op_complementar"].Value = op.ult_id_pagto_gw_pag_payment_op_complementar;
				cmUpdatePagPaymentVoidCreditCardRespFalha.Parameters["@voided_erro_status"].Value = op.voided_erro_status;
				cmUpdatePagPaymentVoidCreditCardRespFalha.Parameters["@voided_erro_data"].Value = (op.voided_erro_data == DateTime.MinValue ? "" : Global.formataDataYyyyMmDdComSeparador(op.voided_erro_data));
				cmUpdatePagPaymentVoidCreditCardRespFalha.Parameters["@voided_erro_data_hora"].Value = (op.voided_erro_data_hora == DateTime.MinValue ? "" : Global.formataDataYyyyMmDdHhMmSsComSeparador(op.voided_erro_data_hora));
				cmUpdatePagPaymentVoidCreditCardRespFalha.Parameters["@voided_erro_mensagem"].Value = (op.voided_erro_mensagem == null ? "" : op.voided_erro_mensagem);
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagPaymentVoidCreditCardRespFalha);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + " com os dados de resposta do método Braspag VoidCreditCardTransaction(): " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id = " + op.id_pagto_gw_pag_payment.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagPaymentRefundCreditCardRespSucesso ]
		public static bool updatePagPaymentRefundCreditCardRespSucesso(BraspagUpdatePagPaymentRefundCreditCardTransactionResponseSucesso op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagPaymentRefundCreditCardRespSucesso()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters["@id"].Value = op.id_pagto_gw_pag_payment;
				cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters["@ult_GlobalStatus"].Value = (op.ult_GlobalStatus == null ? "" : op.ult_GlobalStatus);
				cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters["@ult_atualizacao_usuario"].Value = op.ult_atualizacao_usuario;
				cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters["@ult_id_pagto_gw_pag_payment_op_complementar"].Value = op.ult_id_pagto_gw_pag_payment_op_complementar;
				cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters["@resp_VoidedDate"].Value = (op.resp_VoidedDate == DateTime.MinValue ? "" : Global.formataDataYyyyMmDdHhMmSsComSeparador(op.resp_VoidedDate));
				cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters["@refunded_status"].Value = op.refunded_status;
				cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters["@refunded_data"].Value = (op.refunded_data == DateTime.MinValue ? "" : Global.formataDataYyyyMmDdComSeparador(op.refunded_data));
				cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters["@refunded_data_hora"].Value = (op.refunded_data_hora == DateTime.MinValue ? "" : Global.formataDataYyyyMmDdHhMmSsComSeparador(op.refunded_data_hora));
				cmUpdatePagPaymentRefundCreditCardRespSucesso.Parameters["@refunded_usuario"].Value = op.refunded_usuario ?? "";
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagPaymentRefundCreditCardRespSucesso);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + " com os dados de resposta do método Braspag RefundCreditCardTransaction(): " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id = " + op.id_pagto_gw_pag_payment.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagPaymentRefundCreditCardRespRefundAccepted ]
		public static bool updatePagPaymentRefundCreditCardRespRefundAccepted(BraspagUpdatePagPaymentRefundCreditCardTransactionResponseRefundAccepted op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagPaymentRefundCreditCardRespRefundAccepted()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagPaymentRefundCreditCardRespRefundAccepted.Parameters["@id"].Value = op.id_pagto_gw_pag_payment;
				cmUpdatePagPaymentRefundCreditCardRespRefundAccepted.Parameters["@refund_pending_usuario"].Value = op.ult_atualizacao_usuario;
				cmUpdatePagPaymentRefundCreditCardRespRefundAccepted.Parameters["@ult_GlobalStatus"].Value = (op.ult_GlobalStatus == null ? "" : op.ult_GlobalStatus);
				cmUpdatePagPaymentRefundCreditCardRespRefundAccepted.Parameters["@ult_atualizacao_usuario"].Value = op.ult_atualizacao_usuario;
				cmUpdatePagPaymentRefundCreditCardRespRefundAccepted.Parameters["@ult_id_pagto_gw_pag_payment_op_complementar"].Value = op.ult_id_pagto_gw_pag_payment_op_complementar;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagPaymentRefundCreditCardRespRefundAccepted);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + " com os dados de resposta do método Braspag RefundCreditCardTransaction() com status 'refund accepted': " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id = " + op.id_pagto_gw_pag_payment.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagPaymentRefundCreditCardRespFalha ]
		public static bool updatePagPaymentRefundCreditCardRespFalha(BraspagUpdatePagPaymentRefundCreditCardTransactionResponseFalha op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagPaymentRefundCreditCardRespFalha()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagPaymentRefundCreditCardRespFalha.Parameters["@id"].Value = op.id_pagto_gw_pag_payment;
				cmUpdatePagPaymentRefundCreditCardRespFalha.Parameters["@ult_atualizacao_usuario"].Value = op.ult_atualizacao_usuario;
				cmUpdatePagPaymentRefundCreditCardRespFalha.Parameters["@ult_id_pagto_gw_pag_payment_op_complementar"].Value = op.ult_id_pagto_gw_pag_payment_op_complementar;
				cmUpdatePagPaymentRefundCreditCardRespFalha.Parameters["@refunded_erro_status"].Value = op.refunded_erro_status;
				cmUpdatePagPaymentRefundCreditCardRespFalha.Parameters["@refunded_erro_data"].Value = (op.refunded_erro_data == DateTime.MinValue ? "" : Global.formataDataYyyyMmDdComSeparador(op.refunded_erro_data));
				cmUpdatePagPaymentRefundCreditCardRespFalha.Parameters["@refunded_erro_data_hora"].Value = (op.refunded_erro_data_hora == DateTime.MinValue ? "" : Global.formataDataYyyyMmDdHhMmSsComSeparador(op.refunded_erro_data_hora));
				cmUpdatePagPaymentRefundCreditCardRespFalha.Parameters["@refunded_erro_mensagem"].Value = (op.refunded_erro_mensagem == null ? "" : op.refunded_erro_mensagem);
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagPaymentRefundCreditCardRespFalha);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + " com os dados de resposta do método Braspag RefundCreditCardTransaction(): " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id = " + op.id_pagto_gw_pag_payment.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagPaymentPagtoRegPedido ]
		public static bool updatePagPaymentPagtoRegPedido(BraspagUpdatePagPaymentPagtoRegPedido op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagPaymentPagtoRegPedido()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				if (op.pagto_registrado_no_pedido_usuario == null) op.pagto_registrado_no_pedido_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				if (op.pagto_registrado_no_pedido_usuario.Trim().Length == 0) op.pagto_registrado_no_pedido_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;

				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagPaymentPagtoRegPedido.Parameters["@id"].Value = op.id_pagto_gw_pag_payment;
				cmUpdatePagPaymentPagtoRegPedido.Parameters["@pagto_registrado_no_pedido_status"].Value = op.pagto_registrado_no_pedido_status;
				cmUpdatePagPaymentPagtoRegPedido.Parameters["@pagto_registrado_no_pedido_tipo_operacao"].Value = op.pagto_registrado_no_pedido_tipo_operacao;
				cmUpdatePagPaymentPagtoRegPedido.Parameters["@pagto_registrado_no_pedido_usuario"].Value = op.pagto_registrado_no_pedido_usuario;
				cmUpdatePagPaymentPagtoRegPedido.Parameters["@pagto_registrado_no_pedido_id_pedido_pagamento"].Value = op.pagto_registrado_no_pedido_id_pedido_pagamento;
				cmUpdatePagPaymentPagtoRegPedido.Parameters["@pagto_registrado_no_pedido_st_pagto_anterior"].Value = (op.pagto_registrado_no_pedido_st_pagto_anterior == null ? "" : op.pagto_registrado_no_pedido_st_pagto_anterior);
				cmUpdatePagPaymentPagtoRegPedido.Parameters["@pagto_registrado_no_pedido_st_pagto_novo"].Value = (op.pagto_registrado_no_pedido_st_pagto_novo == null ? "" : op.pagto_registrado_no_pedido_st_pagto_novo);
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagPaymentPagtoRegPedido);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + " as informações de que o pagamento foi registrado no pedido: " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id = " + op.id_pagto_gw_pag_payment.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagPaymentEstornoRegPedido ]
		public static bool updatePagPaymentEstornoRegPedido(BraspagUpdatePagPaymentEstornoRegPedido op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagPaymentEstornoRegPedido()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				if (op.estorno_registrado_no_pedido_usuario == null) op.estorno_registrado_no_pedido_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				if (op.estorno_registrado_no_pedido_usuario.Trim().Length == 0) op.estorno_registrado_no_pedido_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;

				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagPaymentEstornoRegPedido.Parameters["@id"].Value = op.id_pagto_gw_pag_payment;
				cmUpdatePagPaymentEstornoRegPedido.Parameters["@estorno_registrado_no_pedido_status"].Value = op.estorno_registrado_no_pedido_status;
				cmUpdatePagPaymentEstornoRegPedido.Parameters["@estorno_registrado_no_pedido_tipo_operacao"].Value = op.estorno_registrado_no_pedido_tipo_operacao;
				cmUpdatePagPaymentEstornoRegPedido.Parameters["@estorno_registrado_no_pedido_usuario"].Value = op.estorno_registrado_no_pedido_usuario;
				cmUpdatePagPaymentEstornoRegPedido.Parameters["@estorno_registrado_no_pedido_id_pedido_pagamento"].Value = op.estorno_registrado_no_pedido_id_pedido_pagamento;
				cmUpdatePagPaymentEstornoRegPedido.Parameters["@estorno_registrado_no_pedido_st_pagto_anterior"].Value = (op.estorno_registrado_no_pedido_st_pagto_anterior == null ? "" : op.estorno_registrado_no_pedido_st_pagto_anterior);
				cmUpdatePagPaymentEstornoRegPedido.Parameters["@estorno_registrado_no_pedido_st_pagto_novo"].Value = (op.estorno_registrado_no_pedido_st_pagto_novo == null ? "" : op.estorno_registrado_no_pedido_st_pagto_novo);
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagPaymentEstornoRegPedido);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + " as informações de que o estorno foi registrado no pedido: " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id = " + op.id_pagto_gw_pag_payment.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagPaymentAFFinalizado ]
		public static bool updatePagPaymentAFFinalizado(int id, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagPaymentAFFinalizado()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros  ]
				cmUpdatePagPaymentAFFinalizado.Parameters["@id"].Value = id;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagPaymentAFFinalizado);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = "t_PAGTO_GW_PAG_PAYMENT.id = " + id.ToString();
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.Message;

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = "t_PAGTO_GW_PAG_PAYMENT.id = " + id.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagPaymentRefundPendingConfirmado ]
		public static bool updatePagPaymentRefundPendingConfirmado(BraspagUpdatePagPaymentRefundPendingConfirmado op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagPaymentRefundPendingConfirmado()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				if ((op.refund_pending_confirmado_usuario ?? "").Trim().Length == 0) op.refund_pending_confirmado_usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;

				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagPaymentRefundPendingConfirmado.Parameters["@id"].Value = op.id_pagto_gw_pag_payment;
				cmUpdatePagPaymentRefundPendingConfirmado.Parameters["@refund_pending_confirmado_usuario"].Value = op.refund_pending_confirmado_usuario;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagPaymentRefundPendingConfirmado);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + " com as informações de que o estorno pendente foi confirmado: " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id = " + op.id_pagto_gw_pag_payment.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagPaymentRefundPendingFalha ]
		public static bool updatePagPaymentRefundPendingFalha(BraspagUpdatePagPaymentRefundPendingFalha op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updatePagPaymentRefundPendingFalha()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagPaymentRefundPendingFalha.Parameters["@id"].Value = op.id_pagto_gw_pag_payment;
				cmUpdatePagPaymentRefundPendingFalha.Parameters["@refund_pending_falha_motivo"].Value = op.refund_pending_falha_motivo;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagPaymentRefundPendingFalha);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + " com as informações de que a verificação do estorno pendente foi abortado definitivamente por falha: " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id = " + op.id_pagto_gw_pag_payment.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updateWebhookQueryDadosComplementaresFalhaDefinitiva ]
		public static bool updateWebhookQueryDadosComplementaresFalhaDefinitiva(BraspagUpdateWebhookQueryDadosComplementaresFalhaDefinitiva op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updateWebhookQueryDadosComplementaresFalhaDefinitiva()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva.Parameters["@Id"].Value = op.id_braspag_webhook;
				cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva.Parameters["@EmailEnviadoStatus"].Value = op.EmailEnviadoStatus;
				cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva.Parameters["@BraspagDadosComplementaresQueryStatus"].Value = op.BraspagDadosComplementaresQueryStatus;
				cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva.Parameters["@BraspagDadosComplementaresQueryTentativas"].Value = op.BraspagDadosComplementaresQueryTentativas;
				cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva.Parameters["@MsgErro"].Value = (op.MsgErro ?? "");
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateWebhookQueryDadosComplementaresFalhaDefinitiva);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com as informações indicando falha definitiva na consulta dos dados complementares: " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id = " + op.id_braspag_webhook.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updateWebhookQueryDadosComplementaresFalhaTemporaria ]
		public static bool updateWebhookQueryDadosComplementaresFalhaTemporaria(BraspagUpdateWebhookQueryDadosComplementaresFalhaTemporaria op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updateWebhookQueryDadosComplementaresFalhaTemporaria()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateWebhookQueryDadosComplementaresFalhaTemporaria.Parameters["@Id"].Value = op.id_braspag_webhook;
				cmUpdateWebhookQueryDadosComplementaresFalhaTemporaria.Parameters["@BraspagDadosComplementaresQueryStatus"].Value = op.BraspagDadosComplementaresQueryStatus;
				cmUpdateWebhookQueryDadosComplementaresFalhaTemporaria.Parameters["@BraspagDadosComplementaresQueryTentativas"].Value = op.BraspagDadosComplementaresQueryTentativas;
				cmUpdateWebhookQueryDadosComplementaresFalhaTemporaria.Parameters["@MsgErroTemporario"].Value = (op.MsgErroTemporario ?? "");
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateWebhookQueryDadosComplementaresFalhaTemporaria);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com as informações indicando falha temporária na consulta dos dados complementares: " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id = " + op.id_braspag_webhook.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updateWebhookQueryDadosComplementaresQtdeTentativas ]
		public static bool updateWebhookQueryDadosComplementaresQtdeTentativas(BraspagUpdateWebhookQueryDadosComplementaresQtdeTentativas op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updateWebhookQueryDadosComplementaresQtdeTentativas()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateWebhookQueryDadosComplementaresQtdeTentativas.Parameters["@Id"].Value = op.id_braspag_webhook;
				cmUpdateWebhookQueryDadosComplementaresQtdeTentativas.Parameters["@BraspagDadosComplementaresQueryTentativas"].Value = op.BraspagDadosComplementaresQueryTentativas;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateWebhookQueryDadosComplementaresQtdeTentativas);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com as informações indicando falha temporária na consulta dos dados complementares: " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id = " + op.id_braspag_webhook.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updateWebhookQueryDadosComplementaresSucesso ]
		public static bool updateWebhookQueryDadosComplementaresSucesso(BraspagUpdateWebhookQueryDadosComplementaresSucesso op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updateWebhookQueryDadosComplementaresSucesso()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateWebhookQueryDadosComplementaresSucesso.Parameters["@Id"].Value = op.id_braspag_webhook;
				cmUpdateWebhookQueryDadosComplementaresSucesso.Parameters["@BraspagDadosComplementaresQueryStatus"].Value = op.BraspagDadosComplementaresQueryStatus;
				cmUpdateWebhookQueryDadosComplementaresSucesso.Parameters["@BraspagDadosComplementaresQueryTentativas"].Value = op.BraspagDadosComplementaresQueryTentativas;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateWebhookQueryDadosComplementaresSucesso);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com as informações indicando sucesso na consulta dos dados complementares: " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id = " + op.id_braspag_webhook.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ insereWebhookQueryDadosComplementares ]
		public static bool insereWebhookQueryDadosComplementares(BraspagInsertWebhookQueryDadosComplementares op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.insereWebhookQueryDadosComplementares()";
			int generatedId;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmInsertWebhookQueryDadosComplementares.Parameters["@id_braspag_webhook"].Value = op.id_braspag_webhook;
				cmInsertWebhookQueryDadosComplementares.Parameters["@BraspagTransactionId"].Value = (op.BraspagTransactionId ?? "");
				cmInsertWebhookQueryDadosComplementares.Parameters["@BraspagOrderId"].Value = (op.BraspagOrderId ?? "");
				cmInsertWebhookQueryDadosComplementares.Parameters["@PaymentMethod"].Value = (op.PaymentMethod ?? "");
				cmInsertWebhookQueryDadosComplementares.Parameters["@GlobalStatus"].Value = (op.GlobalStatus ?? "");
				cmInsertWebhookQueryDadosComplementares.Parameters["@ReceivedDate"].Value = Global.formataDataYyyyMmDdHhMmSsComSeparador(op.ReceivedDate);
				cmInsertWebhookQueryDadosComplementares.Parameters["@CapturedDate"].Value = Global.formataDataYyyyMmDdHhMmSsComSeparador(op.CapturedDate);
				cmInsertWebhookQueryDadosComplementares.Parameters["@CustomerName"].Value = (op.CustomerName ?? "");
				cmInsertWebhookQueryDadosComplementares.Parameters["@BoletoExpirationDate"].Value = Global.formataDataYyyyMmDdHhMmSsComSeparador(op.BoletoExpirationDate);
				cmInsertWebhookQueryDadosComplementares.Parameters["@Amount"].Value = (op.Amount ?? "");
				cmInsertWebhookQueryDadosComplementares.Parameters["@ValorAmount"].Value = op.ValorAmount;
				cmInsertWebhookQueryDadosComplementares.Parameters["@PaidAmount"].Value = (op.PaidAmount ?? "");
				cmInsertWebhookQueryDadosComplementares.Parameters["@ValorPaidAmount"].Value = op.ValorPaidAmount;
				cmInsertWebhookQueryDadosComplementares.Parameters["@pedido"].Value = (op.pedido ?? "");
				#endregion

				#region [ Tenta inserir o registro ]
				try
				{
					generatedId = (int)BD.executeScalar(ref cmInsertWebhookQueryDadosComplementares);
					op.Id = generatedId;
				}
				catch (Exception ex)
				{
					generatedId = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.Message;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion
				}
				#endregion

				#region [ Gravou o registro? ]
				if (generatedId == 0)
				{
					msg_erro = "Falha ao tentar gravar o registro na tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + "\n" + msg_erro;
					return false;
				}
				#endregion

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Sucesso na gravação dos dados (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + op.Id.ToString() + ")");

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.Message;

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updateWebhookComplementarPagtoRegPedido ]
		public static bool updateWebhookComplementarPagtoRegPedido(BraspagWebhookComplementarUpdatePagtoRegPedido op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updateWebhookComplementarPagtoRegPedido()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateWebhookComplementarPagtoRegPedido.Parameters["@Id"].Value = op.id_braspag_webhook_complementar;
				cmUpdateWebhookComplementarPagtoRegPedido.Parameters["@PagtoRegistradoNoPedidoStatus"].Value = op.PagtoRegistradoNoPedidoStatus;
				cmUpdateWebhookComplementarPagtoRegPedido.Parameters["@PagtoRegistradoNoPedidoTipoOperacao"].Value = op.PagtoRegistradoNoPedidoTipoOperacao;
				cmUpdateWebhookComplementarPagtoRegPedido.Parameters["@PagtoRegistradoNoPedido_id_pedido_pagamento"].Value = op.PagtoRegistradoNoPedido_id_pedido_pagamento;
				cmUpdateWebhookComplementarPagtoRegPedido.Parameters["@PagtoRegistradoNoPedidoStPagtoAnterior"].Value = (op.PagtoRegistradoNoPedidoStPagtoAnterior == null ? "" : op.PagtoRegistradoNoPedidoStPagtoAnterior);
				cmUpdateWebhookComplementarPagtoRegPedido.Parameters["@PagtoRegistradoNoPedidoStPagtoNovo"].Value = (op.PagtoRegistradoNoPedidoStPagtoNovo == null ? "" : op.PagtoRegistradoNoPedidoStPagtoNovo);
				cmUpdateWebhookComplementarPagtoRegPedido.Parameters["@AnaliseCreditoStatusAnterior"].Value = op.AnaliseCreditoStatusAnterior;
				cmUpdateWebhookComplementarPagtoRegPedido.Parameters["@AnaliseCreditoStatusNovo"].Value = op.AnaliseCreditoStatusNovo;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateWebhookComplementarPagtoRegPedido);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + " as informações de que o pagamento foi registrado no pedido: " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id = " + op.id_braspag_webhook_complementar.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(op);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.serializaObjectToXml(op);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updateWebhookEmailEnviadoStatusFalha ]
		public static bool updateWebhookEmailEnviadoStatusFalha(int id_braspag_webhook, byte status, string mensagemErro, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updateWebhookEmailEnviadoStatusFalha()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateWebhookEmailEnviadoStatusFalha.Parameters["@Id"].Value = id_braspag_webhook;
				cmUpdateWebhookEmailEnviadoStatusFalha.Parameters["@EmailEnviadoStatus"].Value = status;
				cmUpdateWebhookEmailEnviadoStatusFalha.Parameters["@MsgErro"].Value = (mensagemErro ?? "");
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateWebhookEmailEnviadoStatusFalha);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id = " + id_braspag_webhook.ToString();
					svcLog.complemento_2 = "EmailEnviadoStatus = " + status.ToString();
					svcLog.complemento_3 = "MsgErro = " + mensagemErro;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com as informações indicando falha no envio do email informativo: " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id = " + id_braspag_webhook.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id = " + id_braspag_webhook.ToString();
					svcLog.complemento_2 = "EmailEnviadoStatus = " + status.ToString();
					svcLog.complemento_3 = "MsgErro = " + mensagemErro;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id = " + id_braspag_webhook.ToString();
				svcLog.complemento_2 = "EmailEnviadoStatus = " + status.ToString();
				svcLog.complemento_3 = "MsgErro = " + mensagemErro;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updateWebhookEmailEnviadoStatusSucesso ]
		public static bool updateWebhookEmailEnviadoStatusSucesso(int id_braspag_webhook, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updateWebhookEmailEnviadoStatusSucesso()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateWebhookEmailEnviadoStatusSucesso.Parameters["@Id"].Value = id_braspag_webhook;
				cmUpdateWebhookEmailEnviadoStatusSucesso.Parameters["@EmailEnviadoStatus"].Value = Global.Cte.Braspag.Webhook.EmailEnviadoStatus.EnviadoComSucesso;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateWebhookEmailEnviadoStatusSucesso);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + id_braspag_webhook.ToString();
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com as informações indicando sucesso no envio do email informativo: " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id = " + id_braspag_webhook.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + id_braspag_webhook.ToString();
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + id_braspag_webhook.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updateWebhookProcessamentoErpStatusFalha ]
		public static bool updateWebhookProcessamentoErpStatusFalha(int id_braspag_webhook, int status, string mensagemErro, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updateWebhookProcessamentoErpStatusFalha()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateWebhookProcessamentoErpStatusFalha.Parameters["@Id"].Value = id_braspag_webhook;
				cmUpdateWebhookProcessamentoErpStatusFalha.Parameters["@ProcessamentoErpStatus"].Value = status;
				cmUpdateWebhookProcessamentoErpStatusFalha.Parameters["@MsgErro"].Value = (mensagemErro ?? "");
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateWebhookProcessamentoErpStatusFalha);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id = " + id_braspag_webhook.ToString();
					svcLog.complemento_2 = "ProcessamentoErpStatus = " + status.ToString();
					svcLog.complemento_3 = "MsgErro = " + mensagemErro;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com as informações indicando falha no processamento ERP: " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id = " + id_braspag_webhook.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id = " + id_braspag_webhook.ToString();
					svcLog.complemento_2 = "ProcessamentoErpStatus = " + status.ToString();
					svcLog.complemento_3 = "MsgErro = " + mensagemErro;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id = " + id_braspag_webhook.ToString();
				svcLog.complemento_2 = "ProcessamentoErpStatus = " + status.ToString();
				svcLog.complemento_3 = "MsgErro = " + mensagemErro;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updateWebhookProcessamentoErpStatusSucesso ]
		public static bool updateWebhookProcessamentoErpStatusSucesso(int id_braspag_webhook, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.updateWebhookProcessamentoErpStatusSucesso()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateWebhookProcessamentoErpStatusSucesso.Parameters["@Id"].Value = id_braspag_webhook;
				cmUpdateWebhookProcessamentoErpStatusSucesso.Parameters["@ProcessamentoErpStatus"].Value = Global.Cte.Braspag.Webhook.ProcessamentoErpStatus.ProcessadoComSucesso;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateWebhookProcessamentoErpStatusSucesso);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + id_braspag_webhook.ToString();
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				if (intRetorno != 1)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = "Falha ao tentar atualizar o registro da tabela " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + " com as informações indicando sucesso no processamento ERP: " + intRetorno.ToString() + " registro(s) alterado(s) ao invés de 1 (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id = " + id_braspag_webhook.ToString() + ")";

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + id_braspag_webhook.ToString();
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + id_braspag_webhook.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ transacaoJaRegistrouPagtoNoPedido ]
		public static bool transacaoJaRegistrouPagtoNoPedido(string braspagTransactionId, out BraspagWebhookComplementar braspagWebhookComplementar, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.transacaoJaRegistrouPagtoNoPedido()";
			string msg_erro_aux;
			string strSql;
			int id;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			braspagWebhookComplementar = null;
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
							" Id" +
						" FROM t_BRASPAG_WEBHOOK_COMPLEMENTAR" +
						" WHERE" +
							" (BraspagTransactionId = '" + braspagTransactionId + "')" +
							" AND (GlobalStatus = '" + Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA + "')" +
							" AND (PagtoRegistradoNoPedidoStatus = 1)" +
						" ORDER BY" +
							" Id DESC";

				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count == 0) return false;

				id = BD.readToInt(dtbResultado.Rows[0]["Id"]);
				braspagWebhookComplementar = getBraspagWebhookComplementarById(id, out msg_erro_aux);
				
				if (braspagWebhookComplementar == null) return false;
				if (braspagWebhookComplementar.Id == 0) return false;
				
				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = ex.Message;

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = $"BraspagTransactionId={braspagTransactionId}";
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ registraPagamentoNoPedido ]
		public static bool registraPagamentoNoPedido(Global.Cte.Braspag.Pagador.Transacao tipoTransacao, int id_pagto_gw_pag_payment, out string msg_erro)
		{
			// IMPORTANTE: AS ALTERAÇÕES NAS REGRAS DEVEM ESTAR SINCRONIZADAS ENTRE AS ROTINAS BraspagClearsaleRegistraPagtoNoPedido() de BraspagCS.asp E registraPagamentoNoPedido() DE BraspagDAO.cs (FinanceiroService)
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.registraPagamentoNoPedido()";
			bool blnUpdateHistPagto;
			int id_emailsndsvc_mensagem;
			int? st_analise_credito_novo = null;
			int statusHistPagto;
			Decimal vlTotalFamiliaPrecoVenda;
			Decimal vlTotalFamiliaPrecoNF;
			Decimal vlTotalFamiliaPago;
			Decimal vlTotalFamiliaDevolucaoPrecoVenda;
			Decimal vlTotalFamiliaDevolucaoPrecoNF;
			Decimal vlTotalFamiliaPagoNovo;
			string msg_erro_aux;
			string strSubject;
			string strBody;
			string strMsg;
			string id_pedido_base;
			string st_pagto;
			string st_pagto_novo = "";
			string s_ult_AF_status = "";
			string s_descricao_ult_AF_status = "(sem status)";
			string s_log = "";
			BraspagPag pag;
			BraspagPagPayment payment = null;
			PedidoPagamento pedidoPagto;
			Pedido pedidoBase;
			ClearsaleAF clearsaleAF;
			BraspagUpdatePagPaymentPagtoRegPedido updatePagPaymentPagtoRegPedido;
			BraspagUpdatePagPaymentEstornoRegPedido updatePagPaymentEstornoRegPedido;
			Global.Cte.Braspag.Pagador.OperacaoRegistraPagtoPedido opRegistraPagtoPedido = null;
			PedidoHistPagto histPagto = null;
			#endregion

			msg_erro = "";
			try
			{
				// Obtém os dados atualizados do BD
				payment = BraspagDAO.getBraspagPagPaymentById(id_pagto_gw_pag_payment, out msg_erro_aux);
				pag = BraspagDAO.getBraspagPagById(payment.id_pagto_gw_pag, out msg_erro_aux);
				id_pedido_base = Global.retornaNumeroPedidoBase(pag.pedido);

				#region [ Calcula valores do pedido ]
				if (!PedidoDAO.calculaPagamentos(pag.pedido, out vlTotalFamiliaPrecoVenda, out vlTotalFamiliaPrecoNF, out vlTotalFamiliaPago, out vlTotalFamiliaDevolucaoPrecoVenda, out vlTotalFamiliaDevolucaoPrecoNF, out st_pagto, out msg_erro_aux))
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = msg_erro_aux;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(payment);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar calcular valores do pedido durante operação de registrar o pagamento no pedido [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\nFalha ao tentar calcular valores do pedido durante operação de registrar o pagamento no pedido " + pag.pedido + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\nChamada à rotina PedidoDAO.calculaPagamentos() retornou erro:\r\n" + msg_erro;
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
					}

					return false;
				}
				#endregion

				#region [ Registra o pagamento no pedido ]
				// Somente nos seguintes casos:
				//		1) Captura
				//		2) Estorno (implica que a captura já foi realizada em alguma data passada)
				//		3) Cancelamento (Void) se a captura foi realizada no mesmo dia (lembrando que o método Void é usado p/ cancelar uma transação Autorizada ou p/ uma transação Capturada até a meia-noite do mesmo dia)
				if (
					tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.CaptureCreditCardTransaction.GetCodOpLog())
					||
					tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.RefundCreditCardTransaction.GetCodOpLog())
					||
					(tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.VoidCreditCardTransaction.GetCodOpLog()) && (payment.captura_confirmada_status == 1))
					)
				{
					#region [ Grava pagamento no pedido ]
					pedidoPagto = new PedidoPagamento();
					pedidoPagto.pedido = pag.pedido;
					pedidoPagto.valor = payment.valor_transacao;
					// Se for estorno/cancelamento, o valor deve ser subtraído dos créditos do pedido
					if (
						tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.RefundCreditCardTransaction.GetCodOpLog())
						||
						(tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.VoidCreditCardTransaction.GetCodOpLog()) && (payment.captura_confirmada_status == 1))
						)
					{
						if (pedidoPagto.valor > 0) pedidoPagto.valor = -1 * pedidoPagto.valor;
					}

					pedidoPagto.tipo_pagto = Global.Cte.PedidoPagtoTipoOperacao.GW_BRASPAG_CLEARSALE;
					pedidoPagto.usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
					pedidoPagto.id_pagto_gw_pag_payment = payment.id;
					if (!PedidoDAO.inserePedidoPagamento(pedidoPagto, out msg_erro_aux))
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = Global.serializaObjectToXml(payment);
						svcLog.complemento_2 = Global.serializaObjectToXml(pedidoPagto);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar registrar o pagamento no pedido [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\nFalha ao tentar registrar o pagamento no pedido " + pag.pedido + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\nChamada à rotina PedidoDAO.inserePedidoPagamento() retornou erro:\r\n" + msg_erro;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}

						return false;
					}
					#endregion

					#region [ Determina alterações no status de pagamento do pedido (Quitado, Pago Parcial, Não-Pago) e no status de análise de crédito ]
					pedidoBase = PedidoDAO.getPedido(id_pedido_base);

					clearsaleAF = ClearsaleDAO.getClearsaleAFByIdPagtoGwPagPayment(payment.id, out msg_erro_aux);
					if (clearsaleAF != null)
					{
						if (clearsaleAF.ult_Status != null)
						{
							s_ult_AF_status = clearsaleAF.ult_Status;
							s_descricao_ult_AF_status = s_ult_AF_status + " - " + Global.Cte.Clearsale.StatusAF.GetDescription(s_ult_AF_status);
						}
					}

					if ((vlTotalFamiliaDevolucaoPrecoNF + vlTotalFamiliaPago + pedidoPagto.valor) >= (vlTotalFamiliaPrecoNF - Global.Cte.Etc.MAX_VALOR_MARGEM_ERRO_PAGAMENTO))
					{
						#region [ Pago (Quitado) ]
						st_pagto_novo = Global.Cte.StPagtoPedido.ST_PAGTO_PAGO;
						s_log += "Status de pagamento do pedido: quitado (st_pagto: " + pedidoBase.st_pagto + " => " + st_pagto_novo + ")";
						if ((vlTotalFamiliaDevolucaoPrecoNF + vlTotalFamiliaPago + pedidoPagto.valor) > vlTotalFamiliaPrecoNF)
						{
							s_log += " (excedeu " + Global.Cte.Etc.SIMBOLO_MONETARIO + " " + Global.formataMoeda((vlTotalFamiliaDevolucaoPrecoNF + vlTotalFamiliaPago + pedidoPagto.valor) - vlTotalFamiliaPrecoNF) + ")";
						}
						else if ((vlTotalFamiliaDevolucaoPrecoNF + vlTotalFamiliaPago + pedidoPagto.valor) < vlTotalFamiliaPrecoNF)
						{
							s_log += " (faltou " + Global.Cte.Etc.SIMBOLO_MONETARIO + " " + Global.formataMoeda(vlTotalFamiliaPrecoNF - (vlTotalFamiliaDevolucaoPrecoNF + vlTotalFamiliaPago + pedidoPagto.valor)) + ")";
						}

						// IMPORTANTE: AS ALTERAÇÕES NAS REGRAS DEVEM ESTAR SINCRONIZADAS ENTRE AS ROTINAS BraspagClearsaleRegistraPagtoNoPedido() de BraspagCS.asp E registraPagamentoNoPedido() DE BraspagDAO.cs (FinanceiroService)
						// ===========
						// OBSERVAÇÃO: NO FLUXO COM A CLEARSALE, O FLUXO SEGUE A SEGUINTE SEQUENCIA:
						//		1) ENVIO DA TRANSAÇÃO DE AUTORIZAÇÃO DO PAGAMENTO PARA A BRASPAG
						//		2) CASO O PAGAMENTO TENHA SIDO AUTORIZADO, É ENVIADA A TRANSAÇÃO PARA ANÁLISE ANTIFRAUDE (CLEARSALE)
						//		3-A) SE A ANÁLISE ANTIFRAUDE APROVOU A TRANSAÇÃO, É FEITA A CAPTURA DA TRANSAÇÃO
						//		3-B) SE A ANÁLISE ANTIFRAUDE REPROVOU A TRANSAÇÃO, É FEITO O CANCELAMENTO/ESTORNO
						// A ANÁLISE DE ANTIFRAUDE É FEITA POR EQUIPE PRÓPRIA E ASSUME-SE QUE DURANTE A ANÁLISE ANTIFRAUDE FORAM FEITAS TODAS AS VERIFICAÇÕES NECESSÁRIAS
						// PORTANTO, A APROVAÇÃO POR PARTE DO ANALISTA INTERNO SIGNIFICA QUE A ANÁLISE DE CRÉDITO ESTÁ OK
						// IMPORTANTE: O CLIENTE PODE REALIZAR O PAGAMENTO UTILIZANDO VÁRIOS CARTÕES DE CRÉDITO. ESSAS N TRANSAÇÕES SERÃO ENVIADAS JUNTAS P/ UMA ÚNICA REQUISIÇÃO DE
						// ==========  ANÁLISE ANTIFRAUDE P/ A CLEARSALE. QUANDO A ANÁLISE AF ESTIVER CONCLUÍDA, O PROCESSAMENTO FINAL DE CAPTURA OU CANCELAMENTO/ESTORNO DAS TRANSAÇÕES
						// SERÁ FEITA NO FINANCEIROSERVICE. CADA TRANSAÇÃO DE CARTÃO ENVOLVIDA NO PAGAMENTO DO PEDIDO IRÁ ACIONAR UMA VEZ A ROTINA BraspagDAO.registraPagamentoNoPedido()
						// PORTANTO, A 1ª TRANSAÇÃO IRÁ ALTERAR O PEDIDO PARA O STATUS DE PAGAMENTO 'PARCIAL' E APENAS QUANDO A ÚLTIMA TRANSAÇÃO FOR PROCESSADA, O PEDIDO FICARÁ
						// COM O STATUS 'PAGO' E A ANÁLISE DE CRÉDITO DEVE FICAR C/ O STATUS 'OK'.
						// ENTRETANTO, O SISTEMA NÃO SABE DE ANTEMÃO SE O PAGAMENTO SERÁ INTEGRALIZADO OU NÃO, PORTANTO, ENQUANTO O STATUS DE PAGAMENTO ESTIVER PARCIAL, O STATUS
						// DA ANÁLISE DE CRÉDITO DEVE FICAR COMO 'PENDENTE VENDAS', POIS CASO O VALOR NÃO SEJA INTEGRALIZADO, ESSE SERÁ O STATUS FINAL.
						if (s_ult_AF_status.Equals(Global.Cte.Clearsale.StatusAF.APROVACAO_MANUAL.GetValue()))
						{
							#region [ Tratamento para Aprovação Manual ]
							if (
									(
										(pedidoBase.analise_credito == (short)Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.ST_INICIAL)
										||
										(pedidoBase.analise_credito == (short)Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS)
									)
									&&
									(pedidoBase.analise_credito != (short)Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK)
								)
							{
								st_analise_credito_novo = Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK;
								if (s_log.Length > 0) s_log += "; ";
								s_log += " Análise de crédito: " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " => " + Global.obtemDescricaoAnaliseCredito((int)st_analise_credito_novo) + " (status AF Clearsale: '" + s_descricao_ult_AF_status + "')";
							}
							else
							{
								if (s_log.Length > 0) s_log += "; ";
								s_log += " Análise de crédito: status não foi alterado porque pedido encontra-se em " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " (status AF Clearsale: '" + s_descricao_ult_AF_status + "')";
							}
							#endregion
						}
						else if (s_ult_AF_status.Equals(Global.Cte.Clearsale.StatusAF.APROVACAO_AUTOMATICA.GetValue()) || s_ult_AF_status.Equals(Global.Cte.Clearsale.StatusAF.APROVACAO_POR_POLITICA.GetValue()))
						{
							#region [ Tratamento para Aprovação Automática ]
							// EM CASO DE APROVAÇÃO AUTOMÁTICA, COLOCA-SE EM 'PENDENTE VENDAS' PARA DAR OPORTUNIDADE AO ANALISTA CONFERIR A TITULARIDADE DO CARTÃO
							if (
									(pedidoBase.analise_credito != Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS)
									&&
									(pedidoBase.analise_credito != (short)Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK)
								)
							{
								st_analise_credito_novo = Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS;
								if (s_log.Length > 0) s_log += "; ";
								s_log += " Análise de crédito: " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " => " + Global.obtemDescricaoAnaliseCredito((int)st_analise_credito_novo) + " (status AF Clearsale: '" + s_descricao_ult_AF_status + "')";
							}
							else
							{
								if (s_log.Length > 0) s_log += "; ";
								s_log += " Análise de crédito: status não foi alterado porque pedido encontra-se em " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " (status AF Clearsale: '" + s_descricao_ult_AF_status + "')";
							}
							#endregion
						}
						else
						{
							#region [ Tratamento para outros status ]
							if (
									(pedidoBase.analise_credito != Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS)
									&&
									(pedidoBase.analise_credito != (short)Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK)
								)
							{
								st_analise_credito_novo = Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS;
								if (s_log.Length > 0) s_log += "; ";
								s_log += " Análise de crédito: " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " => " + Global.obtemDescricaoAnaliseCredito((int)st_analise_credito_novo) + " (status AF Clearsale: '" + s_descricao_ult_AF_status + "')";
							}
							else
							{
								if (s_log.Length > 0) s_log += "; ";
								s_log += " Análise de crédito: status não foi alterado porque pedido encontra-se em " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " (status AF Clearsale: '" + s_descricao_ult_AF_status + "')";
							}
							#endregion
						}
						#endregion
					}
					else if ((vlTotalFamiliaPago + pedidoPagto.valor) > 0)
					{
						#region [ Pagamento Parcial ]
						st_pagto_novo = Global.Cte.StPagtoPedido.ST_PAGTO_PARCIAL;
						s_log += "Status de pagamento do pedido: pago parcial (st_pagto: " + pedidoBase.st_pagto + " => " + st_pagto_novo + ")";

						// SE O STATUS É 'PAGO PARCIAL', PODE TER HAVIDO UMA OPERAÇÃO DE CAPTURA OU DE CANCELAMENTO/ESTORNO. NESTE CASO, AS SEGUINTES PREMISSAS SÃO SEGUIDAS:
						//	1) SE O PEDIDO ESTIVER COM 'CRÉDITO OK', NÃO SERÁ ALTERADO DEVIDO A CANCELAMENTO/ESTORNO (DEFINIDO PELA ROSE EM 22/06/2016)
						//	2) SE O PEDIDO ESTIVER COM 'PENDENTE VENDAS', NÃO SERÁ ALTERADO. NÃO HÁ NECESSIDADE DE ATUALIZAR A DATA DA ÚLTIMA ALTERAÇÃO DE STATUS, POIS PEDIDOS C/ STATUS DE PAGTO 'PAGO PARCIAL' NÃO SÃO CANCELADOS AUTOMATICAMENTE
						if (
								(pedidoBase.analise_credito != Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS)
								&&
								(pedidoBase.analise_credito != (short)Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK)
							)
						{
							st_analise_credito_novo = Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS;
							if (s_log.Length > 0) s_log += "; ";
							s_log += " Análise de crédito: " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " => " + Global.obtemDescricaoAnaliseCredito((int)st_analise_credito_novo) + " (status AF Clearsale: '" + s_descricao_ult_AF_status + "')";
						}
						else
						{
							if (s_log.Length > 0) s_log += "; ";
							s_log += " Análise de crédito: status não foi alterado porque pedido encontra-se em " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " (status AF Clearsale: '" + s_descricao_ult_AF_status + "')";
						}
						#endregion
					}
					else
					{
						#region [ Não-Pago ]
						st_pagto_novo = Global.Cte.StPagtoPedido.ST_PAGTO_NAO_PAGO;
						s_log += "Status de pagamento do pedido: não-pago (st_pagto: " + pedidoBase.st_pagto + " => " + st_pagto_novo + ")";

						// SE O STATUS É 'NÃO PAGO', ENTÃO OCORREU UMA OPERAÇÃO DE CANCELAMENTO/ESTORNO. NESTE CASO, AS SEGUINTES PREMISSAS SÃO SEGUIDAS:
						//	1) SE O PEDIDO ESTIVER COM 'CRÉDITO OK', NÃO SERÁ ALTERADO DEVIDO A CANCELAMENTO/ESTORNO (DEFINIDO PELA ROSE EM 22/06/2016)
						//	2) SE O PEDIDO ESTIVER COMO 'PENDENTE VENDAS', CONTINUA COMO ESTÁ E A DATA DA ÚLTIMA ALTERAÇÃO DE STATUS NÃO É ALTERADA, MANTENDO A CONTAGEM ORIGINAL DO PERÍODO DE CANCELAMENTO AUTOMÁTICO DE PEDIDOS
						if (
								(pedidoBase.analise_credito != Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS)
								&&
								(pedidoBase.analise_credito != (short)Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK)
							)
						{
							st_analise_credito_novo = Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS;
							if (s_log.Length > 0) s_log += "; ";
							s_log += " Análise de crédito: " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " => " + Global.obtemDescricaoAnaliseCredito((int)st_analise_credito_novo) + " (status AF Clearsale: '" + s_descricao_ult_AF_status + "')";
						}
						else
						{
							if (s_log.Length > 0) s_log += "; ";
							s_log += " Análise de crédito: status não foi alterado porque pedido encontra-se em " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " (status AF Clearsale: '" + s_descricao_ult_AF_status + "')";
						}
						#endregion
					}
					#endregion

					#region [ Atualiza status de pagamento do pedido ]
					if (st_pagto_novo.Length > 0)
					{
						if (!PedidoDAO.atualizaPedidoStatusPagto(id_pedido_base, st_pagto_novo, out msg_erro_aux))
						{
							// Retorna mensagem de erro p/ rotina chamadora
							msg_erro = msg_erro_aux;

							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = msg_erro;
							svcLog.complemento_1 = Global.serializaObjectToXml(payment);
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar atualizar o status de pagamento do pedido " + id_pedido_base + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o status de pagamento do pedido " + id_pedido_base + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\nChamada à rotina PedidoDAO.atualizaPedidoStatusPagto() retornou erro:\r\n" + msg_erro;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}

							return false;
						}
					}
					#endregion

					#region [ Atualiza status da análise de crédito do pedido ]
					if (st_analise_credito_novo != null)
					{
						if (!PedidoDAO.atualizaPedidoStatusAnaliseCredito(id_pedido_base, st_analise_credito_novo, Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA, out msg_erro_aux))
						{
							// Retorna mensagem de erro p/ rotina chamadora
							msg_erro = msg_erro_aux;

							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = msg_erro;
							svcLog.complemento_1 = Global.serializaObjectToXml(payment);
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar atualizar o status da análise de crédito do pedido " + id_pedido_base + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o status da análise de crédito do pedido " + id_pedido_base + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\nChamada à rotina PedidoDAO.atualizaPedidoStatusAnaliseCredito() retornou erro:\r\n" + msg_erro;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}

							return false;
						}
					}
					#endregion

					#region [ Atualiza o campo 'vl_pago_familia' no pedido ]
					vlTotalFamiliaPagoNovo = vlTotalFamiliaPago + pedidoPagto.valor;
					if (!PedidoDAO.atualizaPedidoVlPagoFamilia(id_pedido_base, vlTotalFamiliaPagoNovo, out msg_erro_aux))
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = Global.serializaObjectToXml(payment);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar atualizar o valor total pago (família de pedidos) do pedido " + id_pedido_base + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o valor total pago (família de pedidos) do pedido " + id_pedido_base + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\nChamada à rotina PedidoDAO.atualizaPedidoVlPagoFamilia() retornou erro:\r\n" + msg_erro;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}

						return false;
					}
					#endregion

					#region [ Registra em t_PAGTO_GW_PAG_PAYMENT que foi gerado pagamento/estorno + Grava log ]
					// Atualiza t_PAGTO_GW_PAG_PAYMENT os campos que indicam se o pagamento/estorno foi registrado no pedido
					if (tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.CaptureCreditCardTransaction.GetCodOpLog()))
					{
						opRegistraPagtoPedido = Global.Cte.Braspag.Pagador.OperacaoRegistraPagtoPedido.CAPTURA;
					}
					else if (tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.VoidCreditCardTransaction.GetCodOpLog()) && (payment.captura_confirmada_status == 1))
					{
						opRegistraPagtoPedido = Global.Cte.Braspag.Pagador.OperacaoRegistraPagtoPedido.CANCELAMENTO;
					}
					else if (tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.RefundCreditCardTransaction.GetCodOpLog()))
					{
						opRegistraPagtoPedido = Global.Cte.Braspag.Pagador.OperacaoRegistraPagtoPedido.ESTORNO;
					}
					else
					{
						// Situação não prevista
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = "Tipo de operação não previsto ao processar o registro do pagamento no pedido " + pag.pedido + ": " + tipoTransacao.GetCodOpLog();

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = Global.serializaObjectToXml(payment);
						svcLog.complemento_2 = Global.serializaObjectToXml(tipoTransacao);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Tipo de operação não previsto ao processar o registro do pagamento no pedido " + pag.pedido + ": " + tipoTransacao.GetCodOpLog() + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\nTipo de operação não previsto ao processar o registro do pagamento no pedido " + pag.pedido + ": " + tipoTransacao.GetCodOpLog();
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}

						return false;
					}

					if (tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.CaptureCreditCardTransaction.GetCodOpLog()))
					{
						#region [ Update nos campos que registram pagamento (crédito) ]
						updatePagPaymentPagtoRegPedido = new BraspagUpdatePagPaymentPagtoRegPedido();
						updatePagPaymentPagtoRegPedido.pagto_registrado_no_pedido_status = 1;
						updatePagPaymentPagtoRegPedido.id_pagto_gw_pag_payment = payment.id;
						updatePagPaymentPagtoRegPedido.pagto_registrado_no_pedido_tipo_operacao = opRegistraPagtoPedido.GetValue();
						updatePagPaymentPagtoRegPedido.pagto_registrado_no_pedido_id_pedido_pagamento = pedidoPagto.id;
						updatePagPaymentPagtoRegPedido.pagto_registrado_no_pedido_st_pagto_anterior = pedidoBase.st_pagto;
						updatePagPaymentPagtoRegPedido.pagto_registrado_no_pedido_st_pagto_novo = st_pagto_novo;
						if (!BraspagDAO.updatePagPaymentPagtoRegPedido(updatePagPaymentPagtoRegPedido, out msg_erro_aux))
						{
							// Retorna mensagem de erro p/ rotina chamadora
							msg_erro = msg_erro_aux;

							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = msg_erro;
							svcLog.complemento_1 = Global.serializaObjectToXml(payment);
							svcLog.complemento_2 = Global.serializaObjectToXml(updatePagPaymentPagtoRegPedido);
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar atualizar o registro em " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + payment.id.ToString() + " com as informações de que o pagamento foi registrado no pedido " + pag.pedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o registro em " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + payment.id.ToString() + " com as informações de que o pagamento foi registrado no pedido " + pag.pedido + "\r\nChamada à rotina BraspagDAO.updatePagPaymentPagtoRegPedido() retornou erro:\r\n" + msg_erro;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}

							return false;
						}

						s_log = "Registro automático de pagamento decorrente de operação de '" + opRegistraPagtoPedido.GetDescription() + "' na Braspag no valor de " + Global.formataMoeda(payment.valor_transacao) + " foi registrado com sucesso no pedido " + pag.pedido + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ", t_PEDIDO_PAGAMENTO.id=" + pedidoPagto.id + "): " + s_log;
						GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PEDIDO_PAGTO_CONTABILIZADO_BRASPAG_CLEARSALE, pag.pedido, s_log, out msg_erro_aux);
						#endregion
					}
					else
					{
						#region [ Update nos campos que registram estorno (débito) ]
						updatePagPaymentEstornoRegPedido = new BraspagUpdatePagPaymentEstornoRegPedido();
						updatePagPaymentEstornoRegPedido.estorno_registrado_no_pedido_status = 1;
						updatePagPaymentEstornoRegPedido.id_pagto_gw_pag_payment = payment.id;
						updatePagPaymentEstornoRegPedido.estorno_registrado_no_pedido_tipo_operacao = opRegistraPagtoPedido.GetValue();
						updatePagPaymentEstornoRegPedido.estorno_registrado_no_pedido_id_pedido_pagamento = pedidoPagto.id;
						updatePagPaymentEstornoRegPedido.estorno_registrado_no_pedido_st_pagto_anterior = pedidoBase.st_pagto;
						updatePagPaymentEstornoRegPedido.estorno_registrado_no_pedido_st_pagto_novo = st_pagto_novo;
						if (!BraspagDAO.updatePagPaymentEstornoRegPedido(updatePagPaymentEstornoRegPedido, out msg_erro_aux))
						{
							// Retorna mensagem de erro p/ rotina chamadora
							msg_erro = msg_erro_aux;

							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = msg_erro;
							svcLog.complemento_1 = Global.serializaObjectToXml(payment);
							svcLog.complemento_2 = Global.serializaObjectToXml(updatePagPaymentEstornoRegPedido);
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar atualizar o registro em " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + payment.id.ToString() + " com as informações de que o pagamento foi registrado no pedido " + pag.pedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o registro em " + Global.Cte.FIN.NSU.T_PAGTO_GW_PAG_PAYMENT + ".id=" + payment.id.ToString() + " com as informações de que o pagamento foi registrado no pedido " + pag.pedido + "\r\nChamada à rotina BraspagDAO.updatePagPaymentEstornoRegPedido() retornou erro:\r\n" + msg_erro;
							if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
							{
								strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
								Global.gravaLogAtividade(strMsg);
							}

							return false;
						}

						s_log = "Registro automático de estorno decorrente de operação de '" + opRegistraPagtoPedido.GetDescription() + "' na Braspag no valor de " + Global.formataMoeda(payment.valor_transacao) + " foi registrado com sucesso no pedido " + pag.pedido + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ", t_PEDIDO_PAGAMENTO.id=" + pedidoPagto.id + "): " + s_log;
						GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PEDIDO_PAGTO_CONTABILIZADO_BRASPAG_CLEARSALE, pag.pedido, s_log, out msg_erro_aux);
						#endregion
					}
					#endregion
				}
				#endregion

				#region [ Registra no histórico de pagamentos do pedido ]
				blnUpdateHistPagto = false;
				if (
					tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.CaptureCreditCardTransaction.GetCodOpLog())
					||
					(tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.VoidCreditCardTransaction.GetCodOpLog()) && (payment.captura_confirmada_status == 0))
					)
				{
					// Se é captura, tenta localizar o registro do histórico de pagamentos do pedido gerado durante a requisição de autorização p/ apenas atualizar o status
					histPagto = PedidoDAO.getPedidoHistPagtoByCtrlPagtoIdParcela(Global.Cte.FIN.CtrlPagtoModulo.BRASPAG_CLEARSALE, payment.id, out msg_erro_aux);
					if (histPagto != null)
					{
						if (histPagto.status == Global.Cte.T_FIN_PEDIDO_HIST_PAGTO__status.PREVISAO)
						{
							blnUpdateHistPagto = true;
							if (tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.CaptureCreditCardTransaction.GetCodOpLog()))
							{
								histPagto.status = Global.Cte.T_FIN_PEDIDO_HIST_PAGTO__status.QUITADO;
							}
							else if (tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.VoidCreditCardTransaction.GetCodOpLog()) && (payment.captura_confirmada_status == 0))
							{
								histPagto.status = Global.Cte.T_FIN_PEDIDO_HIST_PAGTO__status.CANCELADO;
							}
						}
					}
				}

				if (blnUpdateHistPagto)
				{
					#region [ Atualiza o status do registro no histórico de pagamentos do pedido ]
					if (!PedidoDAO.updateFinPedidoHistPagtoCampoStatus(histPagto, out msg_erro_aux))
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = Global.serializaObjectToXml(payment);
						svcLog.complemento_2 = Global.serializaObjectToXml(histPagto);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar atualizar os dados no histórico de pagamentos do pedido " + pag.pedido + " (" + Global.Cte.Nsu.T_FIN_PEDIDO_HIST_PAGTO + ".id=" + histPagto.id.ToString() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar os dados no histórico de pagamentos do pedido " + pag.pedido + " (" + Global.Cte.Nsu.T_FIN_PEDIDO_HIST_PAGTO + ".id=" + histPagto.id.ToString() + ", t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\nChamada à rotina PedidoDAO.updateFinPedidoHistPagtoCampoStatus() retornou erro:\r\n" + msg_erro;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}

						return false;
					}
					#endregion
				}
				else
				{
					#region [ Insere novo registro em t_FIN_PEDIDO_HIST_PAGTO ]
					histPagto = new PedidoHistPagto();
					histPagto.pedido = pag.pedido;

					#region [ Status e descrição ]
					histPagto.descricao = Global.Cte.Braspag.Bandeira.GetDescription(payment.bandeira) + ": " + Global.formataMoeda(payment.valor_transacao);
					if (tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.CaptureCreditCardTransaction.GetCodOpLog())
						||
						tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.AuthorizeTransaction.GetCodOpLog()))
					{
						histPagto.descricao += " em " + payment.req_PaymentDataRequest_NumberOfPayments.Trim() + "x";
					}

					if (tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.CaptureCreditCardTransaction.GetCodOpLog()))
					{
						statusHistPagto = Global.Cte.T_FIN_PEDIDO_HIST_PAGTO__status.QUITADO;
					}
					else if (tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.AuthorizeTransaction.GetCodOpLog()))
					{
						statusHistPagto = Global.Cte.T_FIN_PEDIDO_HIST_PAGTO__status.PREVISAO;
					}
					else if (tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.RefundCreditCardTransaction.GetCodOpLog()))
					{
						statusHistPagto = Global.Cte.T_FIN_PEDIDO_HIST_PAGTO__status.CANCELADO;
						histPagto.descricao = "(" + opRegistraPagtoPedido.GetDescription() + ") " + histPagto.descricao;
					}
					else if (tipoTransacao.GetCodOpLog().Equals(Global.Cte.Braspag.Pagador.Transacao.VoidCreditCardTransaction.GetCodOpLog()) && (payment.captura_confirmada_status == 1))
					{
						statusHistPagto = Global.Cte.T_FIN_PEDIDO_HIST_PAGTO__status.CANCELADO;
						histPagto.descricao = "(" + opRegistraPagtoPedido.GetDescription() + ") " + histPagto.descricao;
					}
					else
					{
						statusHistPagto = 0;

						#region [ Registra log devido a situação inesperada ]
						msg_erro = "Situação não prevista no processamento do histórico de pagamentos do pedido (insert): t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ", tipo de operação=" + tipoTransacao.GetCodOpLog() + ", t_PAGTO_GW_PAG_PAYMENT.captura_confirmada_status=" + payment.captura_confirmada_status.ToString();

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = Global.serializaObjectToXml(payment);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar incluir novo registro no histórico de pagamentos do pedido " + pag.pedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\nFalha ao tentar incluir novo registro no histórico de pagamentos do pedido " + pag.pedido + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\n" + msg_erro;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}
						#endregion
					}
					#endregion

					if (statusHistPagto != 0)
					{
						histPagto.status = (byte)statusHistPagto;
						histPagto.ctrl_pagto_id_parcela = payment.id;
						histPagto.ctrl_pagto_modulo = Global.Cte.FIN.CtrlPagtoModulo.BRASPAG_CLEARSALE;
						histPagto.valor_total = payment.valor_transacao;
						histPagto.valor_rateado = payment.valor_transacao;

						if (!PedidoDAO.insereFinPedidoHistPagto(histPagto, out msg_erro_aux))
						{
							// Retorna mensagem de erro p/ rotina chamadora
							msg_erro = msg_erro_aux;

							#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
							FinSvcLog svcLog = new FinSvcLog();
							svcLog.operacao = NOME_DESTA_ROTINA;
							svcLog.descricao = msg_erro;
							svcLog.complemento_1 = Global.serializaObjectToXml(payment);
							svcLog.complemento_2 = Global.serializaObjectToXml(histPagto);
							GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
							#endregion

							strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar gravar dados no histórico de pagamentos do pedido " + pag.pedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
							strBody = "Mensagem de Financeiro Service\nFalha ao tentar gravar dados no histórico de pagamentos do pedido " + pag.pedido + " (t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\nChamada à rotina PedidoDAO.insereFinPedidoHistPagto() retornou erro:\r\n" + msg_erro;
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
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = tipoTransacao.GetCodOpLog();
				svcLog.complemento_2 = Global.serializaObjectToXml(payment);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar registrar o pagamento no pedido (operação: " + tipoTransacao.GetCodOpLog() + ") [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
				strBody = "Mensagem de Financeiro Service\nFalha ao tentar registrar o pagamento no pedido (operação " + tipoTransacao.GetCodOpLog() + ", t_PAGTO_GW_PAG_PAYMENT.id=" + payment.id.ToString() + ")\r\n" + msg_erro_aux;
				if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
				{
					strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
					Global.gravaLogAtividade(strMsg);
				}

				return false;
			}
		}
		#endregion

		#region [ registraPagamentoBoletoECNoPedido ]
		public static bool registraPagamentoBoletoECNoPedido(int id_braspag_webhook_complementar, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "BraspagDAO.registraPagamentoBoletoECNoPedido()";
			int id_emailsndsvc_mensagem;
			int? st_analise_credito_novo = null;
			Decimal vlTotalFamiliaPrecoVenda;
			Decimal vlTotalFamiliaPrecoNF;
			Decimal vlTotalFamiliaPago;
			Decimal vlTotalFamiliaDevolucaoPrecoVenda;
			Decimal vlTotalFamiliaDevolucaoPrecoNF;
			Decimal vlTotalFamiliaPagoNovo;
			string msg_erro_aux;
			string strSubject;
			string strBody;
			string strMsg;
			string id_pedido_base;
			string st_pagto;
			string st_pagto_novo = "";
			string s_log = "";
			BraspagWebhook braspagWebhook = null;
			BraspagWebhookComplementar braspagWebhookComplementar = null;
			PedidoPagamento pedidoPagto;
			Pedido pedidoBase;
			Global.Cte.Braspag.Pagador.OperacaoRegistraPagtoPedido opRegistraPagtoPedido = null;
			BraspagWebhookComplementarUpdatePagtoRegPedido updatePagtoRegPedido;
			PedidoHistPagto histPagto = null;
			#endregion

			msg_erro = "";
			try
			{
				braspagWebhookComplementar = getBraspagWebhookComplementarById(id_braspag_webhook_complementar, out msg_erro_aux);
				if (braspagWebhookComplementar == null)
				{
					msg_erro = "Falha ao recuperar os dados complementares do Webhook Braspag (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + id_braspag_webhook_complementar.ToString() + ")" + (msg_erro_aux.Length > 0 ? ": " + msg_erro_aux : "");
					return false;
				}

				braspagWebhook = getBraspagWebhookById(braspagWebhookComplementar.id_braspag_webhook, out msg_erro_aux);
				if (braspagWebhook == null)
				{
					msg_erro = "Falha ao recuperar os dados gravados pelo Webhook Braspag (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + braspagWebhookComplementar.id_braspag_webhook.ToString() + ")" + (msg_erro_aux.Length > 0 ? ": " + msg_erro_aux : "");
					return false;
				}

				if ((braspagWebhookComplementar.pedido ?? "").Length == 0)
				{
					msg_erro = "Não é possível registrar o pagamento no pedido porque não foi identificado o número do pedido no ERP a partir do OrderId da Braspag (OrderId=" + braspagWebhook.NumPedido + ")";
					return false;
				}

				id_pedido_base = Global.retornaNumeroPedidoBase(braspagWebhookComplementar.pedido);

				#region [ Calcula valores do pedido ]
				if (!PedidoDAO.calculaPagamentos(braspagWebhookComplementar.pedido, out vlTotalFamiliaPrecoVenda, out vlTotalFamiliaPrecoNF, out vlTotalFamiliaPago, out vlTotalFamiliaDevolucaoPrecoVenda, out vlTotalFamiliaDevolucaoPrecoNF, out st_pagto, out msg_erro_aux))
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = msg_erro_aux;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					strMsg = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + id_braspag_webhook_complementar.ToString();
					if (braspagWebhookComplementar != null) strMsg = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + braspagWebhookComplementar.id_braspag_webhook.ToString() + ", " + strMsg;
					svcLog.complemento_1 = strMsg;
					svcLog.complemento_2 = Global.serializaObjectToXml(braspagWebhook);
					svcLog.complemento_3 = Global.serializaObjectToXml(braspagWebhookComplementar);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Webhook Braspag: Falha ao tentar calcular valores do pedido durante operação de registrar o pagamento no pedido [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\r\nWebhook Braspag\r\nFalha ao tentar calcular valores do pedido durante operação de registrar o pagamento no pedido " + braspagWebhookComplementar.pedido + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + id_braspag_webhook_complementar.ToString() + ")\r\nChamada à rotina PedidoDAO.calculaPagamentos() retornou erro:\r\n" + msg_erro;
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
					}

					return false;
				}
				#endregion

				#region [ Registra o pagamento no pedido ]
				// Esta rotina é específica para registrar pagamento de boleto de e-commerce (ex: Bradesco SPS)
				if ((!braspagWebhook.CODPAGAMENTO.Equals(Global.Cte.Braspag.PaymentMethod.Boleto_Bradesco_SPS.GetValue())) && (!braspagWebhook.CODPAGAMENTO.Equals(Global.Cte.Braspag.PaymentMethod.Boleto_Registrado_Bradesco.GetValue())))
				{
					msg_erro = "Não foi possível registrar o pagamento no pedido porque o meio de pagamento desta transação não é boleto de e-commerce (PaymentMethod=" + braspagWebhook.CODPAGAMENTO + " - " + Global.Cte.Braspag.PaymentMethod.GetDescription(braspagWebhook.CODPAGAMENTO) + ")";
					return false;
				}

				if (!braspagWebhookComplementar.GlobalStatus.Equals(Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue()))
				{
					msg_erro = "Não foi possível registrar o pagamento no pedido porque o status da transação não é 'capturada' (" + braspagWebhookComplementar.GlobalStatus + " - " + Global.Cte.Braspag.Pagador.GlobalStatus.GetDescription(braspagWebhookComplementar.GlobalStatus) + ")";
					return false;
				}

				#region [ Grava pagamento no pedido ]
				pedidoPagto = new PedidoPagamento();
				pedidoPagto.pedido = braspagWebhookComplementar.pedido;
				pedidoPagto.valor = braspagWebhookComplementar.ValorPaidAmount;
				pedidoPagto.tipo_pagto = Global.Cte.PedidoPagtoTipoOperacao.BRASPAG_WEBHOOK;
				pedidoPagto.usuario = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
				pedidoPagto.id_braspag_webhook_complementar = braspagWebhookComplementar.Id;
				if (!PedidoDAO.inserePedidoPagamentoBoletoEC(pedidoPagto, out msg_erro_aux))
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = msg_erro_aux;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(braspagWebhook);
					svcLog.complemento_2 = Global.serializaObjectToXml(braspagWebhookComplementar);
					svcLog.complemento_3 = Global.serializaObjectToXml(pedidoPagto);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar registrar o pagamento no pedido [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\nFalha ao tentar registrar o pagamento no pedido " + braspagWebhookComplementar.pedido + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + braspagWebhookComplementar.Id.ToString() + ")\r\nChamada à rotina PedidoDAO.inserePedidoPagamentoBoletoEC() retornou erro:\r\n" + msg_erro;
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
					}

					return false;
				}
				#endregion

				#region [ Determina alterações no status de pagamento do pedido (Quitado, Pago Parcial, Não-Pago) e no status de análise de crédito ]
				pedidoBase = PedidoDAO.getPedido(id_pedido_base);
				if ((vlTotalFamiliaDevolucaoPrecoNF + vlTotalFamiliaPago + pedidoPagto.valor) >= (vlTotalFamiliaPrecoNF - Global.Cte.Etc.MAX_VALOR_MARGEM_ERRO_PAGAMENTO))
				{
					#region [ Pago (Quitado) ]
					st_pagto_novo = Global.Cte.StPagtoPedido.ST_PAGTO_PAGO;
					s_log += "Status de pagamento do pedido: quitado (st_pagto: '" + pedidoBase.st_pagto + "' => '" + st_pagto_novo + "')";
					if ((vlTotalFamiliaDevolucaoPrecoNF + vlTotalFamiliaPago + pedidoPagto.valor) > vlTotalFamiliaPrecoNF)
					{
						s_log += " (excedeu " + Global.Cte.Etc.SIMBOLO_MONETARIO + " " + Global.formataMoeda((vlTotalFamiliaDevolucaoPrecoNF + vlTotalFamiliaPago + pedidoPagto.valor) - vlTotalFamiliaPrecoNF) + ")";
					}
					else if ((vlTotalFamiliaDevolucaoPrecoNF + vlTotalFamiliaPago + pedidoPagto.valor) < vlTotalFamiliaPrecoNF)
					{
						s_log += " (faltou " + Global.Cte.Etc.SIMBOLO_MONETARIO + " " + Global.formataMoeda(vlTotalFamiliaPrecoNF - (vlTotalFamiliaDevolucaoPrecoNF + vlTotalFamiliaPago + pedidoPagto.valor)) + ")";
					}

					#region [ Status da análise de crédito ]
					if (
							(
								(pedidoBase.analise_credito == (short)Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.ST_INICIAL)
								||
								(pedidoBase.analise_credito == (short)Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS)
							)
							&&
							(pedidoBase.analise_credito != (short)Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK)
						)
					{
						st_analise_credito_novo = Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK;
						if (s_log.Length > 0) s_log += "; ";
						s_log += " Análise de crédito: " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " => " + Global.obtemDescricaoAnaliseCredito((int)st_analise_credito_novo) + " (" + braspagWebhook.CODPAGAMENTO + " - " + Global.Cte.Braspag.PaymentMethod.GetDescription(braspagWebhook.CODPAGAMENTO) + ")";
					}
					else
					{
						if (s_log.Length > 0) s_log += "; ";
						s_log += " Análise de crédito: status não foi alterado porque pedido encontra-se em " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " (" + braspagWebhook.CODPAGAMENTO + " - " + Global.Cte.Braspag.PaymentMethod.GetDescription(braspagWebhook.CODPAGAMENTO) + ")";
					}
					#endregion

					#endregion
				}
				else if ((vlTotalFamiliaPago + pedidoPagto.valor) > 0)
				{
					#region [ Pagamento Parcial ]
					st_pagto_novo = Global.Cte.StPagtoPedido.ST_PAGTO_PARCIAL;
					s_log += "Status de pagamento do pedido: pago parcial (st_pagto: '" + pedidoBase.st_pagto + "' => '" + st_pagto_novo + "')";

					if (
							(pedidoBase.analise_credito != Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS)
							&&
							(pedidoBase.analise_credito != (short)Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK)
						)
					{
						st_analise_credito_novo = Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS;
						if (s_log.Length > 0) s_log += "; ";
						s_log += " Análise de crédito: " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " => " + Global.obtemDescricaoAnaliseCredito((int)st_analise_credito_novo) + " (" + braspagWebhook.CODPAGAMENTO + " - " + Global.Cte.Braspag.PaymentMethod.GetDescription(braspagWebhook.CODPAGAMENTO) + ")";
					}
					else
					{
						if (s_log.Length > 0) s_log += "; ";
						s_log += " Análise de crédito: status não foi alterado porque pedido encontra-se em " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " (" + braspagWebhook.CODPAGAMENTO + " - " + Global.Cte.Braspag.PaymentMethod.GetDescription(braspagWebhook.CODPAGAMENTO) + ")";
					}
					#endregion
				}
				else
				{
					#region [ Não-Pago ]
					st_pagto_novo = Global.Cte.StPagtoPedido.ST_PAGTO_NAO_PAGO;
					s_log += "Status de pagamento do pedido: não-pago (st_pagto: '" + pedidoBase.st_pagto + "' => '" + st_pagto_novo + "')";

					if (
							(pedidoBase.analise_credito != Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS)
							&&
							(pedidoBase.analise_credito != (short)Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.CREDITO_OK)
						)
					{
						st_analise_credito_novo = Global.Cte.T_PEDIDO__ANALISE_CREDITO_STATUS.PENDENTE_VENDAS;
						if (s_log.Length > 0) s_log += "; ";
						s_log += " Análise de crédito: " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " => " + Global.obtemDescricaoAnaliseCredito((int)st_analise_credito_novo) + " (" + braspagWebhook.CODPAGAMENTO + " - " + Global.Cte.Braspag.PaymentMethod.GetDescription(braspagWebhook.CODPAGAMENTO) + ")";
					}
					else
					{
						if (s_log.Length > 0) s_log += "; ";
						s_log += " Análise de crédito: status não foi alterado porque pedido encontra-se em " + Global.obtemDescricaoAnaliseCredito(pedidoBase.analise_credito) + " (" + braspagWebhook.CODPAGAMENTO + " - " + Global.Cte.Braspag.PaymentMethod.GetDescription(braspagWebhook.CODPAGAMENTO) + ")";
					}
					#endregion
				}
				#endregion

				#region [ Atualiza status de pagamento do pedido ]
				if (st_pagto_novo.Length > 0)
				{
					if (!PedidoDAO.atualizaPedidoStatusPagto(id_pedido_base, st_pagto_novo, out msg_erro_aux))
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = Global.serializaObjectToXml(braspagWebhook);
						svcLog.complemento_2 = Global.serializaObjectToXml(braspagWebhookComplementar);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar atualizar o status de pagamento do pedido " + id_pedido_base + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o status de pagamento do pedido " + id_pedido_base + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + braspagWebhookComplementar.Id.ToString() + ")\r\nChamada à rotina PedidoDAO.atualizaPedidoStatusPagto() retornou erro:\r\n" + msg_erro;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}

						return false;
					}
				}
				#endregion

				#region [ Atualiza status da análise de crédito do pedido ]
				if (st_analise_credito_novo != null)
				{
					if (!PedidoDAO.atualizaPedidoStatusAnaliseCredito(id_pedido_base, st_analise_credito_novo, Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA, out msg_erro_aux))
					{
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = msg_erro_aux;

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = msg_erro;
						svcLog.complemento_1 = Global.serializaObjectToXml(braspagWebhook);
						svcLog.complemento_2 = Global.serializaObjectToXml(braspagWebhookComplementar);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion

						strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar atualizar o status da análise de crédito do pedido " + id_pedido_base + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
						strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o status da análise de crédito do pedido " + id_pedido_base + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + braspagWebhookComplementar.Id.ToString() + ")\r\nChamada à rotina PedidoDAO.atualizaPedidoStatusAnaliseCredito() retornou erro:\r\n" + msg_erro;
						if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
						{
							strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
							Global.gravaLogAtividade(strMsg);
						}

						return false;
					}
				}
				#endregion

				#region [ Atualiza o campo 'vl_pago_familia' no pedido ]
				vlTotalFamiliaPagoNovo = vlTotalFamiliaPago + pedidoPagto.valor;
				if (!PedidoDAO.atualizaPedidoVlPagoFamilia(id_pedido_base, vlTotalFamiliaPagoNovo, out msg_erro_aux))
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = msg_erro_aux;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(braspagWebhook);
					svcLog.complemento_2 = Global.serializaObjectToXml(braspagWebhookComplementar);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar atualizar o valor total pago (família de pedidos) do pedido " + id_pedido_base + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o valor total pago (família de pedidos) do pedido " + id_pedido_base + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + braspagWebhookComplementar.Id.ToString() + ")\r\nChamada à rotina PedidoDAO.atualizaPedidoVlPagoFamilia() retornou erro:\r\n" + msg_erro;
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
					}

					return false;
				}
				#endregion

				#region [ Registra em t_BRASPAG_WEBHOOK_COMPLEMENTAR que foi gerado pagamento/estorno + Grava log ]
				// Atualiza t_BRASPAG_WEBHOOK_COMPLEMENTAR os campos que indicam se o pagamento foi registrado no pedido
				opRegistraPagtoPedido = Global.Cte.Braspag.Pagador.OperacaoRegistraPagtoPedido.CAPTURA;

				#region [ Update nos campos que registram pagamento ]
				updatePagtoRegPedido = new BraspagWebhookComplementarUpdatePagtoRegPedido();
				updatePagtoRegPedido.PagtoRegistradoNoPedidoStatus = 1;
				updatePagtoRegPedido.id_braspag_webhook_complementar = braspagWebhookComplementar.Id;
				updatePagtoRegPedido.PagtoRegistradoNoPedidoTipoOperacao = opRegistraPagtoPedido.GetValue();
				updatePagtoRegPedido.PagtoRegistradoNoPedido_id_pedido_pagamento= pedidoPagto.id;
				updatePagtoRegPedido.PagtoRegistradoNoPedidoStPagtoAnterior = pedidoBase.st_pagto;
				updatePagtoRegPedido.PagtoRegistradoNoPedidoStPagtoNovo = st_pagto_novo;
				updatePagtoRegPedido.AnaliseCreditoStatusAnterior = pedidoBase.analise_credito;
				if (st_analise_credito_novo != null)
				{
					updatePagtoRegPedido.AnaliseCreditoStatusNovo = (int)st_analise_credito_novo;
				}
				else
				{
					updatePagtoRegPedido.AnaliseCreditoStatusNovo = pedidoBase.analise_credito;
				}

				if (!BraspagDAO.updateWebhookComplementarPagtoRegPedido(updatePagtoRegPedido, out msg_erro_aux))
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = msg_erro_aux;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(braspagWebhook);
					svcLog.complemento_2 = Global.serializaObjectToXml(braspagWebhookComplementar);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar atualizar o registro em " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + braspagWebhookComplementar.Id.ToString() + " com as informações de que o pagamento foi registrado no pedido " + braspagWebhookComplementar.pedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\nFalha ao tentar atualizar o registro em " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + braspagWebhookComplementar.Id.ToString() + " com as informações de que o pagamento foi registrado no pedido " + braspagWebhookComplementar.pedido + "\r\nChamada à rotina BraspagDAO.updateWebhookComplementarPagtoRegPedido() retornou erro:\r\n" + msg_erro;
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
					}

					return false;
				}

				s_log = "Registro automático de pagamento decorrente de operação de '" + opRegistraPagtoPedido.GetDescription() + "' (" + braspagWebhook.CODPAGAMENTO + " - " + Global.Cte.Braspag.PaymentMethod.GetDescription(braspagWebhook.CODPAGAMENTO) + ") na Braspag no valor de " + Global.formataMoeda(braspagWebhookComplementar.ValorPaidAmount) + " foi registrado com sucesso no pedido " + braspagWebhookComplementar.pedido + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + braspagWebhookComplementar.Id.ToString() + ", t_PEDIDO_PAGAMENTO.id=" + pedidoPagto.id + "): " + s_log;
				GeralDAO.gravaLog(Global.Cte.LogBd.Operacao.OP_LOG_PEDIDO_PAGTO_CONTABILIZADO_BRASPAG_WEBHOOK, braspagWebhookComplementar.pedido, s_log, out msg_erro_aux);
				#endregion

				#endregion

				#endregion

				#region [ Registra no histórico de pagamentos do pedido ]
				histPagto = new PedidoHistPagto();
				histPagto.pedido = braspagWebhookComplementar.pedido;
				histPagto.status = (byte)Global.Cte.T_FIN_PEDIDO_HIST_PAGTO__status.QUITADO;
				histPagto.ctrl_pagto_id_parcela = braspagWebhookComplementar.Id;
				histPagto.ctrl_pagto_modulo = Global.Cte.FIN.CtrlPagtoModulo.BRASPAG_WEBHOOK;
				histPagto.dt_vencto = (braspagWebhookComplementar.BoletoExpirationDate ?? DateTime.MinValue);
				histPagto.dt_credito = (braspagWebhookComplementar.CapturedDate ?? DateTime.MinValue);
				histPagto.valor_total = braspagWebhookComplementar.ValorAmount;
				histPagto.valor_rateado = braspagWebhookComplementar.ValorAmount;
				histPagto.valor_pago = braspagWebhookComplementar.ValorPaidAmount;
				histPagto.descricao = Global.Cte.Braspag.PaymentMethod.GetDescription(braspagWebhook.CODPAGAMENTO) +
										" (Emissão: " + Global.formataDataDdMmYyyyComSeparador(braspagWebhookComplementar.ReceivedDate) + ")";
				if (!PedidoDAO.insereFinPedidoHistPagtoBoletoEC(histPagto, out msg_erro_aux))
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = msg_erro_aux;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(braspagWebhook);
					svcLog.complemento_2 = Global.serializaObjectToXml(braspagWebhookComplementar);
					svcLog.complemento_3 = Global.serializaObjectToXml(histPagto);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar gravar dados no histórico de pagamentos do pedido " + braspagWebhookComplementar.pedido + " [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
					strBody = "Mensagem de Financeiro Service\nFalha ao tentar gravar dados no histórico de pagamentos do pedido " + braspagWebhookComplementar.pedido + " (" + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + braspagWebhookComplementar.Id.ToString() + ")\r\nChamada à rotina PedidoDAO.insereFinPedidoHistPagtoBoletoEC() retornou erro:\r\n" + msg_erro;
					if (!EmailSndSvcDAO.gravaMensagemParaEnvio(Global.Cte.Clearsale.Email.REMETENTE_MSG_ALERTA_SISTEMA, Global.Cte.Clearsale.Email.DESTINATARIO_MSG_ALERTA_SISTEMA, null, null, strSubject, strBody, DateTime.Now, out id_emailsndsvc_mensagem, out msg_erro_aux))
					{
						strMsg = NOME_DESTA_ROTINA + ": Falha ao tentar inserir email de alerta na fila de mensagens!!\n" + msg_erro_aux;
						Global.gravaLogAtividade(strMsg);
					}

					return false;
				}
				#endregion

				return true;
			}
			catch (Exception ex)
			{
				msg_erro = ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				strMsg = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + id_braspag_webhook_complementar.ToString();
				if (braspagWebhookComplementar != null) strMsg = Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + braspagWebhookComplementar.id_braspag_webhook.ToString() + ", " + strMsg;
				svcLog.complemento_1 = strMsg;
				svcLog.complemento_2 = Global.serializaObjectToXml(braspagWebhook);
				svcLog.complemento_3 = Global.serializaObjectToXml(braspagWebhookComplementar);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				strSubject = Global.montaIdInstanciaServicoEmailSubject() + " Braspag: Falha ao tentar registrar o pagamento do boleto de e-commerce no pedido [" + Global.formataDataDdMmYyyyHhMmSsComSeparador(DateTime.Now) + "]";
				strBody = "Mensagem de Financeiro Service\nFalha ao tentar registrar o pagamento do boleto de e-commerce no pedido (pedido=" + (braspagWebhook != null ? braspagWebhook.NumPedido : "(undefined)") + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK + ".Id=" + (braspagWebhook != null ? braspagWebhook.Id.ToString() : "(undefined)") + ", " + Global.Cte.FIN.NSU.T_BRASPAG_WEBHOOK_COMPLEMENTAR + ".Id=" + (braspagWebhookComplementar != null ? braspagWebhookComplementar.Id.ToString() : "(undefined)") + ")\r\n" + msg_erro_aux;
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
