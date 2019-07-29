using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace FinanceiroService
{
	class ClearsaleDAO
	{
		#region [ Atributos ]
		private static SqlCommand cmInsertAF;
		private static SqlCommand cmInsertAFItem;
		private static SqlCommand cmInsertAFPayment;
		private static SqlCommand cmInsertAFPhone;
		private static SqlCommand cmInsertAFXml;
		private static SqlCommand cmInsertAFNsu;
		private static SqlCommand cmInsertAFOpComplementar;
		private static SqlCommand cmInsertAFOpComplementarXml;
		private static SqlCommand cmUpdateAFAnulaRegistroTentativaAnterior;
		private static SqlCommand cmUpdateAFErroRx;
		private static SqlCommand cmUpdateAFIdRegistroXmlTx;
		private static SqlCommand cmUpdateAFIdRegistroXmlRx;
		private static SqlCommand cmUpdatePagPaymentStEnviadoAnaliseAF;
		private static SqlCommand cmUpdateAFNsu;
		private static SqlCommand cmUpdateAFSendOrdersResponse;
		private static SqlCommand cmUpdateAFGetReturnAnalysisResponse;
		private static SqlCommand cmUpdateAFOpComplementar;
		private static SqlCommand cmUpdateAFSetOrderAsReturnedPendente;
		private static SqlCommand cmUpdateAFSetOrderAsReturnedSucesso;
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
		static ClearsaleDAO()
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

			#region [ cmInsertAF ]
			strSql = "INSERT INTO t_PAGTO_GW_AF (" +
						"id, " +
						"usuario, " +
						"owner, " +
						"loja, " +
						"id_cliente, " +
						"pedido, " +
						"pedido_com_sufixo_nsu, " +
						"valor_pedido, " +
						"req_entityCode, " +
						"req_Order_ID, " +
						"req_Order_FingerPrint_SessionID, " +
						"req_Order_Date, " +
						"req_Order_Email, " +
						"req_Order_B2B_B2C, " +
						"req_Order_ShippingPrice, " +
						"req_Order_TotalItems, " +
						"req_Order_TotalOrder, " +
						"req_Order_QtyInstallments, " +
						"req_Order_DeliveryTimeCD, " +
						"req_Order_QtyItems, " +
						"req_Order_QtyPaymentTypes, " +
						"req_Order_IP, " +
						"req_Order_Status, " +
						"req_Order_Reanalise, " +
						"req_Order_Origin, " +
						"req_Order_BillingData_ID, " +
						"req_Order_BillingData_Type, " +
						"req_Order_BillingData_LegalDocument1, " +
						"req_Order_BillingData_LegalDocument2, " +
						"req_Order_BillingData_Name, " +
						"req_Order_BillingData_BirthDate, " +
						"req_Order_BillingData_Email, " +
						"req_Order_BillingData_Gender, " +
						"req_Order_BillingData_Address_Street, " +
						"req_Order_BillingData_Address_Number, " +
						"req_Order_BillingData_Address_Comp, " +
						"req_Order_BillingData_Address_County, " +
						"req_Order_BillingData_Address_City, " +
						"req_Order_BillingData_Address_State, " +
						"req_Order_BillingData_Address_Country, " +
						"req_Order_BillingData_Address_ZipCode, " +
						"req_Order_BillingData_Address_Reference, " +
						"req_Order_ShippingData_ID, " +
						"req_Order_ShippingData_Type, " +
						"req_Order_ShippingData_LegalDocument1, " +
						"req_Order_ShippingData_LegalDocument2, " +
						"req_Order_ShippingData_Name, " +
						"req_Order_ShippingData_BirthDate, " +
						"req_Order_ShippingData_Email, " +
						"req_Order_ShippingData_Gender, " +
						"req_Order_ShippingData_Address_Street, " +
						"req_Order_ShippingData_Address_Number, " +
						"req_Order_ShippingData_Address_Comp, " +
						"req_Order_ShippingData_Address_County, " +
						"req_Order_ShippingData_Address_City, " +
						"req_Order_ShippingData_Address_State, " +
						"req_Order_ShippingData_Address_Country, " +
						"req_Order_ShippingData_Address_ZipCode, " +
						"req_Order_ShippingData_Address_Reference" +
					") VALUES (" +
						"@id, " +
						"@usuario, " +
						"@owner, " +
						"@loja, " +
						"@id_cliente, " +
						"@pedido, " +
						"@pedido_com_sufixo_nsu, " +
						"@valor_pedido, " +
						"@req_entityCode, " +
						"@req_Order_ID, " +
						"@req_Order_FingerPrint_SessionID, " +
						"@req_Order_Date, " +
						"@req_Order_Email, " +
						"@req_Order_B2B_B2C, " +
						"@req_Order_ShippingPrice, " +
						"@req_Order_TotalItems, " +
						"@req_Order_TotalOrder, " +
						"@req_Order_QtyInstallments, " +
						"@req_Order_DeliveryTimeCD, " +
						"@req_Order_QtyItems, " +
						"@req_Order_QtyPaymentTypes, " +
						"@req_Order_IP, " +
						"@req_Order_Status, " +
						"@req_Order_Reanalise, " +
						"@req_Order_Origin, " +
						"@req_Order_BillingData_ID, " +
						"@req_Order_BillingData_Type, " +
						"@req_Order_BillingData_LegalDocument1, " +
						"@req_Order_BillingData_LegalDocument2, " +
						"@req_Order_BillingData_Name, " +
						"@req_Order_BillingData_BirthDate, " +
						"@req_Order_BillingData_Email, " +
						"@req_Order_BillingData_Gender, " +
						"@req_Order_BillingData_Address_Street, " +
						"@req_Order_BillingData_Address_Number, " +
						"@req_Order_BillingData_Address_Comp, " +
						"@req_Order_BillingData_Address_County, " +
						"@req_Order_BillingData_Address_City, " +
						"@req_Order_BillingData_Address_State, " +
						"@req_Order_BillingData_Address_Country, " +
						"@req_Order_BillingData_Address_ZipCode, " +
						"@req_Order_BillingData_Address_Reference, " +
						"@req_Order_ShippingData_ID, " +
						"@req_Order_ShippingData_Type, " +
						"@req_Order_ShippingData_LegalDocument1, " +
						"@req_Order_ShippingData_LegalDocument2, " +
						"@req_Order_ShippingData_Name, " +
						"@req_Order_ShippingData_BirthDate, " +
						"@req_Order_ShippingData_Email, " +
						"@req_Order_ShippingData_Gender, " +
						"@req_Order_ShippingData_Address_Street, " +
						"@req_Order_ShippingData_Address_Number, " +
						"@req_Order_ShippingData_Address_Comp, " +
						"@req_Order_ShippingData_Address_County, " +
						"@req_Order_ShippingData_Address_City, " +
						"@req_Order_ShippingData_Address_State, " +
						"@req_Order_ShippingData_Address_Country, " +
						"@req_Order_ShippingData_Address_ZipCode, " +
						"@req_Order_ShippingData_Address_Reference" +
					")";
			cmInsertAF = BD.criaSqlCommand();
			cmInsertAF.CommandText = strSql;
			cmInsertAF.Parameters.Add("@id", SqlDbType.Int);
			cmInsertAF.Parameters.Add("@usuario", SqlDbType.VarChar, 10);
			cmInsertAF.Parameters.Add("@owner", SqlDbType.SmallInt);
			cmInsertAF.Parameters.Add("@loja", SqlDbType.VarChar, 3);
			cmInsertAF.Parameters.Add("@id_cliente", SqlDbType.VarChar, 12);
			cmInsertAF.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmInsertAF.Parameters.Add("@pedido_com_sufixo_nsu", SqlDbType.VarChar, 15);
			cmInsertAF.Parameters.Add("@valor_pedido", SqlDbType.Money);
			cmInsertAF.Parameters.Add("@req_entityCode", SqlDbType.VarChar, 72);
			cmInsertAF.Parameters.Add("@req_Order_ID", SqlDbType.VarChar, 50);
			cmInsertAF.Parameters.Add("@req_Order_FingerPrint_SessionID", SqlDbType.VarChar, 128);
			cmInsertAF.Parameters.Add("@req_Order_Date", SqlDbType.VarChar, 19);
			cmInsertAF.Parameters.Add("@req_Order_Email", SqlDbType.VarChar, 150);
			cmInsertAF.Parameters.Add("@req_Order_B2B_B2C", SqlDbType.VarChar, 3);
			cmInsertAF.Parameters.Add("@req_Order_ShippingPrice", SqlDbType.VarChar, 24);
			cmInsertAF.Parameters.Add("@req_Order_TotalItems", SqlDbType.VarChar, 24);
			cmInsertAF.Parameters.Add("@req_Order_TotalOrder", SqlDbType.VarChar, 24);
			cmInsertAF.Parameters.Add("@req_Order_QtyInstallments", SqlDbType.VarChar, 6);
			cmInsertAF.Parameters.Add("@req_Order_DeliveryTimeCD", SqlDbType.VarChar, 50);
			cmInsertAF.Parameters.Add("@req_Order_QtyItems", SqlDbType.VarChar, 6);
			cmInsertAF.Parameters.Add("@req_Order_QtyPaymentTypes", SqlDbType.VarChar, 6);
			cmInsertAF.Parameters.Add("@req_Order_IP", SqlDbType.VarChar, 50);
			cmInsertAF.Parameters.Add("@req_Order_Status", SqlDbType.VarChar, 2);
			cmInsertAF.Parameters.Add("@req_Order_Reanalise", SqlDbType.VarChar, 2);
			cmInsertAF.Parameters.Add("@req_Order_Origin", SqlDbType.VarChar, 150);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_ID", SqlDbType.VarChar, 50);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_Type", SqlDbType.VarChar, 1);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_LegalDocument1", SqlDbType.VarChar, 100);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_LegalDocument2", SqlDbType.VarChar, 100);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_Name", SqlDbType.VarChar, 500);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_BirthDate", SqlDbType.VarChar, 19);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_Email", SqlDbType.VarChar, 150);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_Gender", SqlDbType.VarChar, 1);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_Address_Street", SqlDbType.VarChar, 200);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_Address_Number", SqlDbType.VarChar, 15);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_Address_Comp", SqlDbType.VarChar, 250);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_Address_County", SqlDbType.VarChar, 150);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_Address_City", SqlDbType.VarChar, 150);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_Address_State", SqlDbType.VarChar, 2);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_Address_Country", SqlDbType.VarChar, 150);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_Address_ZipCode", SqlDbType.VarChar, 10);
			cmInsertAF.Parameters.Add("@req_Order_BillingData_Address_Reference", SqlDbType.VarChar, 250);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_ID", SqlDbType.VarChar, 50);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_Type", SqlDbType.VarChar, 1);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_LegalDocument1", SqlDbType.VarChar, 100);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_LegalDocument2", SqlDbType.VarChar, 100);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_Name", SqlDbType.VarChar, 500);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_BirthDate", SqlDbType.VarChar, 19);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_Email", SqlDbType.VarChar, 150);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_Gender", SqlDbType.VarChar, 1);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_Address_Street", SqlDbType.VarChar, 200);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_Address_Number", SqlDbType.VarChar, 15);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_Address_Comp", SqlDbType.VarChar, 250);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_Address_County", SqlDbType.VarChar, 150);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_Address_City", SqlDbType.VarChar, 150);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_Address_State", SqlDbType.VarChar, 2);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_Address_Country", SqlDbType.VarChar, 150);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_Address_ZipCode", SqlDbType.VarChar, 10);
			cmInsertAF.Parameters.Add("@req_Order_ShippingData_Address_Reference", SqlDbType.VarChar, 250);
			cmInsertAF.Prepare();
			#endregion

			#region [ cmInsertAFItem ]
			strSql = "INSERT INTO t_PAGTO_GW_AF_ITEM (" +
						"id, " +
						"id_pagto_gw_af, " +
						"af_ID, " +
						"af_Name, " +
						"af_ItemValue, " +
						"af_Qty, " +
						"af_CategoryID, " +
						"af_CategoryName" +
					") VALUES (" +
						"@id, " +
						"@id_pagto_gw_af, " +
						"@af_ID, " +
						"@af_Name, " +
						"@af_ItemValue, " +
						"@af_Qty, " +
						"@af_CategoryID, " +
						"@af_CategoryName" +
					")";
			cmInsertAFItem = BD.criaSqlCommand();
			cmInsertAFItem.CommandText = strSql;
			cmInsertAFItem.Parameters.Add("@id", SqlDbType.Int);
			cmInsertAFItem.Parameters.Add("@id_pagto_gw_af", SqlDbType.Int);
			cmInsertAFItem.Parameters.Add("@af_ID", SqlDbType.VarChar, 50);
			cmInsertAFItem.Parameters.Add("@af_Name", SqlDbType.VarChar, 150);
			cmInsertAFItem.Parameters.Add("@af_ItemValue", SqlDbType.VarChar, 24);
			cmInsertAFItem.Parameters.Add("@af_Qty", SqlDbType.VarChar, 6);
			cmInsertAFItem.Parameters.Add("@af_CategoryID", SqlDbType.VarChar, 6);
			cmInsertAFItem.Parameters.Add("@af_CategoryName", SqlDbType.VarChar, 200);
			cmInsertAFItem.Prepare();
			#endregion

			#region [ cmInsertAFPayment ]
			strSql = "INSERT INTO t_PAGTO_GW_AF_PAYMENT (" +
						"id, " +
						"id_pagto_gw_af, " +
						"id_pagto_gw_pag_payment, " +
						"ordem, " +
						"bandeira, " +
						"valor_transacao, " +
						"af_Sequential, " +
						"af_Date, " +
						"af_Amount, " +
						"af_PaymentTypeID, " +
						"af_QtyInstallments, " +
						"af_Interest, " +
						"af_InterestValue, " +
						"af_CardNumber, " +
						"af_CardBin, " +
						"af_CardEndNumber, " +
						"af_CardType, " +
						"af_CardExpirationDate, " +
						"af_Name, " +
						"af_LegalDocument, " +
						"af_Address_Street, " +
						"af_Address_Number, " +
						"af_Address_Comp, " +
						"af_Address_County, " +
						"af_Address_City, " +
						"af_Address_State, " +
						"af_Address_Country, " +
						"af_Address_ZipCode, " +
						"af_Address_Reference, " +
						"af_Nsu, " +
						"af_Currency" +
					") VALUES (" +
						"@id, " +
						"@id_pagto_gw_af, " +
						"@id_pagto_gw_pag_payment, " +
						"@ordem, " +
						"@bandeira, " +
						"@valor_transacao, " +
						"@af_Sequential, " +
						"@af_Date, " +
						"@af_Amount, " +
						"@af_PaymentTypeID, " +
						"@af_QtyInstallments, " +
						"@af_Interest, " +
						"@af_InterestValue, " +
						"@af_CardNumber, " +
						"@af_CardBin, " +
						"@af_CardEndNumber, " +
						"@af_CardType, " +
						"@af_CardExpirationDate, " +
						"@af_Name, " +
						"@af_LegalDocument, " +
						"@af_Address_Street, " +
						"@af_Address_Number, " +
						"@af_Address_Comp, " +
						"@af_Address_County, " +
						"@af_Address_City, " +
						"@af_Address_State, " +
						"@af_Address_Country, " +
						"@af_Address_ZipCode, " +
						"@af_Address_Reference, " +
						"@af_Nsu, " +
						"@af_Currency" +
					")";
			cmInsertAFPayment = BD.criaSqlCommand();
			cmInsertAFPayment.CommandText = strSql;
			cmInsertAFPayment.Parameters.Add("@id", SqlDbType.Int);
			cmInsertAFPayment.Parameters.Add("@id_pagto_gw_af", SqlDbType.Int);
			cmInsertAFPayment.Parameters.Add("@id_pagto_gw_pag_payment", SqlDbType.Int);
			cmInsertAFPayment.Parameters.Add("@ordem", SqlDbType.SmallInt);
			cmInsertAFPayment.Parameters.Add("@bandeira", SqlDbType.VarChar, 20);
			cmInsertAFPayment.Parameters.Add("@valor_transacao", SqlDbType.Money);
			cmInsertAFPayment.Parameters.Add("@af_Sequential", SqlDbType.VarChar, 2);
			cmInsertAFPayment.Parameters.Add("@af_Date", SqlDbType.VarChar, 19);
			cmInsertAFPayment.Parameters.Add("@af_Amount", SqlDbType.VarChar, 24);
			cmInsertAFPayment.Parameters.Add("@af_PaymentTypeID", SqlDbType.VarChar, 2);
			cmInsertAFPayment.Parameters.Add("@af_QtyInstallments", SqlDbType.VarChar, 2);
			cmInsertAFPayment.Parameters.Add("@af_Interest", SqlDbType.VarChar, 6);
			cmInsertAFPayment.Parameters.Add("@af_InterestValue", SqlDbType.VarChar, 24);
			cmInsertAFPayment.Parameters.Add("@af_CardNumber", SqlDbType.VarChar, 200);
			cmInsertAFPayment.Parameters.Add("@af_CardBin", SqlDbType.VarChar, 6);
			cmInsertAFPayment.Parameters.Add("@af_CardEndNumber", SqlDbType.VarChar, 4);
			cmInsertAFPayment.Parameters.Add("@af_CardType", SqlDbType.VarChar, 2);
			cmInsertAFPayment.Parameters.Add("@af_CardExpirationDate", SqlDbType.VarChar, 50);
			cmInsertAFPayment.Parameters.Add("@af_Name", SqlDbType.VarChar, 150);
			cmInsertAFPayment.Parameters.Add("@af_LegalDocument", SqlDbType.VarChar, 100);
			cmInsertAFPayment.Parameters.Add("@af_Address_Street", SqlDbType.VarChar, 200);
			cmInsertAFPayment.Parameters.Add("@af_Address_Number", SqlDbType.VarChar, 15);
			cmInsertAFPayment.Parameters.Add("@af_Address_Comp", SqlDbType.VarChar, 250);
			cmInsertAFPayment.Parameters.Add("@af_Address_County", SqlDbType.VarChar, 150);
			cmInsertAFPayment.Parameters.Add("@af_Address_City", SqlDbType.VarChar, 150);
			cmInsertAFPayment.Parameters.Add("@af_Address_State", SqlDbType.VarChar, 2);
			cmInsertAFPayment.Parameters.Add("@af_Address_Country", SqlDbType.VarChar, 150);
			cmInsertAFPayment.Parameters.Add("@af_Address_ZipCode", SqlDbType.VarChar, 10);
			cmInsertAFPayment.Parameters.Add("@af_Address_Reference", SqlDbType.VarChar, 250);
			cmInsertAFPayment.Parameters.Add("@af_Nsu", SqlDbType.VarChar, 50);
			cmInsertAFPayment.Parameters.Add("@af_Currency", SqlDbType.VarChar, 4);
			cmInsertAFPayment.Prepare();
			#endregion

			#region [ cmInsertAFPhone ]
			strSql = "INSERT INTO t_PAGTO_GW_AF_PHONE (" +
						"id, " +
						"id_pagto_gw_af, " +
						"IdBlocoXml, " +
						"af_Type, " +
						"af_DDI, " +
						"af_DDD, " +
						"af_Number, " +
						"af_Extension" +
					") VALUES (" +
						"@id, " +
						"@id_pagto_gw_af, " +
						"@IdBlocoXml, " +
						"@af_Type, " +
						"@af_DDI, " +
						"@af_DDD, " +
						"@af_Number, " +
						"@af_Extension" +
					")";
			cmInsertAFPhone = BD.criaSqlCommand();
			cmInsertAFPhone.CommandText = strSql;
			cmInsertAFPhone.Parameters.Add("@id", SqlDbType.Int);
			cmInsertAFPhone.Parameters.Add("@id_pagto_gw_af", SqlDbType.Int);
			cmInsertAFPhone.Parameters.Add("@IdBlocoXml", SqlDbType.VarChar, 100);
			cmInsertAFPhone.Parameters.Add("@af_Type", SqlDbType.VarChar, 1);
			cmInsertAFPhone.Parameters.Add("@af_DDI", SqlDbType.VarChar, 3);
			cmInsertAFPhone.Parameters.Add("@af_DDD", SqlDbType.VarChar, 2);
			cmInsertAFPhone.Parameters.Add("@af_Number", SqlDbType.VarChar, 9);
			cmInsertAFPhone.Parameters.Add("@af_Extension", SqlDbType.VarChar, 10);
			cmInsertAFPhone.Prepare();
			#endregion

			#region [ cmInsertAFXml ]
			strSql = "INSERT INTO t_PAGTO_GW_AF_XML (" +
						"id, " +
						"id_pagto_gw_af, " +
						"tipo_transacao, " +
						"fluxo_xml, " +
						"xml" +
					") VALUES (" +
						"@id, " +
						"@id_pagto_gw_af, " +
						"@tipo_transacao, " +
						"@fluxo_xml, " +
						"@xml" +
					")";
			cmInsertAFXml = BD.criaSqlCommand();
			cmInsertAFXml.CommandText = strSql;
			cmInsertAFXml.Parameters.Add("@id", SqlDbType.Int);
			cmInsertAFXml.Parameters.Add("@id_pagto_gw_af", SqlDbType.Int);
			cmInsertAFXml.Parameters.Add("@tipo_transacao", SqlDbType.VarChar, 20);
			cmInsertAFXml.Parameters.Add("@fluxo_xml", SqlDbType.VarChar, 2);
			cmInsertAFXml.Parameters.Add("@xml", SqlDbType.VarChar, -1); // varchar(max)
			cmInsertAFXml.Prepare();
			#endregion

			#region [ cmInsertAFNsu ]
			strSql = "INSERT INTO t_PAGTO_GW_AF_NSU (" +
						"pedido, " +
						"nsu, " +
						"dt_hr_atualizacao, " +
						"usuario_atualizacao" +
					") VALUES (" +
						"@pedido, " +
						"@nsu, " +
						"getdate(), " +
						"@usuario_atualizacao" +
					")";
			cmInsertAFNsu = BD.criaSqlCommand();
			cmInsertAFNsu.CommandText = strSql;
			cmInsertAFNsu.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmInsertAFNsu.Parameters.Add("@nsu", SqlDbType.Int);
			cmInsertAFNsu.Parameters.Add("@usuario_atualizacao", SqlDbType.VarChar, 10);
			cmInsertAFNsu.Prepare();
			#endregion

			#region [ cmInsertAFOpComplementar ]
			strSql = "INSERT INTO t_PAGTO_GW_AF_OP_COMPLEMENTAR (" +
						"id, " +
						"id_pagto_gw_af, " +
						"usuario, " +
						"operacao, " +
						"trx_TX_data, " +
						"trx_TX_data_hora" +
					") VALUES (" +
						"@id, " +
						"@id_pagto_gw_af, " +
						"@usuario, " +
						"@operacao, " +
						Global.sqlMontaGetdateSomenteData() + ", " +
						"getdate()" +
					")";
			cmInsertAFOpComplementar = BD.criaSqlCommand();
			cmInsertAFOpComplementar.CommandText = strSql;
			cmInsertAFOpComplementar.Parameters.Add("@id", SqlDbType.Int);
			cmInsertAFOpComplementar.Parameters.Add("@id_pagto_gw_af", SqlDbType.Int);
			cmInsertAFOpComplementar.Parameters.Add("@usuario", SqlDbType.VarChar, 10);
			cmInsertAFOpComplementar.Parameters.Add("@operacao", SqlDbType.VarChar, 40);
			cmInsertAFOpComplementar.Prepare();
			#endregion

			#region [ cmInsertAFOpComplementarXml ]
			strSql = "INSERT INTO t_PAGTO_GW_AF_OP_COMPLEMENTAR_XML (" +
						"id, " +
						"id_pagto_gw_af_op_complementar, " +
						"tipo_transacao, " +
						"fluxo_xml, " +
						"xml" +
					") VALUES (" +
						"@id, " +
						"@id_pagto_gw_af_op_complementar, " +
						"@tipo_transacao, " +
						"@fluxo_xml, " +
						"@xml" +
					")";
			cmInsertAFOpComplementarXml = BD.criaSqlCommand();
			cmInsertAFOpComplementarXml.CommandText = strSql;
			cmInsertAFOpComplementarXml.Parameters.Add("@id", SqlDbType.Int);
			cmInsertAFOpComplementarXml.Parameters.Add("@id_pagto_gw_af_op_complementar", SqlDbType.Int);
			cmInsertAFOpComplementarXml.Parameters.Add("@tipo_transacao", SqlDbType.VarChar, 20);
			cmInsertAFOpComplementarXml.Parameters.Add("@fluxo_xml", SqlDbType.VarChar, 2);
			cmInsertAFOpComplementarXml.Parameters.Add("@xml", SqlDbType.VarChar, -1); // varchar(max)
			cmInsertAFOpComplementarXml.Prepare();
			#endregion

			#region [ cmUpdateAFAnulaRegistroTentativaAnterior ]
			strSql = "UPDATE t_PAGTO_GW_AF SET " +
						" anulado_status = 1, " +
						" anulado_data = " + Global.sqlMontaGetdateSomenteData() + ", " +
						" anulado_data_hora = getdate(), " +
						" anulado_por_id_pagto_gw_af = @anulado_por_id_pagto_gw_af" +
					" WHERE" +
						" (pedido = @pedido)" +
						" AND (anulado_status = 0)" +
						" AND (" +
							"(trx_RX_status = 0)" +
							" OR " +
							"(trx_RX_vazio_status = 1)" +
							" OR " +
							"(trx_erro_status = 1)" +
							")";
			cmUpdateAFAnulaRegistroTentativaAnterior = BD.criaSqlCommand();
			cmUpdateAFAnulaRegistroTentativaAnterior.CommandText = strSql;
			cmUpdateAFAnulaRegistroTentativaAnterior.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmUpdateAFAnulaRegistroTentativaAnterior.Parameters.Add("@anulado_por_id_pagto_gw_af", SqlDbType.Int);
			cmUpdateAFAnulaRegistroTentativaAnterior.Prepare();
			#endregion

			#region [ cmUpdateAFErro ]
			strSql = "UPDATE t_PAGTO_GW_AF SET " +
						" trx_RX_status = 1," +
						" trx_RX_data = " + Global.sqlMontaGetdateSomenteData() + ", " +
						" trx_RX_data_hora = getdate(), " +
						" trx_RX_vazio_status = @trx_RX_vazio_status, " +
						" trx_erro_status = @trx_erro_status, " +
						" trx_erro_codigo = @trx_erro_codigo, " +
						" trx_erro_mensagem = @trx_erro_mensagem" +
					" WHERE" +
						" (id = @id)";
			cmUpdateAFErroRx = BD.criaSqlCommand();
			cmUpdateAFErroRx.CommandText = strSql;
			cmUpdateAFErroRx.Parameters.Add("@id", SqlDbType.Int);
			cmUpdateAFErroRx.Parameters.Add("@trx_RX_vazio_status", SqlDbType.TinyInt);
			cmUpdateAFErroRx.Parameters.Add("@trx_erro_status", SqlDbType.TinyInt);
			cmUpdateAFErroRx.Parameters.Add("@trx_erro_codigo", SqlDbType.VarChar, 6);
			cmUpdateAFErroRx.Parameters.Add("@trx_erro_mensagem", SqlDbType.VarChar, -1);
			cmUpdateAFErroRx.Prepare();
			#endregion

			#region [ cmUpdateAFIdRegistroXmlTx ]
			strSql = "UPDATE t_PAGTO_GW_AF SET " +
						" trx_TX_id_pagto_gw_af_xml = @trx_TX_id_pagto_gw_af_xml, " +
						" trx_TX_data = " + Global.sqlMontaGetdateSomenteData() + ", " +
						" trx_TX_data_hora = getdate() " +
					" WHERE" +
						" (id = @id)";
			cmUpdateAFIdRegistroXmlTx = BD.criaSqlCommand();
			cmUpdateAFIdRegistroXmlTx.CommandText = strSql;
			cmUpdateAFIdRegistroXmlTx.Parameters.Add("@id", SqlDbType.Int);
			cmUpdateAFIdRegistroXmlTx.Parameters.Add("@trx_TX_id_pagto_gw_af_xml", SqlDbType.Int);
			cmUpdateAFIdRegistroXmlTx.Prepare();
			#endregion

			#region [ cmUpdateAFIdRegistroXmlRx ]
			strSql = "UPDATE t_PAGTO_GW_AF SET " +
						" trx_RX_id_pagto_gw_af_xml = @trx_RX_id_pagto_gw_af_xml, " +
						" trx_RX_status = 1, " +
						" trx_RX_vazio_status = @trx_RX_vazio_status, " +
						" trx_RX_data = " + Global.sqlMontaGetdateSomenteData() + ", " +
						" trx_RX_data_hora = getdate() " +
					" WHERE" +
						" (id = @id)";
			cmUpdateAFIdRegistroXmlRx = BD.criaSqlCommand();
			cmUpdateAFIdRegistroXmlRx.CommandText = strSql;
			cmUpdateAFIdRegistroXmlRx.Parameters.Add("@id", SqlDbType.Int);
			cmUpdateAFIdRegistroXmlRx.Parameters.Add("@trx_RX_id_pagto_gw_af_xml", SqlDbType.Int);
			cmUpdateAFIdRegistroXmlRx.Parameters.Add("@trx_RX_vazio_status", SqlDbType.TinyInt);
			cmUpdateAFIdRegistroXmlRx.Prepare();
			#endregion

			#region [ cmUpdatePagPaymentStEnviadoAnaliseAF ]
			strSql = "UPDATE t_PAGTO_GW_PAG_PAYMENT SET " +
						" st_enviado_analise_AF = 1, " +
						" id_pagto_gw_af = @id_pagto_gw_af" +
					" WHERE" +
						" (id = @id)";
			cmUpdatePagPaymentStEnviadoAnaliseAF = BD.criaSqlCommand();
			cmUpdatePagPaymentStEnviadoAnaliseAF.CommandText = strSql;
			cmUpdatePagPaymentStEnviadoAnaliseAF.Parameters.Add("@id", SqlDbType.Int);
			cmUpdatePagPaymentStEnviadoAnaliseAF.Parameters.Add("@id_pagto_gw_af", SqlDbType.Int);
			cmUpdatePagPaymentStEnviadoAnaliseAF.Prepare();
			#endregion

			#region [ cmUpdateAFNsu ]
			strSql = "UPDATE t_PAGTO_GW_AF_NSU SET " +
						" nsu = @nsu, " +
						" dt_hr_atualizacao = getdate(), " +
						" usuario_atualizacao = @usuario_atualizacao" +
					" WHERE" +
						" (pedido = @pedido)";
			cmUpdateAFNsu = BD.criaSqlCommand();
			cmUpdateAFNsu.CommandText = strSql;
			cmUpdateAFNsu.Parameters.Add("@pedido", SqlDbType.VarChar, 9);
			cmUpdateAFNsu.Parameters.Add("@nsu", SqlDbType.Int);
			cmUpdateAFNsu.Parameters.Add("@usuario_atualizacao", SqlDbType.VarChar, 10);
			cmUpdateAFNsu.Prepare();
			#endregion

			#region [ cmUpdateAFSendOrdersResponse ]
			strSql = "UPDATE t_PAGTO_GW_AF SET" +
						" resp_TransactionID = @resp_TransactionID," +
						" resp_StatusCode = @resp_StatusCode," +
						" resp_Message = @resp_Message," +
						" resp_ID = @resp_ID," +
						" resp_Status = @resp_Status," +
						" resp_Score = @resp_Score," +
						" prim_Status = @prim_Status," +
						" prim_atualizacao_data_hora = getdate()," +
						" ult_Status = @ult_Status," +
						" ult_atualizacao_data_hora = getdate()" +
					" WHERE" +
						" (id = @id)";
			cmUpdateAFSendOrdersResponse = BD.criaSqlCommand();
			cmUpdateAFSendOrdersResponse.CommandText = strSql;
			cmUpdateAFSendOrdersResponse.Parameters.Add("@id", SqlDbType.Int);
			cmUpdateAFSendOrdersResponse.Parameters.Add("@resp_TransactionID", SqlDbType.VarChar, 72);
			cmUpdateAFSendOrdersResponse.Parameters.Add("@resp_StatusCode", SqlDbType.VarChar, 10);
			cmUpdateAFSendOrdersResponse.Parameters.Add("@resp_Message", SqlDbType.VarChar, 2048);
			cmUpdateAFSendOrdersResponse.Parameters.Add("@resp_ID", SqlDbType.VarChar, 50);
			cmUpdateAFSendOrdersResponse.Parameters.Add("@resp_Status", SqlDbType.VarChar, 3);
			cmUpdateAFSendOrdersResponse.Parameters.Add("@resp_Score", SqlDbType.VarChar, 8);
			cmUpdateAFSendOrdersResponse.Parameters.Add("@prim_Status", SqlDbType.VarChar, 3);
			cmUpdateAFSendOrdersResponse.Parameters.Add("@ult_Status", SqlDbType.VarChar, 3);
			cmUpdateAFSendOrdersResponse.Prepare();
			#endregion

			#region [ cmUpdateAFGetReturnAnalysisResponse ]
			strSql = "UPDATE t_PAGTO_GW_AF SET" +
						" resp_ID = @resp_ID," +
						" resp_Status = @resp_Status," +
						" resp_Score = @resp_Score," +
						" ult_Status = @ult_Status," +
						" ult_atualizacao_data_hora = getdate()" +
					" WHERE" +
						" (id = @id)";
			cmUpdateAFGetReturnAnalysisResponse = BD.criaSqlCommand();
			cmUpdateAFGetReturnAnalysisResponse.CommandText = strSql;
			cmUpdateAFGetReturnAnalysisResponse.Parameters.Add("@id", SqlDbType.Int);
			cmUpdateAFGetReturnAnalysisResponse.Parameters.Add("@resp_ID", SqlDbType.VarChar, 50);
			cmUpdateAFGetReturnAnalysisResponse.Parameters.Add("@resp_Status", SqlDbType.VarChar, 3);
			cmUpdateAFGetReturnAnalysisResponse.Parameters.Add("@resp_Score", SqlDbType.VarChar, 8);
			cmUpdateAFGetReturnAnalysisResponse.Parameters.Add("@ult_Status", SqlDbType.VarChar, 3);
			cmUpdateAFGetReturnAnalysisResponse.Prepare();
			#endregion

			#region [ cmUpdateAFOpComplementar ]
			strSql = "UPDATE t_PAGTO_GW_AF_OP_COMPLEMENTAR SET" +
						" trx_RX_data = " + Global.sqlMontaGetdateSomenteData() + "," +
						" trx_RX_data_hora = getdate()," +
						" trx_RX_status = @trx_RX_status," +
						" trx_RX_vazio_status = @trx_RX_vazio_status," +
						" st_sucesso = @st_sucesso" +
					" WHERE" +
						" (id = @id)";
			cmUpdateAFOpComplementar = BD.criaSqlCommand();
			cmUpdateAFOpComplementar.CommandText = strSql;
			cmUpdateAFOpComplementar.Parameters.Add("@id", SqlDbType.Int);
			cmUpdateAFOpComplementar.Parameters.Add("@trx_RX_status", SqlDbType.TinyInt);
			cmUpdateAFOpComplementar.Parameters.Add("@trx_RX_vazio_status", SqlDbType.TinyInt);
			cmUpdateAFOpComplementar.Parameters.Add("@st_sucesso", SqlDbType.TinyInt);
			cmUpdateAFOpComplementar.Prepare();
			#endregion

			#region [ cmUpdateAFSetOrderAsReturnedPendente ]
			strSql = "UPDATE t_PAGTO_GW_AF SET" +
						" SetOrderAsReturned_pendente_status = 1," +
						" SetOrderAsReturned_pendente_data_hora = getdate()" +
					" WHERE" +
						" (id = @id)";
			cmUpdateAFSetOrderAsReturnedPendente = BD.criaSqlCommand();
			cmUpdateAFSetOrderAsReturnedPendente.CommandText = strSql;
			cmUpdateAFSetOrderAsReturnedPendente.Parameters.Add("@id", SqlDbType.Int);
			cmUpdateAFSetOrderAsReturnedPendente.Prepare();
			#endregion

			#region [ cmUpdateAFSetOrderAsReturnedSucesso ]
			strSql = "UPDATE t_PAGTO_GW_AF SET" +
						" SetOrderAsReturned_sucesso_status = 1," +
						" SetOrderAsReturned_sucesso_data_hora = getdate()" +
					" WHERE" +
						" (id = @id)";
			cmUpdateAFSetOrderAsReturnedSucesso = BD.criaSqlCommand();
			cmUpdateAFSetOrderAsReturnedSucesso.CommandText = strSql;
			cmUpdateAFSetOrderAsReturnedSucesso.Parameters.Add("@id", SqlDbType.Int);
			cmUpdateAFSetOrderAsReturnedSucesso.Prepare();
			#endregion
		}
		#endregion

		#region [ insereAF ]
		public static bool insereAF(ClearsaleAF clearsaleAF, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.insereAF()";
			bool blnGerouNsu;
			int intRetorno;
			int idPagtoGwAf = 0;
			int idPagtoGwAfItem = 0;
			int idPagtoGwAfPayment = 0;
			int idPagtoGwAfPhone = 0;
			string msg_erro_aux;
			ClearsaleAFPayment afPayment;
			ClearsaleAFItem afItem;
			ClearsaleAFPhone afPhone;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Gera NSU ]

				#region [ Gera o NSU para o registro principal? ]
				if (clearsaleAF.id == 0)
				{
					blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.FIN.NSU.T_PAGTO_GW_AF, out idPagtoGwAf, out msg_erro);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro principal dos dados da análise antifraude!!\n" + msg_erro;
						return false;
					}
					clearsaleAF.id = idPagtoGwAf;
				}
				else
				{
					// O NSU já foi gerado anteriormente na rotina chamadora
					// Isso pode ocorrer quando é necessário referenciar o ID da nova transação antes mesmo dela ter sido gravada (ex:
					// cancelamento de registros de tentativas anteriores)
					idPagtoGwAf = clearsaleAF.id;
				}
				#endregion

				#region [ Gera o NSU para os registros de 'Payment' ]
				for (int i = 0; i < clearsaleAF.Payments.Count; i++)
				{
					blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.FIN.NSU.T_PAGTO_GW_AF_PAYMENT, out idPagtoGwAfPayment, out msg_erro);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro com os dados de 'Payment' da análise antifraude!!\n" + msg_erro;
						return false;
					}
					clearsaleAF.Payments[i].id = idPagtoGwAfPayment;
					clearsaleAF.Payments[i].id_pagto_gw_af = idPagtoGwAf;
				}
				#endregion

				#region [ Gera o NSU para os registros de 'Item' ]
				for (int i = 0; i < clearsaleAF.Items.Count; i++)
				{
					blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.FIN.NSU.T_PAGTO_GW_AF_ITEM, out idPagtoGwAfItem, out msg_erro);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro com os dados de 'Item' da análise antifraude!!\n" + msg_erro;
						return false;
					}
					clearsaleAF.Items[i].id = idPagtoGwAfItem;
					clearsaleAF.Items[i].id_pagto_gw_af = idPagtoGwAf;
				}
				#endregion

				#region [ Gera o NSU para os registros de 'BillingData/Phones' ]
				for (int i = 0; i < clearsaleAF.Order_BillingData_Phones.Count; i++)
				{
					blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.FIN.NSU.T_PAGTO_GW_AF_PHONE, out idPagtoGwAfPhone, out msg_erro);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro com os dados de 'BillingData/Phones' da análise antifraude!!\n" + msg_erro;
						return false;
					}
					clearsaleAF.Order_BillingData_Phones[i].id = idPagtoGwAfPhone;
					clearsaleAF.Order_BillingData_Phones[i].id_pagto_gw_af = idPagtoGwAf;
				}
				#endregion

				#region [ Gera o NSU para os registros de 'ShippingData/Phones' ]
				for (int i = 0; i < clearsaleAF.Order_ShippingData_Phones.Count; i++)
				{
					blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.FIN.NSU.T_PAGTO_GW_AF_PHONE, out idPagtoGwAfPhone, out msg_erro);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro com os dados de 'ShippingData/Phones' da análise antifraude!!\n" + msg_erro;
						return false;
					}
					clearsaleAF.Order_ShippingData_Phones[i].id = idPagtoGwAfPhone;
					clearsaleAF.Order_ShippingData_Phones[i].id_pagto_gw_af = idPagtoGwAf;
				}
				#endregion

				#endregion

				#region [ Tenta gravar o registro principal ]

				#region [ Preenche o valor dos parâmetros ]
				cmInsertAF.Parameters["@id"].Value = clearsaleAF.id;
				cmInsertAF.Parameters["@usuario"].Value = clearsaleAF.usuario;
				cmInsertAF.Parameters["@owner"].Value = clearsaleAF.owner;
				cmInsertAF.Parameters["@loja"].Value = clearsaleAF.loja;
				cmInsertAF.Parameters["@id_cliente"].Value = clearsaleAF.id_cliente;
				cmInsertAF.Parameters["@pedido"].Value = clearsaleAF.pedido;
				cmInsertAF.Parameters["@pedido_com_sufixo_nsu"].Value = clearsaleAF.pedido_com_sufixo_nsu;
				cmInsertAF.Parameters["@valor_pedido"].Value = clearsaleAF.valor_pedido;
				cmInsertAF.Parameters["@req_entityCode"].Value = clearsaleAF.req_entityCode;
				cmInsertAF.Parameters["@req_Order_ID"].Value = clearsaleAF.req_Order_ID;
				cmInsertAF.Parameters["@req_Order_FingerPrint_SessionID"].Value = clearsaleAF.req_Order_FingerPrint_SessionID;
				cmInsertAF.Parameters["@req_Order_Date"].Value = clearsaleAF.req_Order_Date;
				cmInsertAF.Parameters["@req_Order_Email"].Value = clearsaleAF.req_Order_Email;
				cmInsertAF.Parameters["@req_Order_B2B_B2C"].Value = clearsaleAF.req_Order_B2B_B2C;
				cmInsertAF.Parameters["@req_Order_ShippingPrice"].Value = clearsaleAF.req_Order_ShippingPrice;
				cmInsertAF.Parameters["@req_Order_TotalItems"].Value = clearsaleAF.req_Order_TotalItems;
				cmInsertAF.Parameters["@req_Order_TotalOrder"].Value = clearsaleAF.req_Order_TotalOrder;
				cmInsertAF.Parameters["@req_Order_QtyInstallments"].Value = clearsaleAF.req_Order_QtyInstallments;
				cmInsertAF.Parameters["@req_Order_DeliveryTimeCD"].Value = clearsaleAF.req_Order_DeliveryTimeCD;
				cmInsertAF.Parameters["@req_Order_QtyItems"].Value = clearsaleAF.req_Order_QtyItems;
				cmInsertAF.Parameters["@req_Order_QtyPaymentTypes"].Value = clearsaleAF.req_Order_QtyPaymentTypes;
				cmInsertAF.Parameters["@req_Order_IP"].Value = clearsaleAF.req_Order_IP;
				cmInsertAF.Parameters["@req_Order_Status"].Value = clearsaleAF.req_Order_Status;
				cmInsertAF.Parameters["@req_Order_Reanalise"].Value = clearsaleAF.req_Order_Reanalise;
				cmInsertAF.Parameters["@req_Order_Origin"].Value = clearsaleAF.req_Order_Origin;
				cmInsertAF.Parameters["@req_Order_BillingData_ID"].Value = clearsaleAF.req_Order_BillingData_ID;
				cmInsertAF.Parameters["@req_Order_BillingData_Type"].Value = clearsaleAF.req_Order_BillingData_Type;
				cmInsertAF.Parameters["@req_Order_BillingData_LegalDocument1"].Value = clearsaleAF.req_Order_BillingData_LegalDocument1;
				cmInsertAF.Parameters["@req_Order_BillingData_LegalDocument2"].Value = clearsaleAF.req_Order_BillingData_LegalDocument2;
				cmInsertAF.Parameters["@req_Order_BillingData_Name"].Value = clearsaleAF.req_Order_BillingData_Name;
				cmInsertAF.Parameters["@req_Order_BillingData_BirthDate"].Value = clearsaleAF.req_Order_BillingData_BirthDate;
				cmInsertAF.Parameters["@req_Order_BillingData_Email"].Value = clearsaleAF.req_Order_BillingData_Email;
				cmInsertAF.Parameters["@req_Order_BillingData_Gender"].Value = clearsaleAF.req_Order_BillingData_Gender;
				cmInsertAF.Parameters["@req_Order_BillingData_Address_Street"].Value = clearsaleAF.req_Order_BillingData_Address_Street;
				cmInsertAF.Parameters["@req_Order_BillingData_Address_Number"].Value = clearsaleAF.req_Order_BillingData_Address_Number;
				cmInsertAF.Parameters["@req_Order_BillingData_Address_Comp"].Value = clearsaleAF.req_Order_BillingData_Address_Comp;
				cmInsertAF.Parameters["@req_Order_BillingData_Address_County"].Value = clearsaleAF.req_Order_BillingData_Address_County;
				cmInsertAF.Parameters["@req_Order_BillingData_Address_City"].Value = clearsaleAF.req_Order_BillingData_Address_City;
				cmInsertAF.Parameters["@req_Order_BillingData_Address_State"].Value = clearsaleAF.req_Order_BillingData_Address_State;
				cmInsertAF.Parameters["@req_Order_BillingData_Address_Country"].Value = clearsaleAF.req_Order_BillingData_Address_Country;
				cmInsertAF.Parameters["@req_Order_BillingData_Address_ZipCode"].Value = clearsaleAF.req_Order_BillingData_Address_ZipCode;
				cmInsertAF.Parameters["@req_Order_BillingData_Address_Reference"].Value = clearsaleAF.req_Order_BillingData_Address_Reference;
				cmInsertAF.Parameters["@req_Order_ShippingData_ID"].Value = clearsaleAF.req_Order_ShippingData_ID;
				cmInsertAF.Parameters["@req_Order_ShippingData_Type"].Value = clearsaleAF.req_Order_ShippingData_Type;
				cmInsertAF.Parameters["@req_Order_ShippingData_LegalDocument1"].Value = clearsaleAF.req_Order_ShippingData_LegalDocument1;
				cmInsertAF.Parameters["@req_Order_ShippingData_LegalDocument2"].Value = clearsaleAF.req_Order_ShippingData_LegalDocument2;
				cmInsertAF.Parameters["@req_Order_ShippingData_Name"].Value = clearsaleAF.req_Order_ShippingData_Name;
				cmInsertAF.Parameters["@req_Order_ShippingData_BirthDate"].Value = clearsaleAF.req_Order_ShippingData_BirthDate;
				cmInsertAF.Parameters["@req_Order_ShippingData_Email"].Value = clearsaleAF.req_Order_ShippingData_Email;
				cmInsertAF.Parameters["@req_Order_ShippingData_Gender"].Value = clearsaleAF.req_Order_ShippingData_Gender;
				cmInsertAF.Parameters["@req_Order_ShippingData_Address_Street"].Value = clearsaleAF.req_Order_ShippingData_Address_Street;
				cmInsertAF.Parameters["@req_Order_ShippingData_Address_Number"].Value = clearsaleAF.req_Order_ShippingData_Address_Number;
				cmInsertAF.Parameters["@req_Order_ShippingData_Address_Comp"].Value = clearsaleAF.req_Order_ShippingData_Address_Comp;
				cmInsertAF.Parameters["@req_Order_ShippingData_Address_County"].Value = clearsaleAF.req_Order_ShippingData_Address_County;
				cmInsertAF.Parameters["@req_Order_ShippingData_Address_City"].Value = clearsaleAF.req_Order_ShippingData_Address_City;
				cmInsertAF.Parameters["@req_Order_ShippingData_Address_State"].Value = clearsaleAF.req_Order_ShippingData_Address_State;
				cmInsertAF.Parameters["@req_Order_ShippingData_Address_Country"].Value = clearsaleAF.req_Order_ShippingData_Address_Country;
				cmInsertAF.Parameters["@req_Order_ShippingData_Address_ZipCode"].Value = clearsaleAF.req_Order_ShippingData_Address_ZipCode;
				cmInsertAF.Parameters["@req_Order_ShippingData_Address_Reference"].Value = clearsaleAF.req_Order_ShippingData_Address_Reference;
				#endregion

				#region [ Tenta inserir o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmInsertAF);
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
					svcLog.complemento_1 = Global.serializaObjectToXml(clearsaleAF);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion
				}
				#endregion

				#region [ Gravou o registro principal? ]
				if (intRetorno == 0)
				{
					msg_erro = "Falha ao tentar gravar o registro principal com os dados da análise antifraude!!\n" + msg_erro;
					return false;
				}
				#endregion

				#endregion

				#region [ Grava os dados de 'Payment' ]
				for (int i = 0; i < clearsaleAF.Payments.Count; i++)
				{
					afPayment = clearsaleAF.Payments[i];

					#region [ Preenche o valor dos parâmetros ]
					cmInsertAFPayment.Parameters["@id"].Value = afPayment.id;
					cmInsertAFPayment.Parameters["@id_pagto_gw_af"].Value = afPayment.id_pagto_gw_af;
					cmInsertAFPayment.Parameters["@id_pagto_gw_pag_payment"].Value = afPayment.id_pagto_gw_pag_payment;
					cmInsertAFPayment.Parameters["@ordem"].Value = afPayment.ordem;
					cmInsertAFPayment.Parameters["@bandeira"].Value = afPayment.bandeira;
					cmInsertAFPayment.Parameters["@valor_transacao"].Value = afPayment.valor_transacao;
					cmInsertAFPayment.Parameters["@af_Sequential"].Value = afPayment.af_Sequential;
					cmInsertAFPayment.Parameters["@af_Date"].Value = afPayment.af_Date;
					cmInsertAFPayment.Parameters["@af_Amount"].Value = afPayment.af_Amount;
					cmInsertAFPayment.Parameters["@af_PaymentTypeID"].Value = afPayment.af_PaymentTypeID;
					cmInsertAFPayment.Parameters["@af_QtyInstallments"].Value = afPayment.af_QtyInstallments;
					cmInsertAFPayment.Parameters["@af_Interest"].Value = afPayment.af_Interest;
					cmInsertAFPayment.Parameters["@af_InterestValue"].Value = afPayment.af_InterestValue;
					cmInsertAFPayment.Parameters["@af_CardNumber"].Value = afPayment.af_CardNumber;
					cmInsertAFPayment.Parameters["@af_CardBin"].Value = afPayment.af_CardBin;
					cmInsertAFPayment.Parameters["@af_CardEndNumber"].Value = afPayment.af_CardEndNumber;
					cmInsertAFPayment.Parameters["@af_CardType"].Value = afPayment.af_CardType;
					cmInsertAFPayment.Parameters["@af_CardExpirationDate"].Value = afPayment.af_CardExpirationDate;
					cmInsertAFPayment.Parameters["@af_Name"].Value = afPayment.af_Name;
					cmInsertAFPayment.Parameters["@af_LegalDocument"].Value = afPayment.af_LegalDocument;
					cmInsertAFPayment.Parameters["@af_Address_Street"].Value = afPayment.af_Address_Street;
					cmInsertAFPayment.Parameters["@af_Address_Number"].Value = afPayment.af_Address_Number;
					cmInsertAFPayment.Parameters["@af_Address_Comp"].Value = afPayment.af_Address_Comp;
					cmInsertAFPayment.Parameters["@af_Address_County"].Value = afPayment.af_Address_County;
					cmInsertAFPayment.Parameters["@af_Address_City"].Value = afPayment.af_Address_City;
					cmInsertAFPayment.Parameters["@af_Address_State"].Value = afPayment.af_Address_State;
					cmInsertAFPayment.Parameters["@af_Address_Country"].Value = afPayment.af_Address_Country;
					cmInsertAFPayment.Parameters["@af_Address_ZipCode"].Value = afPayment.af_Address_ZipCode;
					cmInsertAFPayment.Parameters["@af_Address_Reference"].Value = afPayment.af_Address_Reference;
					cmInsertAFPayment.Parameters["@af_Nsu"].Value = afPayment.af_Nsu;
					cmInsertAFPayment.Parameters["@af_Currency"].Value = afPayment.af_Currency;
					#endregion

					#region [ Tenta inserir o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmInsertAFPayment);
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
						svcLog.complemento_1 = Global.serializaObjectToXml(afPayment);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion
					}
					#endregion

					#region [ Gravou o registro? ]
					if (intRetorno == 0)
					{
						msg_erro = "Falha ao tentar gravar o registro com os dados de 'Payment' da análise antifraude!!\n" + msg_erro;
						return false;
					}
					#endregion
				}
				#endregion

				#region [ Grava os dados de 'Item' ]
				for (int i = 0; i < clearsaleAF.Items.Count; i++)
				{
					afItem = clearsaleAF.Items[i];

					#region [ Preenche o valor dos parâmetros ]
					cmInsertAFItem.Parameters["@id"].Value = afItem.id;
					cmInsertAFItem.Parameters["@id_pagto_gw_af"].Value = afItem.id_pagto_gw_af;
					cmInsertAFItem.Parameters["@af_ID"].Value = afItem.af_ID;
					cmInsertAFItem.Parameters["@af_Name"].Value = afItem.af_Name;
					cmInsertAFItem.Parameters["@af_ItemValue"].Value = afItem.af_ItemValue;
					cmInsertAFItem.Parameters["@af_Qty"].Value = afItem.af_Qty;
					cmInsertAFItem.Parameters["@af_CategoryID"].Value = afItem.af_CategoryID;
					cmInsertAFItem.Parameters["@af_CategoryName"].Value = afItem.af_CategoryName;
					#endregion

					#region [ Tenta inserir o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmInsertAFItem);
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
						svcLog.complemento_1 = Global.serializaObjectToXml(afItem);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion
					}
					#endregion

					#region [ Gravou o registro? ]
					if (intRetorno == 0)
					{
						msg_erro = "Falha ao tentar gravar o registro com os dados de 'Item' da análise antifraude!!\n" + msg_erro;
						return false;
					}
					#endregion
				}
				#endregion

				#region [ Grava os dados de BillingData/Phones ]
				for (int i = 0; i < clearsaleAF.Order_BillingData_Phones.Count; i++)
				{
					afPhone = clearsaleAF.Order_BillingData_Phones[i];

					#region [ Preenche o valor dos parâmetros ]
					cmInsertAFPhone.Parameters["@id"].Value = afPhone.id;
					cmInsertAFPhone.Parameters["@id_pagto_gw_af"].Value = afPhone.id_pagto_gw_af;
					cmInsertAFPhone.Parameters["@IdBlocoXml"].Value = afPhone.idBlocoXml;
					cmInsertAFPhone.Parameters["@af_Type"].Value = afPhone.af_Type;
					cmInsertAFPhone.Parameters["@af_DDI"].Value = afPhone.af_DDI;
					cmInsertAFPhone.Parameters["@af_DDD"].Value = afPhone.af_DDD;
					cmInsertAFPhone.Parameters["@af_Number"].Value = afPhone.af_Number;
					cmInsertAFPhone.Parameters["@af_Extension"].Value = afPhone.af_Extension;
					#endregion

					#region [ Tenta inserir o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmInsertAFPhone);
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
						svcLog.complemento_1 = Global.serializaObjectToXml(afPhone);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion
					}
					#endregion

					#region [ Gravou o registro? ]
					if (intRetorno == 0)
					{
						msg_erro = "Falha ao tentar gravar o registro com os dados de 'BillingData/Phones' da análise antifraude!!\n" + msg_erro;
						return false;
					}
					#endregion
				}
				#endregion

				#region [ Grava os dados de ShippingData/Phones ]
				for (int i = 0; i < clearsaleAF.Order_ShippingData_Phones.Count; i++)
				{
					afPhone = clearsaleAF.Order_ShippingData_Phones[i];

					#region [ Preenche o valor dos parâmetros ]
					cmInsertAFPhone.Parameters["@id"].Value = afPhone.id;
					cmInsertAFPhone.Parameters["@id_pagto_gw_af"].Value = afPhone.id_pagto_gw_af;
					cmInsertAFPhone.Parameters["@IdBlocoXml"].Value = afPhone.idBlocoXml;
					cmInsertAFPhone.Parameters["@af_Type"].Value = afPhone.af_Type;
					cmInsertAFPhone.Parameters["@af_DDI"].Value = afPhone.af_DDI;
					cmInsertAFPhone.Parameters["@af_DDD"].Value = afPhone.af_DDD;
					cmInsertAFPhone.Parameters["@af_Number"].Value = afPhone.af_Number;
					cmInsertAFPhone.Parameters["@af_Extension"].Value = afPhone.af_Extension;
					#endregion

					#region [ Tenta inserir o registro ]
					try
					{
						intRetorno = BD.executaNonQuery(ref cmInsertAFPhone);
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
						svcLog.complemento_1 = Global.serializaObjectToXml(afPhone);
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion
					}
					#endregion

					#region [ Gravou o registro? ]
					if (intRetorno == 0)
					{
						msg_erro = "Falha ao tentar gravar o registro com os dados de 'ShippingData/Phones' da análise antifraude!!\n" + msg_erro;
						return false;
					}
					#endregion
				}
				#endregion

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Sucesso na gravação dos dados (t_PAGTO_GW_AF.id=" + clearsaleAF.id.ToString() + ")");

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
				svcLog.complemento_1 = Global.serializaObjectToXml(clearsaleAF);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ insereAFXml ]
		public static bool insereAFXml(ClearsaleAFXml clearsaleAFXml, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.insereAFXml()";
			bool blnGerouNsu;
			int idPagtoGwAfXml = 0;
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";

			try
			{
				#region [ Gera NSU ]
				if (clearsaleAFXml.id == 0)
				{
					blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.FIN.NSU.T_PAGTO_GW_AF_XML, out idPagtoGwAfXml, out msg_erro);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro que armazena o XML da transação!!\n" + msg_erro;
						return false;
					}
					clearsaleAFXml.id = idPagtoGwAfXml;
				}
				else
				{
					// O NSU já foi gerado anteriormente na rotina chamadora
					idPagtoGwAfXml = clearsaleAFXml.id;
				}
				#endregion

				#region [ Preenche o valor dos parâmetros ]
				cmInsertAFXml.Parameters["@id"].Value = clearsaleAFXml.id;
				cmInsertAFXml.Parameters["@id_pagto_gw_af"].Value = clearsaleAFXml.id_pagto_gw_af;
				cmInsertAFXml.Parameters["@tipo_transacao"].Value = clearsaleAFXml.tipo_transacao;
				cmInsertAFXml.Parameters["@fluxo_xml"].Value = clearsaleAFXml.fluxo_xml;
				cmInsertAFXml.Parameters["@xml"].Value = clearsaleAFXml.xml;
				#endregion

				#region [ Tenta inserir o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmInsertAFXml);
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
					svcLog.complemento_1 = Global.serializaObjectToXml(clearsaleAFXml);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion
				}
				#endregion

				#region [ Gravou o registro? ]
				if (intRetorno == 0)
				{
					msg_erro = "Falha ao tentar gravar o registro que armazena o XML da transação!!\n" + msg_erro;
					return false;
				}
				#endregion

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Sucesso na gravação dos dados (t_PAGTO_GW_AF_XML.id=" + clearsaleAFXml.id.ToString() + ")");

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
				svcLog.complemento_1 = Global.serializaObjectToXml(clearsaleAFXml);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ insereAFOpComplementar ]
		public static bool insereAFOpComplementar(ClearsaleAFOpComplementar op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.insereAFOpComplementar()";
			bool blnGerouNsu;
			int idPagtoGwAfOpCompl = 0;
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				if (op.id == 0)
				{
					blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.FIN.NSU.T_PAGTO_GW_AF_OP_COMPLEMENTAR, out idPagtoGwAfOpCompl, out msg_erro);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_AF_OP_COMPLEMENTAR + "\n" + msg_erro;
						return false;
					}
					op.id = idPagtoGwAfOpCompl;
				}
				else
				{
					// O NSU já foi gerado anteriormente na rotina chamadora
					idPagtoGwAfOpCompl = op.id;
				}

				#region [ Preenche o valor dos parâmetros ]
				cmInsertAFOpComplementar.Parameters["@id"].Value = op.id;
				cmInsertAFOpComplementar.Parameters["@id_pagto_gw_af"].Value = op.id_pagto_gw_af;
				cmInsertAFOpComplementar.Parameters["@usuario"].Value = op.usuario;
				cmInsertAFOpComplementar.Parameters["@operacao"].Value = op.operacao;
				#endregion

				#region [ Tenta inserir o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmInsertAFOpComplementar);
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
					msg_erro = "Falha ao tentar gravar o registro na tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_AF_OP_COMPLEMENTAR + "\n" + msg_erro;
					return false;
				}
				#endregion

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Sucesso na gravação dos dados (t_PAGTO_GW_AF_OP_COMPLEMENTAR.id=" + op.id.ToString() + ")");

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

		#region [ insereAFOpComplementarXml ]
		public static bool insereAFOpComplementarXml(ClearsaleAFOpComplementarXml op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.insereAFOpComplementarXml()";
			bool blnGerouNsu;
			int idPagtoGwAfOpComplXml = 0;
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				if (op.id == 0)
				{
					blnGerouNsu = BD.geraNsuUsandoTabelaFinControle(Global.Cte.FIN.NSU.T_PAGTO_GW_AF_OP_COMPLEMENTAR_XML, out idPagtoGwAfOpComplXml, out msg_erro);
					if (!blnGerouNsu)
					{
						msg_erro = "Falha ao tentar gerar o NSU para o registro da tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_AF_OP_COMPLEMENTAR_XML + "\n" + msg_erro;
						return false;
					}
					op.id = idPagtoGwAfOpComplXml;
				}
				else
				{
					// O NSU já foi gerado anteriormente na rotina chamadora
					idPagtoGwAfOpComplXml = op.id;
				}

				#region [ Preenche o valor dos parâmetros ]
				cmInsertAFOpComplementarXml.Parameters["@id"].Value = op.id;
				cmInsertAFOpComplementarXml.Parameters["@id_pagto_gw_af_op_complementar"].Value = op.id_pagto_gw_af_op_complementar;
				cmInsertAFOpComplementarXml.Parameters["@tipo_transacao"].Value = op.tipo_transacao;
				cmInsertAFOpComplementarXml.Parameters["@fluxo_xml"].Value = op.fluxo_xml;
				cmInsertAFOpComplementarXml.Parameters["@xml"].Value = op.xml;
				#endregion

				#region [ Tenta inserir o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmInsertAFOpComplementarXml);
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
					msg_erro = "Falha ao tentar gravar o registro na tabela " + Global.Cte.FIN.NSU.T_PAGTO_GW_AF_OP_COMPLEMENTAR_XML + "\n" + msg_erro;
					return false;
				}
				#endregion

				Global.gravaLogAtividade(NOME_DESTA_ROTINA + ": Sucesso na gravação dos dados (t_PAGTO_GW_AF_OP_COMPLEMENTAR_XML.id=" + op.id.ToString() + ")");

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

		#region [ updateAFOpComplementar ]
		public static bool updateAFOpComplementar(ClearsaleAFOpComplementar op, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.updateAFOpComplementar()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateAFOpComplementar.Parameters["@id"].Value = op.id;
				cmUpdateAFOpComplementar.Parameters["@trx_RX_status"].Value = op.trx_RX_status;
				cmUpdateAFOpComplementar.Parameters["@trx_RX_vazio_status"].Value = op.trx_RX_vazio_status;
				cmUpdateAFOpComplementar.Parameters["@st_sucesso"].Value = op.st_sucesso;
				#endregion

				#region [ Tenta alterar o(s) registro(s), se houver algum ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateAFOpComplementar);
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

		#region [ anulaRegistroAFTentativaAnterior ]
		public static bool anulaRegistroAFTentativaAnterior(string pedido, int anulado_por_id_pagto_gw_af, out int qtde_anulados, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.anulaRegistroAFTentativaAnterior()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			qtde_anulados = 0;
			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateAFAnulaRegistroTentativaAnterior.Parameters["@pedido"].Value = pedido;
				cmUpdateAFAnulaRegistroTentativaAnterior.Parameters["@anulado_por_id_pagto_gw_af"].Value = anulado_por_id_pagto_gw_af;
				#endregion

				#region [ Tenta alterar o(s) registro(s), se houver algum ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateAFAnulaRegistroTentativaAnterior);
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
					svcLog.complemento_1 = "Pedido=" + pedido + ", anulado_por_id_pagto_gw_af=" + anulado_por_id_pagto_gw_af.ToString();
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}
				#endregion

				qtde_anulados = intRetorno;
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
				svcLog.complemento_1 = "Pedido=" + pedido + ", anulado_por_id_pagto_gw_af=" + anulado_por_id_pagto_gw_af.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updateRegistroAFRespostaErro ]
		public static bool updateRegistroAFRespostaErro(int id, byte trx_RX_vazio_status, byte trx_erro_status, string trx_erro_codigo, string trx_erro_mensagem, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.updateRegistroAFRespostaErro()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateAFErroRx.Parameters["@id"].Value = id;
				cmUpdateAFErroRx.Parameters["@trx_RX_vazio_status"].Value = trx_RX_vazio_status;
				cmUpdateAFErroRx.Parameters["@trx_erro_status"].Value = trx_erro_status;
				cmUpdateAFErroRx.Parameters["@trx_erro_codigo"].Value = (trx_erro_codigo == null) ? "" : trx_erro_codigo;
				cmUpdateAFErroRx.Parameters["@trx_erro_mensagem"].Value = (trx_erro_mensagem == null) ? "" : trx_erro_mensagem;
				#endregion

				#region [ Tenta alterar o(s) registro(s), se houver algum ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateAFErroRx);
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
					svcLog.complemento_1 = "id=" + id.ToString() + ", trx_RX_vazio_status=" + trx_RX_vazio_status.ToString() + ", trx_erro_status=" + trx_erro_status.ToString() + ", trx_erro_codigo=" + trx_erro_codigo.ToString() + ", trx_erro_mensagem=" + trx_erro_mensagem;
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
				svcLog.complemento_1 = "id=" + id.ToString() + ", trx_RX_vazio_status=" + trx_RX_vazio_status.ToString() + ", trx_erro_status=" + trx_erro_status.ToString() + ", trx_erro_codigo=" + trx_erro_codigo.ToString() + ", trx_erro_mensagem=" + trx_erro_mensagem;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updateRegistroAFTxXml ]
		public static bool updateRegistroAFTxXml(int id, int trx_TX_id_pagto_gw_af_xml, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.updateRegistroAFTxXml()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateAFIdRegistroXmlTx.Parameters["@id"].Value = id;
				cmUpdateAFIdRegistroXmlTx.Parameters["@trx_TX_id_pagto_gw_af_xml"].Value = trx_TX_id_pagto_gw_af_xml;
				#endregion

				#region [ Tenta alterar o(s) registro(s), se houver algum ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateAFIdRegistroXmlTx);
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
					svcLog.complemento_1 = "id=" + id.ToString() + ", trx_TX_id_pagto_gw_af_xml=" + trx_TX_id_pagto_gw_af_xml.ToString();
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
				svcLog.complemento_1 = "id=" + id.ToString() + ", trx_TX_id_pagto_gw_af_xml=" + trx_TX_id_pagto_gw_af_xml.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updateRegistroAFRxXml ]
		public static bool updateRegistroAFRxXml(int id, int trx_RX_id_pagto_gw_af_xml, byte trx_RX_vazio_status, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.updateRegistroAFRxXml()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateAFIdRegistroXmlRx.Parameters["@id"].Value = id;
				cmUpdateAFIdRegistroXmlRx.Parameters["@trx_RX_id_pagto_gw_af_xml"].Value = trx_RX_id_pagto_gw_af_xml;
				cmUpdateAFIdRegistroXmlRx.Parameters["@trx_RX_vazio_status"].Value = trx_RX_vazio_status;
				#endregion

				#region [ Tenta alterar o(s) registro(s), se houver algum ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateAFIdRegistroXmlRx);
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
					svcLog.complemento_1 = "id=" + id.ToString() + ", trx_RX_id_pagto_gw_af_xml=" + trx_RX_id_pagto_gw_af_xml.ToString() + ", trx_RX_vazio_status=" + trx_RX_vazio_status.ToString();
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
				svcLog.complemento_1 = "id=" + id.ToString() + ", trx_RX_id_pagto_gw_af_xml=" + trx_RX_id_pagto_gw_af_xml.ToString() + ", trx_RX_vazio_status=" + trx_RX_vazio_status.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updatePagPaymentStEnviadoAnaliseAF ]
		public static bool updatePagPaymentStEnviadoAnaliseAF(int id, int id_pagto_gw_af, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.updatePagPaymentStEnviadoAnaliseAF()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdatePagPaymentStEnviadoAnaliseAF.Parameters["@id"].Value = id;
				cmUpdatePagPaymentStEnviadoAnaliseAF.Parameters["@id_pagto_gw_af"].Value = id_pagto_gw_af;
				#endregion

				#region [ Tenta alterar o(s) registro(s), se houver algum ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdatePagPaymentStEnviadoAnaliseAF);
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
					svcLog.complemento_1 = "id=" + id.ToString() + ", id_pagto_gw_af=" + id_pagto_gw_af.ToString();
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
				svcLog.complemento_1 = "id=" + id.ToString() + ", id_pagto_gw_af=" + id_pagto_gw_af.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updateRegistroAFSendOrdersResponse ]
		public static bool updateRegistroAFSendOrdersResponse(string resp_TransactionID, string resp_StatusCode, string resp_Message, ClearsaleSendOrdersResponseOrder trx, string entityCode, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.updateRegistroAFSendOrdersResponse()";
			int intRetorno;
			string msg_erro_aux;
			ClearsaleAF clearsaleAF;
			#endregion

			msg_erro = "";
			try
			{
				clearsaleAF = getClearsaleAFByOrderID(trx.ID, entityCode, out msg_erro_aux);
				if (clearsaleAF == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = NOME_DESTA_ROTINA + " - Falha ao tentar obter os dados do registro de AF do pedido " + trx.ID;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(trx);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				#region [ Preenche o valor dos parâmetros ]
				cmUpdateAFSendOrdersResponse.Parameters["@id"].Value = clearsaleAF.id;
				cmUpdateAFSendOrdersResponse.Parameters["@resp_TransactionID"].Value = resp_TransactionID;
				cmUpdateAFSendOrdersResponse.Parameters["@resp_StatusCode"].Value = resp_StatusCode;
				cmUpdateAFSendOrdersResponse.Parameters["@resp_Message"].Value = Texto.leftStr(resp_Message, 2048);
				cmUpdateAFSendOrdersResponse.Parameters["@resp_ID"].Value = trx.ID;
				cmUpdateAFSendOrdersResponse.Parameters["@resp_Status"].Value = trx.Status;
				cmUpdateAFSendOrdersResponse.Parameters["@resp_Score"].Value = trx.Score;
				cmUpdateAFSendOrdersResponse.Parameters["@prim_Status"].Value = trx.Status;
				cmUpdateAFSendOrdersResponse.Parameters["@ult_Status"].Value = trx.Status;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateAFSendOrdersResponse);
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
					svcLog.complemento_1 = Global.serializaObjectToXml(trx);
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
				svcLog.complemento_1 = Global.serializaObjectToXml(trx);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updateRegistroAFGetReturnAnalysisResponse ]
		public static bool updateRegistroAFGetReturnAnalysisResponse(ClearsaleGetReturnAnalysisResponse trx, string entityCode, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.updateRegistroAFGetReturnAnalysisResponse()";
			int intRetorno;
			string msg_erro_aux;
			ClearsaleAF clearsaleAF;
			#endregion

			msg_erro = "";
			try
			{
				clearsaleAF = getClearsaleAFByOrderID(trx.ID, entityCode, out msg_erro_aux);
				if (clearsaleAF == null)
				{
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = NOME_DESTA_ROTINA + " - Falha ao tentar obter os dados do registro de AF do pedido " + trx.ID;

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = msg_erro;
					svcLog.complemento_1 = Global.serializaObjectToXml(trx);
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion

					return false;
				}

				#region [ Preenche o valor dos parâmetros ]
				cmUpdateAFGetReturnAnalysisResponse.Parameters["@id"].Value = clearsaleAF.id;
				cmUpdateAFGetReturnAnalysisResponse.Parameters["@resp_ID"].Value = trx.ID;
				cmUpdateAFGetReturnAnalysisResponse.Parameters["@resp_Status"].Value = trx.Status;
				cmUpdateAFGetReturnAnalysisResponse.Parameters["@resp_Score"].Value = trx.Score;
				cmUpdateAFGetReturnAnalysisResponse.Parameters["@ult_Status"].Value = trx.Status;
				#endregion

				#region [ Tenta alterar o registro ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateAFGetReturnAnalysisResponse);
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
					svcLog.complemento_1 = Global.serializaObjectToXml(trx);
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
				svcLog.complemento_1 = Global.serializaObjectToXml(trx);
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ contagemTentativasFalhaTX ]
		public static int contagemTentativasFalhaTX(string numeroPedido)
		{
			#region [ Declarações ]
			int intContagem;
			string strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			#endregion

			#region [ Cria objetos de BD ]
			cmCommand = BD.criaSqlCommand();
			daAdapter = BD.criaSqlDataAdapter();
			daAdapter.SelectCommand = cmCommand;
			daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			#endregion

			strSql = "SELECT" +
						" Count(*) AS qtde" +
					" FROM t_PAGTO_GW_AF" +
					" WHERE" +
						" (pedido = '" + numeroPedido + "')" +
						" AND (" +
							"(trx_RX_status = 0)" +
							" OR " +
							"(trx_RX_vazio_status = 1)" +
							" OR " +
							"(trx_erro_status = 1)" +
							")";
			cmCommand.CommandText = strSql;
			intContagem = (int)cmCommand.ExecuteScalar();
			return intContagem;
		}
		#endregion

		#region [ obtemDataHoraUltTentativaFalhaTX ]
		public static DateTime obtemDataHoraUltTentativaFalhaTX(string numeroPedido)
		{
			#region [ Declarações ]
			DateTime dtUltTentativaFalhaTX = DateTime.MinValue;
			string strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			Object objResultado;
			#endregion

			#region [ Cria objetos de BD ]
			cmCommand = BD.criaSqlCommand();
			daAdapter = BD.criaSqlDataAdapter();
			daAdapter.SelectCommand = cmCommand;
			daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
			#endregion

			strSql = "SELECT" +
						" Max(data_hora) AS data_hora" +
					" FROM t_PAGTO_GW_AF" +
					" WHERE" +
						" (pedido = '" + numeroPedido + "')" +
						" AND (" +
							"(trx_RX_status = 0)" +
							" OR " +
							"(trx_RX_vazio_status = 1)" +
							" OR " +
							"(trx_erro_status = 1)" +
							")";
			cmCommand.CommandText = strSql;
			objResultado = cmCommand.ExecuteScalar();
			if (!Convert.IsDBNull(objResultado))
			{
				dtUltTentativaFalhaTX = (DateTime)objResultado;
			}

			return dtUltTentativaFalhaTX;
		}
		#endregion

		#region [ geraSufixoPedidoNsuAf ]
		public static bool geraSufixoPedidoNsuAf(string numeroPedido, out int nsuSufixo, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.geraSufixoPedidoNsuAf()";
			int intNsu = 0;
			int intRetorno;
			String strSql;
			string msg_erro_aux;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			nsuSufixo = 0;
			msg_erro = "";

			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				strSql = "SELECT * FROM t_PAGTO_GW_AF_NSU WHERE (pedido = '" + numeroPedido + "')";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);

				if (dtbResultado.Rows.Count == 0)
				{
					#region [ Preenche o valor dos parâmetros ]
					cmInsertAFNsu.Parameters["@pedido"].Value = numeroPedido;
					cmInsertAFNsu.Parameters["@nsu"].Value = 0;
					cmInsertAFNsu.Parameters["@usuario_atualizacao"].Value = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;
					#endregion

					try
					{
						intRetorno = BD.executaNonQuery(ref cmInsertAFNsu);
					}
					catch (Exception ex)
					{
						intRetorno = 0;
						// Retorna mensagem de erro p/ rotina chamadora
						msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

						#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
						FinSvcLog svcLog = new FinSvcLog();
						svcLog.operacao = NOME_DESTA_ROTINA;
						svcLog.descricao = ex.ToString();
						svcLog.complemento_1 = "numeroPedido=" + numeroPedido;
						GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
						#endregion
					}

					if (intRetorno == 0)
					{
						if (msg_erro.Length > 0) msg_erro = "\n" + msg_erro;
						msg_erro = "Falha ao tentar criar o registro de controle do NSU do pedido " + numeroPedido + "!" + msg_erro;
						return false;
					}
				}
				else
				{
					intNsu = BD.readToInt(dtbResultado.Rows[0]["nsu"]);
				}

				intNsu++;

				cmUpdateAFNsu.Parameters["@pedido"].Value = numeroPedido;
				cmUpdateAFNsu.Parameters["@nsu"].Value = intNsu;
				cmUpdateAFNsu.Parameters["@usuario_atualizacao"].Value = Global.Cte.LogBd.Usuario.ID_USUARIO_SISTEMA;

				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateAFNsu);
				}
				catch (Exception ex)
				{
					intRetorno = 0;
					// Retorna mensagem de erro p/ rotina chamadora
					msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

					#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
					FinSvcLog svcLog = new FinSvcLog();
					svcLog.operacao = NOME_DESTA_ROTINA;
					svcLog.descricao = ex.ToString();
					svcLog.complemento_1 = "numeroPedido=" + numeroPedido;
					GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
					#endregion
				}

				if (intRetorno == 0)
				{
					if (msg_erro.Length > 0) msg_erro = "\n" + msg_erro;
					msg_erro = "Falha ao tentar incrementar o NSU do sufixo do pedido " + numeroPedido + "!" + msg_erro;
					return false;
				}

				nsuSufixo = intNsu;
				return true;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.Message;

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = "numeroPedido=" + numeroPedido;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ getPagPaymentRowsParaEnvioAF ]
		public static DataTable getPagPaymentRowsParaEnvioAF(string numeroPedido, out string strMsgErro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.getPagPaymentRowsParaEnvioAF()";
			string msg_erro_aux;
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			#endregion

			strMsgErro = "";

			try
			{
				#region [ Cria objetos de BD ]
				cmCommand = BD.criaSqlCommand();
				daAdapter = BD.criaSqlDataAdapter();
				daAdapter.SelectCommand = cmCommand;
				daAdapter.MissingSchemaAction = MissingSchemaAction.Add;
				#endregion

				strSql = "SELECT" +
							" t_PAG.data," +
							" t_PAG.data_hora," +
							" t_PAG.usuario," +
							" t_PAG.loja," +
							" t_PAG.id_cliente," +
							" t_PAG.pedido," +
							" t_PAG.pedido_com_sufixo_nsu," +
							" t_PAG.owner," +
							" t_PAG.executado_pelo_cliente_status," +
							" t_PAG.origem_endereco_IP," +
							" t_PAG.FingerPrint_SessionID," +
							" t_PAG.valor_pedido," +
							" t_PAYMENT.*" +
						" FROM t_PAGTO_GW_PAG t_PAG" +
							" INNER JOIN t_PAGTO_GW_PAG_PAYMENT t_PAYMENT ON (t_PAG.id = t_PAYMENT.id_pagto_gw_pag)" +
						" WHERE" +
							" (pedido = '" + numeroPedido + "')" +
							" AND (st_enviado_analise_AF = 0)" +
							" AND (st_cancelado_envio_analise_AF = 0)" +
							" AND (ult_GlobalStatus IN (" +
									"'" + Global.Cte.Braspag.Pagador.GlobalStatus.AUTORIZADA.GetValue() + "'," +
									"'" + Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue() + "'))" +
						" ORDER BY" +
							" t_PAYMENT.id";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);

				return dtbResultado;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				strMsgErro = NOME_DESTA_ROTINA + "\n" + ex.Message;

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = "numeroPedido=" + numeroPedido;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ getPagPaymentRowsParaFinalizacaoByIdPagtoGwAf ]
		public static DataTable getPagPaymentRowsParaFinalizacaoByIdPagtoGwAf(int id_pagto_gw_af, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.getPagPaymentRowsParaFinalizacaoByIdPagtoGwAf()";
			string msg_erro_aux;
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
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
							" t_PAG.data," +
							" t_PAG.data_hora," +
							" t_PAG.usuario," +
							" t_PAG.loja," +
							" t_PAG.id_cliente," +
							" t_PAG.pedido," +
							" t_PAG.pedido_com_sufixo_nsu," +
							" t_PAG.owner," +
							" t_PAG.executado_pelo_cliente_status," +
							" t_PAG.origem_endereco_IP," +
							" t_PAG.FingerPrint_SessionID," +
							" t_PAG.valor_pedido," +
							" t_PAYMENT.*" +
						" FROM t_PAGTO_GW_PAG t_PAG" +
							" INNER JOIN t_PAGTO_GW_PAG_PAYMENT t_PAYMENT ON (t_PAG.id = t_PAYMENT.id_pagto_gw_pag)" +
						" WHERE" +
							" (id_pagto_gw_af = " + id_pagto_gw_af.ToString() + ")" +
							" AND (st_enviado_analise_AF = 1)" +
							" AND (st_processamento_AF_finalizado = 0)" +
							" AND (st_processamento_PAG_finalizado = 0)" +
							" AND (st_cancelado_envio_analise_AF = 0)" +
							" AND (ult_GlobalStatus IN (" +
										"'" + Global.Cte.Braspag.Pagador.GlobalStatus.AUTORIZADA.GetValue() + "'," +
										"'" + Global.Cte.Braspag.Pagador.GlobalStatus.CAPTURADA.GetValue() + "'))" +
						" ORDER BY" +
							" t_PAYMENT.id";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);

				return dtbResultado;
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = "id_pagto_gw_af=" + id_pagto_gw_af.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ getClearsaleAFById ]
		public static ClearsaleAF getClearsaleAFById(int id, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.getClearsaleAFById()";
			string msg_erro_aux;
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
			DataRow row;
			ClearsaleAF clearsaleAF;
			ClearsaleAFPayment payment;
			ClearsaleAFItem item;
			ClearsaleAFPhone phone;
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

				#region [ Obtém dados do registro principal ]
				strSql = "SELECT " +
							"*" +
						" FROM t_PAGTO_GW_AF" +
						" WHERE" +
							" (id = " + id.ToString() + ")";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count == 0) return null;

				row = dtbResultado.Rows[0];
				clearsaleAF = new ClearsaleAF();
				clearsaleAF.id = BD.readToInt(row["id"]);
				clearsaleAF.data = BD.readToDateTime(row["data"]);
				clearsaleAF.data_hora = BD.readToDateTime(row["data_hora"]);
				clearsaleAF.usuario = BD.readToString(row["usuario"]);
				clearsaleAF.loja = BD.readToString(row["loja"]);
				clearsaleAF.id_cliente = BD.readToString(row["id_cliente"]);
				clearsaleAF.pedido = BD.readToString(row["pedido"]);
				clearsaleAF.pedido_com_sufixo_nsu = BD.readToString(row["pedido_com_sufixo_nsu"]);
				clearsaleAF.valor_pedido = BD.readToDecimal(row["valor_pedido"]);
				clearsaleAF.trx_TX_data = BD.readToDateTime(row["trx_TX_data"]);
				clearsaleAF.trx_TX_data_hora = BD.readToDateTime(row["trx_TX_data_hora"]);
				clearsaleAF.trx_RX_status = BD.readToByte(row["trx_RX_status"]);
				clearsaleAF.trx_RX_data = BD.readToDateTime(row["trx_RX_data"]);
				clearsaleAF.trx_RX_data_hora = BD.readToDateTime(row["trx_RX_data_hora"]);
				clearsaleAF.trx_RX_vazio_status = BD.readToByte(row["trx_RX_vazio_status"]);
				clearsaleAF.trx_erro_status = BD.readToByte(row["trx_erro_status"]);
				clearsaleAF.trx_erro_codigo = BD.readToString(row["trx_erro_codigo"]);
				clearsaleAF.trx_erro_mensagem = BD.readToString(row["trx_erro_mensagem"]);
				clearsaleAF.trx_TX_id_pagto_gw_af_xml = BD.readToInt(row["trx_TX_id_pagto_gw_af_xml"]);
				clearsaleAF.trx_RX_id_pagto_gw_af_xml = BD.readToInt(row["trx_RX_id_pagto_gw_af_xml"]);
				clearsaleAF.req_entityCode = BD.readToString(row["req_entityCode"]);
				clearsaleAF.req_Order_ID = BD.readToString(row["req_Order_ID"]);
				clearsaleAF.req_Order_FingerPrint_SessionID = BD.readToString(row["req_Order_FingerPrint_SessionID"]);
				clearsaleAF.req_Order_Date = BD.readToString(row["req_Order_Date"]);
				clearsaleAF.req_Order_Email = BD.readToString(row["req_Order_Email"]);
				clearsaleAF.req_Order_B2B_B2C = BD.readToString(row["req_Order_B2B_B2C"]);
				clearsaleAF.req_Order_ShippingPrice = BD.readToString(row["req_Order_ShippingPrice"]);
				clearsaleAF.req_Order_TotalItems = BD.readToString(row["req_Order_TotalItems"]);
				clearsaleAF.req_Order_TotalOrder = BD.readToString(row["req_Order_TotalOrder"]);
				clearsaleAF.req_Order_QtyInstallments = BD.readToString(row["req_Order_QtyInstallments"]);
				clearsaleAF.req_Order_DeliveryTimeCD = BD.readToString(row["req_Order_DeliveryTimeCD"]);
				clearsaleAF.req_Order_QtyItems = BD.readToString(row["req_Order_QtyItems"]);
				clearsaleAF.req_Order_QtyPaymentTypes = BD.readToString(row["req_Order_QtyPaymentTypes"]);
				clearsaleAF.req_Order_IP = BD.readToString(row["req_Order_IP"]);
				clearsaleAF.req_Order_Status = BD.readToString(row["req_Order_Status"]);
				clearsaleAF.req_Order_Reanalise = BD.readToString(row["req_Order_Reanalise"]);
				clearsaleAF.req_Order_Origin = BD.readToString(row["req_Order_Origin"]);
				clearsaleAF.req_Order_BillingData_ID = BD.readToString(row["req_Order_BillingData_ID"]);
				clearsaleAF.req_Order_BillingData_Type = BD.readToString(row["req_Order_BillingData_Type"]);
				clearsaleAF.req_Order_BillingData_LegalDocument1 = BD.readToString(row["req_Order_BillingData_LegalDocument1"]);
				clearsaleAF.req_Order_BillingData_LegalDocument2 = BD.readToString(row["req_Order_BillingData_LegalDocument2"]);
				clearsaleAF.req_Order_BillingData_Name = BD.readToString(row["req_Order_BillingData_Name"]);
				clearsaleAF.req_Order_BillingData_BirthDate = BD.readToString(row["req_Order_BillingData_BirthDate"]);
				clearsaleAF.req_Order_BillingData_Email = BD.readToString(row["req_Order_BillingData_Email"]);
				clearsaleAF.req_Order_BillingData_Gender = BD.readToString(row["req_Order_BillingData_Gender"]);
				clearsaleAF.req_Order_BillingData_Address_Street = BD.readToString(row["req_Order_BillingData_Address_Street"]);
				clearsaleAF.req_Order_BillingData_Address_Number = BD.readToString(row["req_Order_BillingData_Address_Number"]);
				clearsaleAF.req_Order_BillingData_Address_Comp = BD.readToString(row["req_Order_BillingData_Address_Comp"]);
				clearsaleAF.req_Order_BillingData_Address_County = BD.readToString(row["req_Order_BillingData_Address_County"]);
				clearsaleAF.req_Order_BillingData_Address_City = BD.readToString(row["req_Order_BillingData_Address_City"]);
				clearsaleAF.req_Order_BillingData_Address_State = BD.readToString(row["req_Order_BillingData_Address_State"]);
				clearsaleAF.req_Order_BillingData_Address_Country = BD.readToString(row["req_Order_BillingData_Address_Country"]);
				clearsaleAF.req_Order_BillingData_Address_ZipCode = BD.readToString(row["req_Order_BillingData_Address_ZipCode"]);
				clearsaleAF.req_Order_BillingData_Address_Reference = BD.readToString(row["req_Order_BillingData_Address_Reference"]);
				clearsaleAF.req_Order_ShippingData_ID = BD.readToString(row["req_Order_ShippingData_ID"]);
				clearsaleAF.req_Order_ShippingData_Type = BD.readToString(row["req_Order_ShippingData_Type"]);
				clearsaleAF.req_Order_ShippingData_LegalDocument1 = BD.readToString(row["req_Order_ShippingData_LegalDocument1"]);
				clearsaleAF.req_Order_ShippingData_LegalDocument2 = BD.readToString(row["req_Order_ShippingData_LegalDocument2"]);
				clearsaleAF.req_Order_ShippingData_Name = BD.readToString(row["req_Order_ShippingData_Name"]);
				clearsaleAF.req_Order_ShippingData_BirthDate = BD.readToString(row["req_Order_ShippingData_BirthDate"]);
				clearsaleAF.req_Order_ShippingData_Email = BD.readToString(row["req_Order_ShippingData_Email"]);
				clearsaleAF.req_Order_ShippingData_Gender = BD.readToString(row["req_Order_ShippingData_Gender"]);
				clearsaleAF.req_Order_ShippingData_Address_Street = BD.readToString(row["req_Order_ShippingData_Address_Street"]);
				clearsaleAF.req_Order_ShippingData_Address_Number = BD.readToString(row["req_Order_ShippingData_Address_Number"]);
				clearsaleAF.req_Order_ShippingData_Address_Comp = BD.readToString(row["req_Order_ShippingData_Address_Comp"]);
				clearsaleAF.req_Order_ShippingData_Address_County = BD.readToString(row["req_Order_ShippingData_Address_County"]);
				clearsaleAF.req_Order_ShippingData_Address_City = BD.readToString(row["req_Order_ShippingData_Address_City"]);
				clearsaleAF.req_Order_ShippingData_Address_State = BD.readToString(row["req_Order_ShippingData_Address_State"]);
				clearsaleAF.req_Order_ShippingData_Address_Country = BD.readToString(row["req_Order_ShippingData_Address_Country"]);
				clearsaleAF.req_Order_ShippingData_Address_ZipCode = BD.readToString(row["req_Order_ShippingData_Address_ZipCode"]);
				clearsaleAF.req_Order_ShippingData_Address_Reference = BD.readToString(row["req_Order_ShippingData_Address_Reference"]);
				clearsaleAF.resp_ID = BD.readToString(row["resp_ID"]);
				clearsaleAF.resp_Status = BD.readToString(row["resp_Status"]);
				clearsaleAF.resp_Score = BD.readToString(row["resp_Score"]);
				clearsaleAF.owner = BD.readToInt(row["owner"]);
				clearsaleAF.prim_Status = BD.readToString(row["prim_Status"]);
				clearsaleAF.prim_atualizacao_data_hora = BD.readToDateTime(row["prim_atualizacao_data_hora"]);
				clearsaleAF.ult_Status = BD.readToString(row["ult_Status"]);
				clearsaleAF.ult_atualizacao_data_hora = BD.readToDateTime(row["ult_atualizacao_data_hora"]);
				clearsaleAF.anulado_status = BD.readToByte(row["anulado_status"]);
				clearsaleAF.anulado_data = BD.readToDateTime(row["anulado_data"]);
				clearsaleAF.anulado_data_hora = BD.readToDateTime(row["anulado_data_hora"]);
				clearsaleAF.anulado_por_id_pagto_gw_af = BD.readToInt(row["anulado_por_id_pagto_gw_af"]);
				clearsaleAF.SetOrderAsReturned_pendente_status = BD.readToByte(row["SetOrderAsReturned_pendente_status"]);
				clearsaleAF.SetOrderAsReturned_pendente_data_hora = BD.readToDateTime(row["SetOrderAsReturned_pendente_data_hora"]);
				clearsaleAF.SetOrderAsReturned_sucesso_status = BD.readToByte(row["SetOrderAsReturned_sucesso_status"]);
				clearsaleAF.SetOrderAsReturned_sucesso_data_hora = BD.readToDateTime(row["SetOrderAsReturned_sucesso_data_hora"]);
				#endregion

				#region [ Obtém dados de 'Payment' ]
				strSql = "SELECT " +
							"*" +
						" FROM t_PAGTO_GW_AF_PAYMENT" +
						" WHERE" +
							" (id_pagto_gw_af = " + id.ToString() + ")" +
						" ORDER BY" +
							" ordem";
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daAdapter.Fill(dtbResultado);
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					row = dtbResultado.Rows[i];
					payment = new ClearsaleAFPayment();
					payment.id = BD.readToInt(row["id"]);
					payment.id_pagto_gw_af = BD.readToInt(row["id_pagto_gw_af"]);
					payment.id_pagto_gw_pag_payment = BD.readToInt(row["id_pagto_gw_pag_payment"]);
					payment.ordem = BD.readToInt(row["ordem"]);
					payment.bandeira = BD.readToString(row["bandeira"]);
					payment.valor_transacao = BD.readToDecimal(row["valor_transacao"]);
					payment.af_Sequential = BD.readToString(row["af_Sequential"]);
					payment.af_Date = BD.readToString(row["af_Date"]);
					payment.af_Amount = BD.readToString(row["af_Amount"]);
					payment.af_PaymentTypeID = BD.readToString(row["af_PaymentTypeID"]);
					payment.af_QtyInstallments = BD.readToString(row["af_QtyInstallments"]);
					payment.af_Interest = BD.readToString(row["af_Interest"]);
					payment.af_InterestValue = BD.readToString(row["af_InterestValue"]);
					payment.af_CardNumber = BD.readToString(row["af_CardNumber"]);
					payment.af_CardBin = BD.readToString(row["af_CardBin"]);
					payment.af_CardEndNumber = BD.readToString(row["af_CardEndNumber"]);
					payment.af_CardType = BD.readToString(row["af_CardType"]);
					payment.af_CardExpirationDate = BD.readToString(row["af_CardExpirationDate"]);
					payment.af_Name = BD.readToString(row["af_Name"]);
					payment.af_LegalDocument = BD.readToString(row["af_LegalDocument"]);
					payment.af_Address_Street = BD.readToString(row["af_Address_Street"]);
					payment.af_Address_Number = BD.readToString(row["af_Address_Number"]);
					payment.af_Address_Comp = BD.readToString(row["af_Address_Comp"]);
					payment.af_Address_County = BD.readToString(row["af_Address_County"]);
					payment.af_Address_City = BD.readToString(row["af_Address_City"]);
					payment.af_Address_State = BD.readToString(row["af_Address_State"]);
					payment.af_Address_Country = BD.readToString(row["af_Address_Country"]);
					payment.af_Address_ZipCode = BD.readToString(row["af_Address_ZipCode"]);
					payment.af_Address_Reference = BD.readToString(row["af_Address_Reference"]);
					payment.af_Nsu = BD.readToString(row["af_Nsu"]);
					payment.af_Currency = BD.readToString(row["af_Currency"]);
					clearsaleAF.Payments.Add(payment);
				}
				#endregion

				#region [ Obtém dados de 'Item' ]
				strSql = "SELECT " +
							"*" +
						" FROM t_PAGTO_GW_AF_ITEM" +
						" WHERE" +
							" (id_pagto_gw_af = " + id.ToString() + ")" +
						" ORDER BY" +
							" id";
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daAdapter.Fill(dtbResultado);
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					row = dtbResultado.Rows[i];
					item = new ClearsaleAFItem();
					item.id = BD.readToInt(row["id"]);
					item.id_pagto_gw_af = BD.readToInt(row["id_pagto_gw_af"]);
					item.af_ID = BD.readToString(row["af_ID"]);
					item.af_Name = BD.readToString(row["af_Name"]);
					item.af_ItemValue = BD.readToString(row["af_ItemValue"]);
					item.af_Qty = BD.readToString(row["af_Qty"]);
					item.af_CategoryID = BD.readToString(row["af_CategoryID"]);
					item.af_CategoryName = BD.readToString(row["af_CategoryName"]);
					clearsaleAF.Items.Add(item);
				}
				#endregion

				#region [ Obtém dados de 'Order/BillingData/Phones' ]
				strSql = "SELECT " +
							"*" +
						" FROM t_PAGTO_GW_AF_PHONE" +
						" WHERE" +
							" (id_pagto_gw_af = " + id.ToString() + ")" +
							" AND (IdBlocoXml = '" + Global.Cte.Clearsale.T_PAGTO_GW_AF_PHONE_IdBlocoXml.Order_BillingData_Phones.GetValue() + "')" +
						" ORDER BY" +
							" id";
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daAdapter.Fill(dtbResultado);
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					row = dtbResultado.Rows[i];
					phone = new ClearsaleAFPhone();
					phone.id = BD.readToInt(row["id"]);
					phone.id_pagto_gw_af = BD.readToInt(row["id_pagto_gw_af"]);
					phone.idBlocoXml = BD.readToString(row["IdBlocoXml"]);
					phone.af_Type = BD.readToString(row["af_Type"]);
					phone.af_DDI = BD.readToString(row["af_DDI"]);
					phone.af_DDD = BD.readToString(row["af_DDD"]);
					phone.af_Number = BD.readToString(row["af_Number"]);
					phone.af_Extension = BD.readToString(row["af_Extension"]);
					clearsaleAF.Order_BillingData_Phones.Add(phone);
				}
				#endregion

				#region [ Obtém dados de 'Order/ShippingData/Phones' ]
				strSql = "SELECT " +
							"*" +
						" FROM t_PAGTO_GW_AF_PHONE" +
						" WHERE" +
							" (id_pagto_gw_af = " + id.ToString() + ")" +
							" AND (IdBlocoXml = '" + Global.Cte.Clearsale.T_PAGTO_GW_AF_PHONE_IdBlocoXml.Order_ShippingData_Phones.GetValue() + "')" +
						" ORDER BY" +
							" id";
				cmCommand.CommandText = strSql;
				dtbResultado.Reset();
				daAdapter.Fill(dtbResultado);
				for (int i = 0; i < dtbResultado.Rows.Count; i++)
				{
					row = dtbResultado.Rows[i];
					phone = new ClearsaleAFPhone();
					phone.id = BD.readToInt(row["id"]);
					phone.id_pagto_gw_af = BD.readToInt(row["id_pagto_gw_af"]);
					phone.idBlocoXml = BD.readToString(row["IdBlocoXml"]);
					phone.af_Type = BD.readToString(row["af_Type"]);
					phone.af_DDI = BD.readToString(row["af_DDI"]);
					phone.af_DDD = BD.readToString(row["af_DDD"]);
					phone.af_Number = BD.readToString(row["af_Number"]);
					phone.af_Extension = BD.readToString(row["af_Extension"]);
					clearsaleAF.Order_ShippingData_Phones.Add(phone);
				}
				#endregion

				return clearsaleAF;
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
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ getClearsaleAFByOrderID ]
		public static ClearsaleAF getClearsaleAFByOrderID(string orderID, string entityCode, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.getClearsaleAFByOrderID()";
			int id;
			string msg_erro_aux;
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
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

				strSql = "SELECT TOP 1 " +
							"id" +
						" FROM t_PAGTO_GW_AF" +
						" WHERE" +
							" (req_Order_ID = '" + orderID + "')" +
							" AND (req_entityCode = '" + entityCode + "')" +
							" AND (anulado_status = 0)" +
						" ORDER BY" +
							" id DESC";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count == 0) return null;

				id = BD.readToInt(dtbResultado.Rows[0]["id"]);
				return getClearsaleAFById(id, out msg_erro);
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = "orderID=" + orderID;
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ getClearsaleAFByIdPagtoGwPagPayment ]
		public static ClearsaleAF getClearsaleAFByIdPagtoGwPagPayment(int id_pagto_gw_pag_payment, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.getClearsaleAFByIdPagtoGwPagPayment()";
			int id;
			string msg_erro_aux;
			String strSql;
			SqlCommand cmCommand;
			SqlDataAdapter daAdapter;
			DataTable dtbResultado = new DataTable();
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

				strSql = "SELECT TOP 1 " +
							"t_PAGTO_GW_AF.id AS id" +
						" FROM t_PAGTO_GW_AF" +
							" INNER JOIN t_PAGTO_GW_AF_PAYMENT ON (t_PAGTO_GW_AF.id=t_PAGTO_GW_AF_PAYMENT.id_pagto_gw_af)" +
						" WHERE" +
							" (t_PAGTO_GW_AF_PAYMENT.id_pagto_gw_pag_payment = " + id_pagto_gw_pag_payment.ToString() + ")" +
							" AND (t_PAGTO_GW_AF.anulado_status = 0)" +
						" ORDER BY" +
							" t_PAGTO_GW_AF.id DESC";
				cmCommand.CommandText = strSql;
				daAdapter.Fill(dtbResultado);
				if (dtbResultado.Rows.Count == 0) return null;

				id = BD.readToInt(dtbResultado.Rows[0]["id"]);
				return getClearsaleAFById(id, out msg_erro);
			}
			catch (Exception ex)
			{
				// Retorna mensagem de erro p/ rotina chamadora
				msg_erro = NOME_DESTA_ROTINA + "\n" + ex.ToString();

				#region [ Registra detalhes em t_FINSVC_LOG (chama gravaLogAtividade() automaticamente) ]
				FinSvcLog svcLog = new FinSvcLog();
				svcLog.operacao = NOME_DESTA_ROTINA;
				svcLog.descricao = ex.ToString();
				svcLog.complemento_1 = "id_pagto_gw_pag_payment=" + id_pagto_gw_pag_payment.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return null;
			}
		}
		#endregion

		#region [ isPedidoERPDesteAmbiente ]
		/// <summary>
		/// Analisa se o número do pedido está no formato usado no sistema (ERP)
		/// Formato usado no sistema: 999999X ou 999999X-A
		/// </summary>
		/// <param name="orderID">Número do pedido a ser analisado</param>
		/// <param name="entityCode">Chave EntityCode usada nas requisições Clearsale</param>
		/// <returns>
		/// true = pedido no formato usado no sistema (ERP)
		/// false = pedido fora do padrão do sistema (ERP)
		/// </returns>
		public static bool isPedidoERPDesteAmbiente(string orderID, string entityCode)
		{
			#region [ Declarações ]
			bool blnTemLetra = false;
			int qtdeDigitosPrefixo = 0;
			string strSql;
			string strPedido = "";
			SqlCommand cmCommand;
			SqlDataReader dr;
			#endregion

			// Importante: é necessário verificar se o pedido está cadastrado no ambiente ao qual o serviço está conectado (DIS ou OLD01). Além disso, é importante
			// =========== lembrar que o nº do pedido enviado p/ a Clearsale pode conter um sufixo, caso tenham sido enviados mais do que uma requisição de análise AF. Esta situação
			// pode ocorrer no caso de pagamento com múltiplos cartões em que uma das transações foi negada pelo Pagador e o cliente não fez outra em substituição dentro do tempo
			// limite que o serviço aguarda p/ integralizar o pagamento antes de enviar a requisição AF p/ a Clearsale.
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
					" FROM t_PAGTO_GW_AF" +
					" WHERE" +
						" (req_Order_ID = '" + orderID + "')" +
						" AND (req_entityCode = '" + entityCode + "')" +
						" AND (anulado_status = 0)" +
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

		#region [ updateRegistroAFSetOrderAsReturnedPendente ]
		public static bool updateRegistroAFSetOrderAsReturnedPendente(int id, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.updateRegistroAFSetOrderAsReturnedPendente()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateAFSetOrderAsReturnedPendente.Parameters["@id"].Value = id;
				#endregion

				#region [ Tenta alterar o(s) registro(s), se houver algum ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateAFSetOrderAsReturnedPendente);
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
					svcLog.complemento_1 = "id=" + id.ToString();
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
				svcLog.complemento_1 = "id=" + id.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion

		#region [ updateRegistroAFSetOrderAsReturnedSucesso ]
		public static bool updateRegistroAFSetOrderAsReturnedSucesso(int id, out string msg_erro)
		{
			#region [ Declarações ]
			const string NOME_DESTA_ROTINA = "ClearsaleDAO.updateRegistroAFSetOrderAsReturnedSucesso()";
			int intRetorno;
			string msg_erro_aux;
			#endregion

			msg_erro = "";
			try
			{
				#region [ Preenche o valor dos parâmetros ]
				cmUpdateAFSetOrderAsReturnedSucesso.Parameters["@id"].Value = id;
				#endregion

				#region [ Tenta alterar o(s) registro(s), se houver algum ]
				try
				{
					intRetorno = BD.executaNonQuery(ref cmUpdateAFSetOrderAsReturnedSucesso);
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
					svcLog.complemento_1 = "id=" + id.ToString();
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
				svcLog.complemento_1 = "id=" + id.ToString();
				GeralDAO.gravaFinSvcLog(svcLog, out msg_erro_aux);
				#endregion

				return false;
			}
		}
		#endregion
	}
}
